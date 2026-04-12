using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.UI.Xaml.Media.Imaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using A_Simple_Recorder.Services;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Pickers;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace A_Simple_Recorder
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        private readonly CameraService _cameraService;
        private readonly ScreenRecordingService _screenRecordingService;
        private readonly VideoCallService _videoCallService;
        private readonly MicrophoneService _microphoneService;
        private bool _isRecording = false;
        private bool _isHosting = false; // Track if we're in hosting mode (preview only)
        private bool _isMuted = false;
        private StorageFolder? _saveFolder;
        private System.Timers.Timer? _voiceActivityTimer;
        private Dictionary<string, Border> _participantThumbnails = new();
        private string? _currentSpeakerId;
        private volatile bool _isShuttingDown = false;

        public MainWindow()
        {
            InitializeComponent();
            _cameraService = new CameraService();
            _screenRecordingService = new ScreenRecordingService();
            _videoCallService = new VideoCallService();
            _microphoneService = new MicrophoneService();

            this.Activated += MainWindow_Activated;
            this.Closed += MainWindow_Closed;

            // Load saved save location or use default
            LoadSaveLocation();
            UpdateSaveLocationText();

            // Setup video call events
            _videoCallService.OnVideoFrameReceived += OnRemoteVideoFrameReceived;
            _videoCallService.OnAudioChunkReceived += OnRemoteAudioChunkReceived;
            _videoCallService.OnParticipantJoined += OnParticipantJoined;
            _videoCallService.OnParticipantLeft += OnParticipantLeft;
            _videoCallService.OnVoiceActivity += OnRemoteVoiceActivity;

            // Display your IP
            YourIpText.Text = $"Your IP: {_videoCallService.GetLocalIPAddress()}";

            // Setup voice activity detection timer
            _voiceActivityTimer = new System.Timers.Timer(100); // Check every 100ms
            _voiceActivityTimer.Elapsed += CheckVoiceActivity;
            _voiceActivityTimer.Start();
        }

        private void UpdateSaveLocationText()
        {
            if (_saveFolder != null)
            {
                SaveLocationText.Text = $"Location: {_saveFolder.Path}";
            }
        }

        private async void LoadSaveLocation()
        {
            try
            {
                var localSettings = ApplicationData.Current.LocalSettings;
                if (localSettings.Values.TryGetValue("SaveFolderToken", out var token))
                {
                    var folder = await Windows.Storage.AccessCache.StorageApplicationPermissions
                        .FutureAccessList.GetFolderAsync(token.ToString());
                    if (folder != null)
                    {
                        _saveFolder = folder;
                        System.Diagnostics.Debug.WriteLine($"Loaded save location: {folder.Path}");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading save location: {ex.Message}");
            }

            // Default to LocalFolder if no saved location
            _saveFolder = ApplicationData.Current.LocalFolder;
        }

        private void SaveSaveLocation()
        {
            try
            {
                if (_saveFolder != null)
                {
                    var localSettings = ApplicationData.Current.LocalSettings;
                    Windows.Storage.AccessCache.StorageApplicationPermissions
                        .FutureAccessList.AddOrReplace("SaveFolder", _saveFolder);
                    localSettings.Values["SaveFolderToken"] = "SaveFolder";
                    System.Diagnostics.Debug.WriteLine("Saved location token: SaveFolder");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving location: {ex.Message}");
            }
        }

        private async void MainWindow_Activated(object sender, WindowActivatedEventArgs args)
        {
            if (args.WindowActivationState != WindowActivationState.Deactivated)
            {
                this.Activated -= MainWindow_Activated; // Only run once
                await RefreshCamerasAsync();
                await RefreshMicrophonesAsync();

                // Start audio monitoring for voice activity detection
                await _microphoneService.StartMonitoringAsync();
            }
        }

        private void MainWindow_Closed(object sender, WindowEventArgs args)
        {
            // Signal shutdown first to prevent new UI updates
            _isShuttingDown = true;

            // Stop screen recording service first
            if (_screenRecordingService != null)
            {
                _screenRecordingService.SignalShutdown();
            }

            // Stop timer to prevent new callbacks from starting
            if (_voiceActivityTimer != null)
            {
                _voiceActivityTimer.Stop();
                _voiceActivityTimer.Dispose();
                _voiceActivityTimer = null;
            }

            // Wait longer for any pending timer/capture callbacks to complete
            System.Threading.Thread.Sleep(300);

            // Stop monitoring after timer is fully stopped
            _microphoneService?.StopMonitoring();
        }

        private async void ChooseLocationButton_Click(object sender, RoutedEventArgs e)
        {
            var picker = new FolderPicker();
            picker.SuggestedStartLocation = PickerLocationId.VideosLibrary;
            picker.FileTypeFilter.Add("*");

            // Get the window handle
            var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

            var folder = await picker.PickSingleFolderAsync();
            if (folder != null)
            {
                _saveFolder = folder;
                SaveSaveLocation();
                UpdateSaveLocationText();
                StatusText.Text = "Save location updated";
            }
        }

        private string GetFileName()
        {
            var baseName = string.IsNullOrWhiteSpace(FilenameTextBox.Text) ? "Recording" : FilenameTextBox.Text;

            if (AutoDateTimeCheckBox.IsChecked == true)
            {
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                return $"{baseName}_{timestamp}";
            }

            return baseName;
        }

        private async void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            await RefreshCamerasAsync();
        }

        private async Task RefreshCamerasAsync()
        {
            StatusText.Text = "Detecting cameras...";

            var cameras = await _cameraService.GetAvailableCamerasAsync();

            // Clear and rebuild the list
            SourceComboBox.Items.Clear();

            // Add screen recording option
            var screenItem = new ComboBoxItem
            {
                Content = "Screen Recording",
                Tag = "Screen"
            };
            SourceComboBox.Items.Add(screenItem);

            // Add detected cameras
            foreach (var camera in cameras)
            {
                var item = new ComboBoxItem
                {
                    Content = camera.Name,
                    Tag = camera
                };
                SourceComboBox.Items.Add(item);
            }

            SourceComboBox.SelectedIndex = 0;
            StatusText.Text = $"Found {cameras.Count} camera(s)";
        }

        private async Task RefreshMicrophonesAsync()
        {
            var microphones = await _microphoneService.GetAvailableMicrophonesAsync();

            MicrophoneComboBox.Items.Clear();

            foreach (var mic in microphones)
            {
                var item = new ComboBoxItem
                {
                    Content = mic.Name,
                    Tag = mic
                };
                MicrophoneComboBox.Items.Add(item);
            }

            if (microphones.Count > 0)
            {
                MicrophoneComboBox.SelectedIndex = 0;
            }

            System.Diagnostics.Debug.WriteLine($"Found {microphones.Count} microphone(s)");
        }

        private void SourceComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StatusText == null) return; // Not fully initialized yet

            if (SourceComboBox.SelectedItem is ComboBoxItem item)
            {
                if (item.Tag is string tag && tag == "Screen")
                {
                    StatusText.Text = "Screen recording selected";
                }
                else if (item.Tag is CameraInfo camera)
                {
                    StatusText.Text = $"Camera selected: {camera.Name}";
                }
            }
        }

        private async void RecordButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isRecording)
            {
                // If hosting, start file recording (preview already running)
                if (_isHosting && _screenRecordingService.IsCapturing)
                {
                    var filename = GetFileName() + ".mp4";
                    var micDeviceId = GetSelectedMicrophoneDeviceId();
                    await _screenRecordingService.StartFileRecordingAsync(_saveFolder, filename, micDeviceId);
                    _isRecording = true;
                    RecordButton.Content = "Stop";
                    StatusText.Text = "Recording to file...";
                }
                else
                {
                    // Normal recording (start both preview and file recording)
                    await StartRecordingAsync();
                }
            }
            else
            {
                // Stop file recording
                if (_isHosting)
                {
                    // Just stop file recording, keep preview running
                    var savedPath = await _screenRecordingService.StopFileRecordingAsync();
                    _isRecording = false;
                    RecordButton.Content = "Record";
                    StatusText.Text = $"Recording saved to: {savedPath}";
                }
                else
                {
                    // Stop everything
                    await StopRecordingAsync();
                }
            }
        }

        private async Task StartRecordingAsync()
        {
            try
            {
                if (SourceComboBox.SelectedItem is ComboBoxItem item)
                {
                    if (item.Tag is string tag && tag == "Screen")
                    {
                        // Get filename (will be converted to .avi by service)
                        var filename = GetFileName() + ".mp4";

                        // Get selected microphone
                        var micDeviceId = GetSelectedMicrophoneDeviceId();

                        // Start screen recording with preview
                        await _screenRecordingService.StartRecordingAsync(this, _saveFolder, filename, OnFrameCaptured, micDeviceId);

                        // Hide "Wait" text (image is always visible now)
                        WaitText.Visibility = Visibility.Collapsed;

                        StatusText.Text = $"Recording screen (will be saved as .avi)...";
                    }
                    else if (item.Tag is CameraInfo camera)
                    {
                        // Get filename
                        var filename = GetFileName() + ".mp4";

                        // Start camera recording
                        await _cameraService.StartRecordingAsync(camera, _saveFolder, filename);
                        StatusText.Text = $"Recording from {camera.Name} to {filename}...";
                    }
                }

                _isRecording = true;
                RecordButton.Content = "Stop";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Error: {ex.Message}";
                _isRecording = false;
            }
        }

        private async void OnFrameCaptured(SoftwareBitmap bitmap)
        {
            try
            {
                if (bitmap == null) return;

                // Convert format if needed
                SoftwareBitmap convertedBitmap = bitmap;
                if (bitmap.BitmapPixelFormat != BitmapPixelFormat.Bgra8 || 
                    bitmap.BitmapAlphaMode != BitmapAlphaMode.Premultiplied)
                {
                    convertedBitmap = SoftwareBitmap.Convert(bitmap, BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
                    bitmap.Dispose();
                }

                // Create and set the bitmap source for local preview
                var source = new SoftwareBitmapSource();
                await source.SetBitmapAsync(convertedBitmap);

                // Update the image
                PreviewImage.Source = source;

                // If in video call, send frame to other participants
                if (_videoCallService.IsConnected)
                {
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            // Encode frame as JPEG
                            using var stream = new MemoryStream();
                            var encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.JpegEncoderId, stream.AsRandomAccessStream());
                            encoder.SetSoftwareBitmap(convertedBitmap);
                            encoder.BitmapTransform.ScaledWidth = 640; // Reduce size for network
                            encoder.BitmapTransform.ScaledHeight = 480;
                            await encoder.FlushAsync();

                            var frameData = stream.ToArray();
                            await _videoCallService.SendVideoFrameAsync(frameData);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error sending video frame: {ex.Message}");
                        }
                    });
                }

                // Dispose the bitmap after setting
                convertedBitmap?.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating preview: {ex.Message}");
            }
        }

        private async Task StopRecordingAsync()
        {
            try
            {
                if (SourceComboBox.SelectedItem is ComboBoxItem item)
                {
                    if (item.Tag is string tag && tag == "Screen")
                    {
                        var savedPath = await _screenRecordingService.StopRecordingAsync();

                        // Show "Wait" text again, clear preview
                        PreviewImage.Source = null;
                        WaitText.Visibility = Visibility.Visible;

                        StatusText.Text = $"Recording saved to: {savedPath}";
                    }
                    else if (item.Tag is CameraInfo)
                    {
                        await _cameraService.StopRecordingAsync();
                        StatusText.Text = "Recording stopped and saved.";
                    }
                }

                _isRecording = false;
                RecordButton.Content = "Record";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Error stopping: {ex.Message}";
                _isRecording = false;
            }
        }

        // Video Call Methods
        private async void HostButton_Click(object sender, RoutedEventArgs e)
        {
            var success = await _videoCallService.StartHostingAsync();
            if (success)
            {
                HostButton.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                DisconnectButton.IsEnabled = true;
                StatusText.Text = $"Hosting on {_videoCallService.GetLocalIPAddress()}:9876";

                // Auto-start PREVIEW only (no recording to file)
                if (!_screenRecordingService.IsCapturing)
                {
                    // Set screen recording as source
                    foreach (var item in SourceComboBox.Items)
                    {
                        if (item is ComboBoxItem comboItem && comboItem.Tag is string tag && tag == "Screen")
                        {
                            SourceComboBox.SelectedItem = item;
                            break;
                        }
                    }

                    // Start preview only (not recording)
                    await _screenRecordingService.StartPreviewAsync(this, OnFrameCaptured);

                    if (_screenRecordingService.IsCapturing)
                    {
                        _isHosting = true;
                        WaitText.Visibility = Visibility.Collapsed;
                        StatusText.Text = $"Hosting on {_videoCallService.GetLocalIPAddress()}:9876 - Streaming";
                    }
                }
            }
            else
            {
                StatusText.Text = "Failed to start hosting";
            }
        }

        private async void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            var ipAddress = IpAddressTextBox.Text.Trim();
            if (string.IsNullOrEmpty(ipAddress))
            {
                StatusText.Text = "Please enter an IP address";
                return;
            }

            var success = await _videoCallService.ConnectToHostAsync(ipAddress);
            if (success)
            {
                HostButton.IsEnabled = false;
                ConnectButton.IsEnabled = false;
                DisconnectButton.IsEnabled = true;
                StatusText.Text = $"Connected to {ipAddress}";
            }
            else
            {
                StatusText.Text = $"Failed to connect to {ipAddress}";
            }
        }

        private async void DisconnectButton_Click(object sender, RoutedEventArgs e)
        {
            await _videoCallService.DisconnectAsync();
            HostButton.IsEnabled = true;
            ConnectButton.IsEnabled = true;
            DisconnectButton.IsEnabled = false;
            StatusText.Text = "Disconnected";

            // Stop recording if active
            if (_isRecording)
            {
                await _screenRecordingService.StopFileRecordingAsync();
                _isRecording = false;
                RecordButton.Content = "Record";
            }

            // Stop preview if hosting
            if (_isHosting)
            {
                await _screenRecordingService.StopPreviewAsync();
                _isHosting = false;
                PreviewImage.Source = null;
                WaitText.Visibility = Visibility.Visible;
            }

            // Clear thumbnails
            ThumbnailsPanel.Children.Clear();
            _participantThumbnails.Clear();
            ParticipantsText.Text = "Participants: 0";
        }

        private async void MicrophoneComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // TODO: Switch microphone device
        }

        private void MuteToggleButton_Click(object sender, RoutedEventArgs e)
        {
            _isMuted = MuteToggleButton.IsChecked == true;

            if (_isMuted)
            {
                MicIcon.Glyph = "\uE74F"; // Muted icon
                MuteButtonText.Text = "Unmute";
                VoiceActivityIndicator.Fill = new SolidColorBrush(Microsoft.UI.Colors.Gray);
                VoiceActivityText.Text = "Muted";
            }
            else
            {
                MicIcon.Glyph = "\uE720"; // Mic icon
                MuteButtonText.Text = "Mute";
                VoiceActivityText.Text = "Not speaking";
            }
        }

        private void CheckVoiceActivity(object? sender, System.Timers.ElapsedEventArgs e)
        {
            if (_isMuted || _isShuttingDown) return;

            // Get actual audio level from microphone service
            var audioLevel = _microphoneService.GetCurrentAudioLevel();
            var isSpeaking = audioLevel > 0.01f; // Threshold for speaking detection

            // Check if DispatcherQueue is still available and not shutting down
            if (DispatcherQueue != null && !_isShuttingDown)
            {
                DispatcherQueue.TryEnqueue(() =>
                {
                    // Double-check shutdown flag inside UI thread
                    if (_isShuttingDown) return;

                    if (isSpeaking)
                    {
                        VoiceActivityIndicator.Fill = new SolidColorBrush(Microsoft.UI.Colors.LimeGreen);
                        VoiceActivityText.Text = "Speaking";
                    }
                    else
                    {
                        VoiceActivityIndicator.Fill = new SolidColorBrush(Microsoft.UI.Colors.Gray);
                        VoiceActivityText.Text = "Not speaking";
                    }
                });
            }

            // Send voice activity to other participants only if not shutting down
            if (!_isShuttingDown)
            {
                _ = _videoCallService.SendVoiceActivityAsync(isSpeaking, audioLevel);
            }
        }

        private void OnRemoteVideoFrameReceived(string senderId, byte[] frameData)
        {
            if (DispatcherQueue == null || _isShuttingDown) return;

            DispatcherQueue.TryEnqueue(async () =>
            {
                try
                {
                    // Decode JPEG frame
                    using var stream = new MemoryStream(frameData);
                    var decoder = await BitmapDecoder.CreateAsync(stream.AsRandomAccessStream());
                    var bitmap = await decoder.GetSoftwareBitmapAsync();

                    // Convert to displayable format
                    if (bitmap.BitmapPixelFormat != BitmapPixelFormat.Bgra8 ||
                        bitmap.BitmapAlphaMode != BitmapAlphaMode.Premultiplied)
                    {
                        bitmap = SoftwareBitmap.Convert(bitmap, BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
                    }

                    var source = new SoftwareBitmapSource();
                    await source.SetBitmapAsync(bitmap);

                    // Update thumbnail or main display based on current speaker
                    if (senderId == _currentSpeakerId || _currentSpeakerId == null)
                    {
                        // Show in main display
                        PreviewImage.Source = source;
                        WaitText.Visibility = Visibility.Collapsed;

                        var participant = _videoCallService.Participants.FirstOrDefault(p => p.Id == senderId);
                        if (participant != null)
                        {
                            SpeakerNameText.Text = participant.Name;
                            SpeakerNameBorder.Visibility = Visibility.Visible;
                        }
                    }
                    else
                    {
                        // Update thumbnail
                        UpdateParticipantThumbnail(senderId, source);
                    }

                    bitmap.Dispose();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error displaying remote frame: {ex.Message}");
                }
            });
        }

        private void OnRemoteAudioChunkReceived(string senderId, byte[] audioData)
        {
            // TODO: Play audio from remote participant
        }

        private void OnParticipantJoined(ParticipantInfo participant)
        {
            if (DispatcherQueue == null || _isShuttingDown) return;

            DispatcherQueue.TryEnqueue(() =>
            {
                // Add thumbnail for participant
                var thumbnail = CreateParticipantThumbnail(participant);
                ThumbnailsPanel.Children.Add(thumbnail);
                _participantThumbnails[participant.Id] = thumbnail;

                ParticipantsText.Text = $"Participants: {_participantThumbnails.Count + 1}";
                StatusText.Text = $"{participant.Name} joined";
            });
        }

        private void OnParticipantLeft(string participantId)
        {
            if (DispatcherQueue == null || _isShuttingDown) return;

            DispatcherQueue.TryEnqueue(() =>
            {
                if (_participantThumbnails.TryGetValue(participantId, out var thumbnail))
                {
                    ThumbnailsPanel.Children.Remove(thumbnail);
                    _participantThumbnails.Remove(participantId);
                }

                ParticipantsText.Text = $"Participants: {_participantThumbnails.Count + 1}";
                StatusText.Text = "Participant left";

                if (_currentSpeakerId == participantId)
                {
                    _currentSpeakerId = null;
                    SpeakerNameBorder.Visibility = Visibility.Collapsed;
                }
            });
        }

        private void OnRemoteVoiceActivity(string participantId, bool isSpeaking, float audioLevel)
        {
            if (isSpeaking && participantId != _currentSpeakerId)
            {
                _currentSpeakerId = participantId;

                if (DispatcherQueue == null || _isShuttingDown) return;

                DispatcherQueue.TryEnqueue(() =>
                {
                    var participant = _videoCallService.Participants.FirstOrDefault(p => p.Id == participantId);
                    if (participant != null)
                    {
                        SpeakerNameText.Text = participant.Name;
                        SpeakerNameBorder.Visibility = Visibility.Visible;
                    }

                    // Highlight speaking thumbnail
                    if (_participantThumbnails.TryGetValue(participantId, out var thumbnail))
                    {
                        thumbnail.BorderBrush = new SolidColorBrush(Microsoft.UI.Colors.LimeGreen);
                        thumbnail.BorderThickness = new Thickness(3);
                    }
                });
            }
            else if (!isSpeaking && participantId == _currentSpeakerId)
            {
                if (DispatcherQueue == null || _isShuttingDown) return;

                DispatcherQueue.TryEnqueue(() =>
                {
                    if (_participantThumbnails.TryGetValue(participantId, out var thumbnail))
                    {
                        thumbnail.BorderBrush = new SolidColorBrush(Microsoft.UI.Colors.Gray);
                        thumbnail.BorderThickness = new Thickness(2);
                    }
                });
            }
        }

        private Border CreateParticipantThumbnail(ParticipantInfo participant)
        {
            var image = new Image
            {
                Width = 150,
                Height = 100,
                Stretch = Stretch.UniformToFill
            };

            var nameText = new TextBlock
            {
                Text = participant.Name,
                Foreground = new SolidColorBrush(Microsoft.UI.Colors.White),
                FontSize = 12,
                Margin = new Thickness(5),
                HorizontalAlignment = HorizontalAlignment.Center
            };

            var grid = new Grid();
            grid.Children.Add(image);
            grid.Children.Add(new Border
            {
                Background = new SolidColorBrush(Microsoft.UI.ColorHelper.FromArgb(170, 0, 0, 0)),
                VerticalAlignment = VerticalAlignment.Bottom,
                Child = nameText
            });

            var border = new Border
            {
                Width = 150,
                Height = 100,
                Background = new SolidColorBrush(Microsoft.UI.Colors.Black),
                BorderBrush = new SolidColorBrush(Microsoft.UI.Colors.Gray),
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(4),
                Child = grid
            };

            border.Tag = new { ParticipantId = participant.Id, Image = image };

            return border;
        }

        private string? GetSelectedMicrophoneDeviceId()
        {
            if (MicrophoneComboBox.SelectedItem is ComboBoxItem item && item.Tag is MicrophoneInfo mic)
            {
                return mic.Id;
            }
            return null;
        }

        private void UpdateParticipantThumbnail(string participantId, SoftwareBitmapSource source)
        {
            if (_participantThumbnails.TryGetValue(participantId, out var border))
            {
                if (border.Tag is { } tag)
                {
                    var tagType = tag.GetType();
                    var imageProperty = tagType.GetProperty("Image");
                    if (imageProperty?.GetValue(tag) is Image image)
                    {
                        image.Source = source;
                    }
                }
            }
        }
    }
}
