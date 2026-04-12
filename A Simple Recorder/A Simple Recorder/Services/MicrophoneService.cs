using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Windows.Devices.Enumeration;
using Windows.Media.Capture;
using Windows.Media.MediaProperties;
using Windows.Storage;
using Windows.Media;
using Windows.Media.Audio;
using System.Threading;
using WinRT;

namespace A_Simple_Recorder.Services
{
    public class MicrophoneService
    {
        private MediaCapture? _mediaCapture;
        private bool _isRecording;
        private AudioGraph? _audioGraph;
        private AudioDeviceInputNode? _deviceInputNode;
        private AudioFrameOutputNode? _frameOutputNode;
        private float _currentAudioLevel = 0f;
        private bool _isMonitoring = false;
        private System.Threading.Timer? _levelCheckTimer;
        private readonly object _audioLock = new object();
        private volatile bool _isProcessingAudio = false;

        public async Task<List<MicrophoneInfo>> GetAvailableMicrophonesAsync()
        {
            var microphones = new List<MicrophoneInfo>();

            try
            {
                // Find all audio capture devices
                var devices = await DeviceInformation.FindAllAsync(DeviceClass.AudioCapture);

                foreach (var device in devices)
                {
                    microphones.Add(new MicrophoneInfo
                    {
                        Id = device.Id,
                        Name = device.Name,
                        DeviceInformation = device
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error enumerating microphones: {ex.Message}");
            }

            return microphones;
        }

        public async Task StartRecordingAsync()
        {
            if (_isRecording)
                return;

            try
            {
                _mediaCapture = new MediaCapture();

                var settings = new MediaCaptureInitializationSettings
                {
                    StreamingCaptureMode = StreamingCaptureMode.Audio
                };

                await _mediaCapture.InitializeAsync(settings);

                // Create a file for recording
                var fileName = $"Audio_{DateTime.Now:yyyyMMdd_HHmmss}.m4a";
                var file = await ApplicationData.Current.LocalFolder.CreateFileAsync(
                    fileName,
                    CreationCollisionOption.GenerateUniqueName);

                // Start recording
                var encodingProfile = MediaEncodingProfile.CreateM4a(AudioEncodingQuality.Auto);
                await _mediaCapture.StartRecordToStorageFileAsync(encodingProfile, file);

                _isRecording = true;
                System.Diagnostics.Debug.WriteLine($"Microphone recording to: {file.Path}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error starting microphone recording: {ex.Message}");
                throw;
            }
        }

        public async Task StopRecordingAsync()
        {
            if (!_isRecording || _mediaCapture == null)
                return;

            try
            {
                await _mediaCapture.StopRecordAsync();
                _isRecording = false;

                _mediaCapture.Dispose();
                _mediaCapture = null;

                System.Diagnostics.Debug.WriteLine("Microphone recording stopped");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping microphone recording: {ex.Message}");
                throw;
            }
        }

        // Start monitoring audio levels without recording
        public async Task<bool> StartMonitoringAsync(string? deviceId = null)
        {
            if (_isMonitoring)
                return true;

            try
            {
                var settings = new AudioGraphSettings(Windows.Media.Render.AudioRenderCategory.Communications)
                {
                    QuantumSizeSelectionMode = QuantumSizeSelectionMode.LowestLatency
                };

                var result = await AudioGraph.CreateAsync(settings);
                if (result.Status != AudioGraphCreationStatus.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create AudioGraph: {result.Status}");
                    return false;
                }

                _audioGraph = result.Graph;

                // Create device input node
                CreateAudioDeviceInputNodeResult deviceInputNodeResult;
                if (!string.IsNullOrEmpty(deviceId))
                {
                    // Find the device
                    var devices = await DeviceInformation.FindAllAsync(DeviceClass.AudioCapture);
                    var device = devices.FirstOrDefault(d => d.Id == deviceId);
                    if (device != null)
                    {
                        deviceInputNodeResult = await _audioGraph.CreateDeviceInputNodeAsync(
                            Windows.Media.Capture.MediaCategory.Communications,
                            _audioGraph.EncodingProperties,
                            device);
                    }
                    else
                    {
                        deviceInputNodeResult = await _audioGraph.CreateDeviceInputNodeAsync(
                            Windows.Media.Capture.MediaCategory.Communications);
                    }
                }
                else
                {
                    deviceInputNodeResult = await _audioGraph.CreateDeviceInputNodeAsync(
                        Windows.Media.Capture.MediaCategory.Communications);
                }

                if (deviceInputNodeResult.Status != AudioDeviceNodeCreationStatus.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create input node: {deviceInputNodeResult.Status}");
                    return false;
                }

                _deviceInputNode = deviceInputNodeResult.DeviceInputNode;

                // Create frame output node to read audio data
                _frameOutputNode = _audioGraph.CreateFrameOutputNode();
                _deviceInputNode.AddOutgoingConnection(_frameOutputNode);

                // Subscribe to quantum started event to calculate audio levels
                _audioGraph.QuantumStarted += OnQuantumStarted;

                _audioGraph.Start();
                _isMonitoring = true;

                System.Diagnostics.Debug.WriteLine("Audio monitoring started");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error starting audio monitoring: {ex.Message}");
                return false;
            }
        }

        public void StopMonitoring()
        {
            if (!_isMonitoring)
                return;

            try
            {
                if (_audioGraph != null)
                {
                    _audioGraph.QuantumStarted -= OnQuantumStarted;
                    _audioGraph.Stop();

                    // Wait for any in-progress audio processing to complete
                    var timeout = DateTime.UtcNow.AddMilliseconds(500);
                    while (_isProcessingAudio && DateTime.UtcNow < timeout)
                    {
                        Thread.Sleep(10);
                    }

                    _audioGraph.Dispose();
                    _audioGraph = null;
                }

                _deviceInputNode = null;
                _frameOutputNode = null;
                _isMonitoring = false;
                _currentAudioLevel = 0f;

                System.Diagnostics.Debug.WriteLine("Audio monitoring stopped");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping audio monitoring: {ex.Message}");
            }
        }

        private void OnQuantumStarted(AudioGraph sender, object args)
        {
            if (_frameOutputNode == null || !_isMonitoring)
                return;

            _isProcessingAudio = true;
            try
            {
                // Get audio frame from the output node
                var frame = _frameOutputNode.GetFrame();
                if (frame != null)
                {
                    using (var buffer = frame.LockBuffer(AudioBufferAccessMode.Read))
                    using (var reference = buffer.CreateReference())
                    {
                        unsafe
                        {
                            // Use proper WinRT-to-COM interop casting
                            byte* dataInBytes;
                            uint capacityInBytes;
                            var byteAccess = reference.As<IMemoryBufferByteAccess>();
                            byteAccess.GetBuffer(out dataInBytes, out capacityInBytes);

                            // Calculate RMS (Root Mean Square) for audio level
                            float sum = 0;
                            int sampleCount = (int)capacityInBytes / sizeof(float);
                            float* dataInFloat = (float*)dataInBytes;

                            for (int i = 0; i < sampleCount && i < 1000; i++) // Limit samples for performance
                            {
                                float sample = dataInFloat[i];
                                sum += sample * sample;
                            }

                            if (sampleCount > 0)
                            {
                                float rms = (float)Math.Sqrt(sum / Math.Min(sampleCount, 1000));
                                lock (_audioLock)
                                {
                                    _currentAudioLevel = rms * 10f; // Amplify for better visualization
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing audio: {ex.Message}");
            }
            finally
            {
                _isProcessingAudio = false;
            }
        }

        public float GetCurrentAudioLevel()
        {
            lock (_audioLock)
            {
                return _currentAudioLevel;
            }
        }

        public bool IsMonitoring => _isMonitoring;
    }
}
