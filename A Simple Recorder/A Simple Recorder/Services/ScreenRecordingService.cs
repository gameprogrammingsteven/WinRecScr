using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using Windows.Graphics.Capture;
using Microsoft.UI.Xaml;
using Windows.Graphics.Imaging;
using Microsoft.UI.Dispatching;
using Microsoft.Graphics.Canvas;
using Windows.Storage;
using Windows.Graphics.DirectX;
using SharpAvi.Output;
using SharpAvi.Codecs;
using System.IO;
using Windows.Media.MediaProperties;
using WinRT;
using Windows.Media;
using Windows.Media.Audio;
using System.Collections.Concurrent;

namespace A_Simple_Recorder.Services
{
    public class ScreenRecordingService
    {
        private GraphicsCaptureItem? _captureItem;
        private Direct3D11CaptureFramePool? _framePool;
        private GraphicsCaptureSession? _session;
        private bool _isCapturing;
        private bool _isRecording;
        private Action<SoftwareBitmap>? _onFrameCaptured;
        private DispatcherQueue? _dispatcherQueue;
        private CanvasDevice? _canvasDevice;
        private StorageFolder? _saveFolder;
        private string? _fileName;
        private DateTime _lastPreviewTime = DateTime.MinValue;
        private readonly TimeSpan _previewInterval = TimeSpan.FromMilliseconds(100);
        private StorageFile? _outputFile;
        private AviWriter? _aviWriter;
        private IAviVideoStream? _videoStream;
        private IAviAudioStream? _audioStream;
        private int _frameWidth;
        private int _frameHeight;
        private AudioGraph? _audioGraph;
        private AudioDeviceInputNode? _audioInputNode;
        private AudioFrameOutputNode? _audioFrameOutputNode;
        private ConcurrentQueue<byte[]> _audioBuffer = new ConcurrentQueue<byte[]>();
        private int _audioBufferSize = 0;
        private readonly object _audioLock = new object();
        private volatile bool _isShuttingDown = false;
        private string? _microphoneDeviceId;
        private int _actualFrameCount = 0;
        private DateTime _recordingStartTime;
        private const int TARGET_FPS = 30;

        // Frame writing timer and current frame storage
        private Timer? _frameWriteTimer;
        private byte[]? _currentFrameData;
        private readonly object _frameLock = new object();

        // Start preview only (no file recording)
        public async Task StartPreviewAsync(Window window, Action<SoftwareBitmap>? onFrameCaptured = null)
        {
            if (_isCapturing)
                return;

            try
            {
                _onFrameCaptured = onFrameCaptured;
                _dispatcherQueue = DispatcherQueue.GetForCurrentThread();

                // Create a GraphicsCapturePicker
                var picker = new GraphicsCapturePicker();

                // Get the window handle
                var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(window);
                WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

                // Let the user pick what to capture
                _captureItem = await picker.PickSingleItemAsync();

                if (_captureItem == null)
                    return;

                // Store dimensions
                _frameWidth = _captureItem.Size.Width;
                _frameHeight = _captureItem.Size.Height;

                // Create capture session (no file writing)
                _canvasDevice = new CanvasDevice();

                _framePool = Direct3D11CaptureFramePool.CreateFreeThreaded(
                    _canvasDevice,
                    DirectXPixelFormat.B8G8R8A8UIntNormalized,
                    2,
                    _captureItem.Size);

                _framePool.FrameArrived += OnFrameArrived;
                _session = _framePool.CreateCaptureSession(_captureItem);
                _session.StartCapture();

                _isCapturing = true;
                System.Diagnostics.Debug.WriteLine($"Screen preview started");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error starting screen preview: {ex.Message}");
                throw;
            }
        }

        // Start recording to file (can be called while previewing)
        public async Task StartFileRecordingAsync(StorageFolder? saveFolder, string fileName, string? microphoneDeviceId = null)
        {
            if (_isRecording || !_isCapturing)
                return;

            _microphoneDeviceId = microphoneDeviceId;

            try
            {
                _saveFolder = saveFolder ?? ApplicationData.Current.LocalFolder;

                // Change extension to .avi since we're using SharpAvi
                _fileName = Path.ChangeExtension(fileName, ".avi");

                // Create output file
                _outputFile = await _saveFolder.CreateFileAsync(_fileName, CreationCollisionOption.GenerateUniqueName);

                // Initialize AVI writer at 30 FPS (we'll duplicate frames to match this)
                _aviWriter = new AviWriter(_outputFile.Path)
                {
                    FramesPerSecond = TARGET_FPS,
                    EmitIndex1 = true
                };

                // Create uncompressed video stream
                _videoStream = _aviWriter.AddUncompressedVideoStream(_frameWidth, _frameHeight);
                _videoStream.Width = _frameWidth;
                _videoStream.Height = _frameHeight;

                // Add audio stream (PCM 16-bit, 44100 Hz, stereo)
                _audioStream = _aviWriter.AddAudioStream(
                    channelCount: 2,
                    samplesPerSecond: 44100,
                    bitsPerSample: 16
                );

                // Initialize state
                _actualFrameCount = 0;
                _recordingStartTime = DateTime.Now;

                // Set recording flag BEFORE starting audio so OnAudioQuantumStarted can capture frames
                _isRecording = true;

                // Initialize audio capture
                await InitializeMicrophoneAsync();

                // Start the frame writing timer - fires every ~33.3ms for 30 FPS
                // This ensures we write exactly 30 frames per second regardless of capture rate
                _frameWriteTimer = new Timer(WriteFrameTimerCallback, null, 0, 1000 / TARGET_FPS);

                System.Diagnostics.Debug.WriteLine($"File recording started - saving to {_fileName}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error starting file recording: {ex.Message}");
                throw;
            }
        }

        // Legacy method that does both preview and recording
        public async Task StartRecordingAsync(Window window, StorageFolder? saveFolder, string fileName, Action<SoftwareBitmap>? onFrameCaptured = null, string? microphoneDeviceId = null)
        {
            await StartPreviewAsync(window, onFrameCaptured);
            if (_isCapturing)
            {
                await StartFileRecordingAsync(saveFolder, fileName, microphoneDeviceId);
            }
        }

        private async Task InitializeMicrophoneAsync()
        {
            try
            {
                var settings = new AudioGraphSettings(Windows.Media.Render.AudioRenderCategory.Media)
                {
                    EncodingProperties = AudioEncodingProperties.CreatePcm(44100, 2, 16)
                };

                var result = await AudioGraph.CreateAsync(settings);
                if (result.Status != AudioGraphCreationStatus.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create audio graph: {result.Status}");
                    return;
                }

                _audioGraph = result.Graph;

                // Create device input node with selected microphone
                CreateAudioDeviceInputNodeResult deviceInputNodeResult;
                if (!string.IsNullOrEmpty(_microphoneDeviceId))
                {
                    // Find the specific device
                    var devices = await Windows.Devices.Enumeration.DeviceInformation.FindAllAsync(Windows.Devices.Enumeration.DeviceClass.AudioCapture);
                    var device = devices.FirstOrDefault(d => d.Id == _microphoneDeviceId);
                    if (device != null)
                    {
                        deviceInputNodeResult = await _audioGraph.CreateDeviceInputNodeAsync(
                            Windows.Media.Capture.MediaCategory.Media,
                            _audioGraph.EncodingProperties,
                            device);
                        System.Diagnostics.Debug.WriteLine($"Using selected microphone: {device.Name}");
                    }
                    else
                    {
                        // Fall back to default
                        deviceInputNodeResult = await _audioGraph.CreateDeviceInputNodeAsync(
                            Windows.Media.Capture.MediaCategory.Media);
                        System.Diagnostics.Debug.WriteLine("Selected microphone not found, using default");
                    }
                }
                else
                {
                    // Use default microphone
                    deviceInputNodeResult = await _audioGraph.CreateDeviceInputNodeAsync(
                        Windows.Media.Capture.MediaCategory.Media);
                    System.Diagnostics.Debug.WriteLine("Using default microphone");
                }

                if (deviceInputNodeResult.Status != AudioDeviceNodeCreationStatus.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create audio input node: {deviceInputNodeResult.Status}");
                    return;
                }

                _audioInputNode = deviceInputNodeResult.DeviceInputNode;

                // Create frame output node with explicit encoding to ensure consistent format
                // Note: AudioGraph always processes in float format internally
                _audioFrameOutputNode = _audioGraph.CreateFrameOutputNode(_audioGraph.EncodingProperties);
                _audioInputNode.AddOutgoingConnection(_audioFrameOutputNode);

                // Log actual audio format being used
                var actualProps = _audioGraph.EncodingProperties;
                System.Diagnostics.Debug.WriteLine($"Audio graph format: {actualProps.SampleRate}Hz, {actualProps.ChannelCount}ch, {actualProps.BitsPerSample}bit, Subtype={actualProps.Subtype}");

                // Boost audio gain for better volume (2.0 = 200% = +6dB)
                _audioInputNode.OutgoingGain = 2.0;

                _audioGraph.QuantumStarted += OnAudioQuantumStarted;
                _audioGraph.Start();

                System.Diagnostics.Debug.WriteLine("Audio capture started");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error initializing audio: {ex.Message}");
            }
        }

        private void OnFrameArrived(Direct3D11CaptureFramePool sender, object args)
        {
            if (!_isCapturing || _isShuttingDown)
                return;

            try
            {
                var frame = sender.TryGetNextFrame();  // No 'using' - we'll manage disposal manually
                if (frame == null)
                    return;

                var now = DateTime.Now;

                _ = Task.Run(async () =>
                {
                    try
                    {
                        var bitmap = await SoftwareBitmap.CreateCopyFromSurfaceAsync(frame.Surface);

                        try
                        {
                            if (bitmap == null)
                                return;

                            // Store the current frame data for the timer to write
                            if (_isRecording && !_isShuttingDown)
                            {
                                var frameData = ConvertBitmapToBgra(bitmap);
                                if (frameData != null)
                                {
                                    lock (_frameLock)
                                    {
                                        _currentFrameData = frameData;
                                    }
                                }
                            }

                            // Update preview (at slower preview rate to reduce UI load)
                            if (now - _lastPreviewTime >= _previewInterval && _onFrameCaptured != null && _dispatcherQueue != null && !_isShuttingDown)
                            {
                                _lastPreviewTime = now;

                                // Make a copy for preview
                                var previewBitmap = SoftwareBitmap.Copy(bitmap);

                                _dispatcherQueue.TryEnqueue(() =>
                                {
                                    try
                                    {
                                        _onFrameCaptured?.Invoke(previewBitmap);
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error in preview callback: {ex.Message}");
                                        previewBitmap?.Dispose();
                                    }
                                });
                            }
                        }
                        finally
                        {
                            bitmap?.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing frame async: {ex.Message}");
                    }
                    finally
                    {
                        frame?.Dispose();  // Always dispose frame, even if CreateCopyFromSurfaceAsync fails
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error capturing frame: {ex.Message}");
            }
        }

        // Timer callback that writes frames at exactly 30 FPS
        private void WriteFrameTimerCallback(object? state)
        {
            if (!_isRecording || _isShuttingDown || _videoStream == null)
                return;

            try
            {
                byte[]? frameToWrite = null;

                lock (_frameLock)
                {
                    frameToWrite = _currentFrameData;
                }

                // Write video frame (actual frame or black frame for timing)
                if (frameToWrite != null)
                {
                    _videoStream.WriteFrame(true, frameToWrite, 0, frameToWrite.Length);
                }
                else
                {
                    // No frame available yet - write a black frame to maintain timing
                    var blackFrame = new byte[_frameWidth * _frameHeight * 4];
                    _videoStream.WriteFrame(true, blackFrame, 0, blackFrame.Length);
                }
                _actualFrameCount++;

                // ALWAYS write audio data from buffer (1/30th of a second worth)
                // This must happen regardless of whether we have a video frame
                if (_audioStream != null)
                {
                    // 44100 samples/sec * 2 channels * 2 bytes/sample / 30 FPS = 5880 bytes per frame
                    const int bytesPerFrame = (44100 * 2 * 2) / TARGET_FPS;

                    lock (_audioLock)
                    {
                        var audioData = new byte[bytesPerFrame];
                        int offset = 0;

                        // Get audio from buffer
                        while (offset < bytesPerFrame && _audioBuffer.TryDequeue(out var chunk))
                        {
                            int copySize = Math.Min(chunk.Length, bytesPerFrame - offset);
                            Array.Copy(chunk, 0, audioData, offset, copySize);
                            offset += copySize;
                            _audioBufferSize -= copySize;

                            if (copySize < chunk.Length)
                            {
                                var remainder = new byte[chunk.Length - copySize];
                                Array.Copy(chunk, copySize, remainder, 0, remainder.Length);
                                _audioBuffer.Enqueue(remainder);
                                _audioBufferSize += remainder.Length;
                            }
                        }

                        // Log audio write status
                        if (_actualFrameCount % 30 == 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"Writing audio: {offset}/{bytesPerFrame} bytes from buffer (buffer size: {_audioBufferSize})");
                        }

                        // Rest is already zeros (silence) if we didn't fill the buffer
                        _audioStream.WriteBlock(audioData, 0, bytesPerFrame);
                    }
                }

                if (_actualFrameCount % 30 == 0)
                {
                    var elapsed = (DateTime.Now - _recordingStartTime).TotalSeconds;
                    System.Diagnostics.Debug.WriteLine($"Recording: {_actualFrameCount} frames ({_actualFrameCount / (double)TARGET_FPS:F1}s video, {elapsed:F1}s real)");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in frame write timer: {ex.Message}");
            }
        }

        private byte[]? ConvertBitmapToBgra(SoftwareBitmap bitmap)
        {
            try
            {
                // Ensure correct format
                if (bitmap.BitmapPixelFormat != BitmapPixelFormat.Bgra8)
                {
                    bitmap = SoftwareBitmap.Convert(bitmap, BitmapPixelFormat.Bgra8);
                }

                // Get pixel data
                using var buffer = bitmap.LockBuffer(BitmapBufferAccessMode.Read);
                using var reference = buffer.CreateReference();

                unsafe
                {
                    // Use proper WinRT-to-COM interop casting
                    byte* data;
                    uint capacity;
                    var byteAccess = reference.As<IMemoryBufferByteAccess>();
                    byteAccess.GetBuffer(out data, out capacity);

                    // Copy to managed array
                    var bytes = new byte[capacity];
                    System.Runtime.InteropServices.Marshal.Copy((IntPtr)data, bytes, 0, (int)capacity);

                    return bytes;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error converting bitmap: {ex.Message}");
                return null;
            }
        }

        private void OnAudioQuantumStarted(AudioGraph sender, object args)
        {
            if (_audioFrameOutputNode == null || !_isRecording || _isShuttingDown)
            {
                return;
            }

            try
            {
                var frame = _audioFrameOutputNode.GetFrame();
                if (frame == null)
                {
                    return;
                }

                using (var buffer = frame.LockBuffer(AudioBufferAccessMode.Read))
                using (var reference = buffer.CreateReference())
                {
                    unsafe
                    {
                        byte* data;
                        uint capacity;
                        var byteAccess = reference.As<IMemoryBufferByteAccess>();
                        byteAccess.GetBuffer(out data, out capacity);

                        if (capacity == 0)
                            return;

                        byte[] pcmData;
                        int nonZeroSamples = 0;

                        // Check the encoding subtype to determine format
                        // AudioGraph internally uses Float format even when PCM is requested
                        var subtype = _audioGraph?.EncodingProperties?.Subtype ?? "";

                        if (subtype.Equals("Float", StringComparison.OrdinalIgnoreCase))
                        {
                            // Float format: convert to 16-bit PCM
                            int floatSampleCount = (int)capacity / sizeof(float);
                            float* floatData = (float*)data;
                            pcmData = new byte[floatSampleCount * sizeof(short)];

                            for (int i = 0; i < floatSampleCount; i++)
                            {
                                float sample = Math.Max(-1.0f, Math.Min(1.0f, floatData[i]));
                                short pcmSample = (short)(sample * short.MaxValue);
                                if (pcmSample != 0) nonZeroSamples++;
                                pcmData[i * 2] = (byte)(pcmSample & 0xFF);
                                pcmData[i * 2 + 1] = (byte)((pcmSample >> 8) & 0xFF);
                            }
                        }
                        else
                        {
                            // PCM format: data is already 16-bit PCM, just copy it
                            pcmData = new byte[capacity];
                            System.Runtime.InteropServices.Marshal.Copy((IntPtr)data, pcmData, 0, (int)capacity);

                            // Count non-zero samples for logging
                            for (int i = 0; i < capacity / 2; i++)
                            {
                                short sample = (short)(pcmData[i * 2] | (pcmData[i * 2 + 1] << 8));
                                if (sample != 0) nonZeroSamples++;
                            }
                        }

                        lock (_audioLock)
                        {
                            _audioBuffer.Enqueue(pcmData);
                            _audioBufferSize += pcmData.Length;

                            // Log less frequently to reduce spam
                            if (_actualFrameCount % 30 == 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"Audio captured: {pcmData.Length} bytes, {nonZeroSamples} non-zero samples, format={subtype}, buffer={_audioBufferSize}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error capturing audio: {ex.Message}");
            }
        }

        // Stop file recording only (keep preview running)
        public async Task<string> StopFileRecordingAsync()
        {
            if (!_isRecording)
                return string.Empty;

            try
            {
                // Stop recording flag first to prevent new frame writes
                _isRecording = false;

                // Stop the frame writing timer
                if (_frameWriteTimer != null)
                {
                    await _frameWriteTimer.DisposeAsync();
                    _frameWriteTimer = null;
                }

                // Log recording statistics
                if (_actualFrameCount > 0)
                {
                    var recordingDuration = (DateTime.Now - _recordingStartTime).TotalSeconds;
                    var videoFileDuration = _actualFrameCount / (double)TARGET_FPS;
                    System.Diagnostics.Debug.WriteLine($"Recording stats:");
                    System.Diagnostics.Debug.WriteLine($"  Recording session: {recordingDuration:F2}s");
                    System.Diagnostics.Debug.WriteLine($"  Frames written: {_actualFrameCount}");
                    System.Diagnostics.Debug.WriteLine($"  Video duration: {videoFileDuration:F2}s at {TARGET_FPS} FPS");
                    System.Diagnostics.Debug.WriteLine($"  Time ratio: {(videoFileDuration / recordingDuration):F2}x (should be ~1.0)");
                }

                // Stop audio capture first
                if (_audioGraph != null)
                {
                    _audioGraph.QuantumStarted -= OnAudioQuantumStarted;
                    _audioGraph.Stop();
                }

                // Collect remaining audio buffer (do this before disposing audio graph)
                List<byte[]> remainingAudio = new List<byte[]>();
                lock (_audioLock)
                {
                    System.Diagnostics.Debug.WriteLine($"Collecting remaining audio buffer: {_audioBufferSize} bytes");
                    while (_audioBuffer.TryDequeue(out var chunk))
                    {
                        remainingAudio.Add(chunk);
                    }
                    _audioBufferSize = 0;
                }

                // Write remaining audio outside of lock (faster, no blocking)
                if (_audioStream != null && remainingAudio.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"Flushing {remainingAudio.Count} audio chunks to file");
                    foreach (var chunk in remainingAudio)
                    {
                        _audioStream.WriteBlock(chunk, 0, chunk.Length);
                    }
                }

                // Dispose audio graph
                if (_audioGraph != null)
                {
                    _audioGraph.Dispose();
                    _audioGraph = null;
                }
                _audioInputNode = null;
                _audioFrameOutputNode = null;

                // Clear current frame data
                lock (_frameLock)
                {
                    _currentFrameData = null;
                }

                // Close AVI writer properly
                if (_aviWriter != null)
                {
                    try
                    {
                        _videoStream = null;
                        _audioStream = null;
                        _aviWriter.Close();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error closing AVI writer: {ex.Message}");
                    }
                    finally
                    {
                        _aviWriter = null;
                    }
                }

                // Clear audio buffer
                lock (_audioLock)
                {
                    _audioBuffer = new ConcurrentQueue<byte[]>();
                    _audioBufferSize = 0;
                }

                var fullPath = _outputFile?.Path ?? "unknown";
                System.Diagnostics.Debug.WriteLine($"File recording stopped - saved to {fullPath}");

                return fullPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping file recording: {ex.Message}");
                throw;
            }
        }

        // Stop preview (and recording if active)
        public async Task StopPreviewAsync()
        {
            if (!_isCapturing)
                return;

            try
            {
                // Stop recording first if active
                if (_isRecording)
                {
                    await StopFileRecordingAsync();
                }

                // Stop capturing to prevent new frames
                _isCapturing = false;

                // Stop capture session
                _session?.Dispose();
                _session = null;

                if (_framePool != null)
                {
                    _framePool.FrameArrived -= OnFrameArrived;
                    _framePool.Dispose();
                    _framePool = null;
                }

                _canvasDevice?.Dispose();
                _canvasDevice = null;
                _captureItem = null;

                // Clear frame data
                lock (_frameLock)
                {
                    _currentFrameData = null;
                }

                System.Diagnostics.Debug.WriteLine($"Screen preview stopped");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping screen preview: {ex.Message}");
                throw;
            }
        }

        // Legacy method: Stop both recording and preview
        public async Task<string> StopRecordingAsync()
        {
            var path = string.Empty;
            if (_isRecording)
            {
                path = await StopFileRecordingAsync();
            }
            await StopPreviewAsync();
            return path;
        }

        public bool IsCapturing => _isCapturing;
        public bool IsRecording => _isRecording;

        public void SignalShutdown()
        {
            _isShuttingDown = true;
        }
    }
}
