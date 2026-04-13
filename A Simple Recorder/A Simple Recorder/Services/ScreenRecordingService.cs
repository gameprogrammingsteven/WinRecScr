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
        private uint _audioSampleRate = 48000;  // Will be set from actual audio graph
        private uint _sourceAudioChannels = 2;
        private uint _audioChannels = 2;

        // Test tone generator for debugging audio pipeline
        private bool _useTestTone = false;  // DISABLED: Use real microphone capture
        private double _testTonePhase = 0;

        // Frame writing timer and current frame storage
        private Timer? _frameWriteTimer;
        private byte[]? _currentFrameData;
        private readonly object _frameLock = new object();
        private byte[]? _audioRemainder;

        // Audio block size calculated once per recording session (per SharpAvi guidelines)
        private int _audioBlockSize;

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

                // Initialize state
                _actualFrameCount = 0;
                _recordingStartTime = DateTime.Now;

                // Set recording flag BEFORE starting audio so OnAudioQuantumStarted can capture frames
                _isRecording = true;

                // Initialize audio capture FIRST to get actual sample rate
                await InitializeMicrophoneAsync();

                // USE the native sample rate from AudioGraph - DO NOT override!
                // The AudioGraph captures at device native rate (usually 48000 Hz)
                // We must use the same rate for the WAV file header
                System.Diagnostics.Debug.WriteLine($"Using audio format: {_audioSampleRate}Hz, {_audioChannels}ch");

                // Initialize AVI writer at 30 FPS
                _aviWriter = new AviWriter(_outputFile.Path)
                {
                    FramesPerSecond = TARGET_FPS,
                    EmitIndex1 = true
                };

                // Create uncompressed video stream
                _videoStream = _aviWriter.AddUncompressedVideoStream(_frameWidth, _frameHeight);
                _videoStream.Width = _frameWidth;
                _videoStream.Height = _frameHeight;

                // Create audio stream in AVI - 16-bit PCM, matching AudioGraph output format
                // SharpAvi 3.x: Use AddAudioStream with waveFormat short (1 = PCM)
                _audioStream = _aviWriter.AddAudioStream(
                    channelCount: (int)_audioChannels,
                    samplesPerSecond: (int)_audioSampleRate,
                    bitsPerSample: 16);

                // Calculate audio block size ONCE using stream properties (per SharpAvi guidelines)
                // audioByteRate = (bitsPerSample / 8) * channelCount * samplesPerSecond
                // audioBlockSize = audioByteRate / framesPerSecond
                var audioByteRate = (_audioStream.BitsPerSample / 8) * _audioStream.ChannelCount * _audioStream.SamplesPerSecond;
                _audioBlockSize = (int)(audioByteRate / _aviWriter.FramesPerSecond);
                if (_audioBlockSize <= 0)
                {
                    _audioBlockSize = (_audioStream.BitsPerSample / 8) * _audioStream.ChannelCount;
                }

                // IMPORTANT: Verify format consistency
                System.Diagnostics.Debug.WriteLine($"=== AUDIO FORMAT VERIFICATION ===");
                System.Diagnostics.Debug.WriteLine($"AudioGraph captures at: {_audioSampleRate}Hz, {_sourceAudioChannels}ch");
                System.Diagnostics.Debug.WriteLine($"AVI stream expects: {_audioStream.SamplesPerSecond}Hz, {_audioStream.ChannelCount}ch, {_audioStream.BitsPerSample}-bit");
                System.Diagnostics.Debug.WriteLine($"Audio block size per frame: {_audioBlockSize} bytes");
                System.Diagnostics.Debug.WriteLine($"Samples per video frame: {_audioBlockSize / (_audioStream.ChannelCount * 2)}");
                System.Diagnostics.Debug.WriteLine($"=================================");

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
                // Create AudioGraph - let it use default/device-native settings for best compatibility
                var settings = new AudioGraphSettings(Windows.Media.Render.AudioRenderCategory.Media);

                var result = await AudioGraph.CreateAsync(settings);
                if (result.Status != AudioGraphCreationStatus.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create audio graph: {result.Status}");
                    return;
                }

                _audioGraph = result.Graph;

                // Create device input node with selected microphone
                // DO NOT pass encoding properties - let device use its native format (prevents resampling artifacts)
                CreateAudioDeviceInputNodeResult deviceInputNodeResult;
                if (!string.IsNullOrEmpty(_microphoneDeviceId))
                {
                    // Find the specific device
                    var devices = await Windows.Devices.Enumeration.DeviceInformation.FindAllAsync(Windows.Devices.Enumeration.DeviceClass.AudioCapture);
                    var device = devices.FirstOrDefault(d => d.Id == _microphoneDeviceId);
                    if (device != null)
                    {
                        // Create WITHOUT encoding properties - device uses native format
                        deviceInputNodeResult = await _audioGraph.CreateDeviceInputNodeAsync(
                            Windows.Media.Capture.MediaCategory.Media,
                            null,  // null = use device native format, no resampling
                            device);
                        System.Diagnostics.Debug.WriteLine($"Using selected microphone (native format): {device.Name}");
                    }
                    else
                    {
                        // Fall back to default - no encoding properties
                        deviceInputNodeResult = await _audioGraph.CreateDeviceInputNodeAsync(
                            Windows.Media.Capture.MediaCategory.Media);
                        System.Diagnostics.Debug.WriteLine("Selected microphone not found, using default (native format)");
                    }
                }
                else
                {
                    // Use default microphone - no encoding properties
                    deviceInputNodeResult = await _audioGraph.CreateDeviceInputNodeAsync(
                        Windows.Media.Capture.MediaCategory.Media);
                    System.Diagnostics.Debug.WriteLine("Using default microphone (native format)");
                }

                if (deviceInputNodeResult.Status != AudioDeviceNodeCreationStatus.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create audio input node: {deviceInputNodeResult.Status}");
                    return;
                }

                _audioInputNode = deviceInputNodeResult.DeviceInputNode;

                // Get ACTUAL sample rate and channels from the INPUT NODE (not the graph!)
                // The device might have different properties than the graph's default
                _audioSampleRate = _audioInputNode.EncodingProperties.SampleRate;
                _sourceAudioChannels = _audioInputNode.EncodingProperties.ChannelCount;
                System.Diagnostics.Debug.WriteLine($"Device input node format: {_audioSampleRate}Hz, {_sourceAudioChannels}ch");

                // Create frame output node WITHOUT forcing encoding properties
                // This avoids any internal resampling that could cause artifacts
                _audioFrameOutputNode = _audioGraph.CreateFrameOutputNode();
                _audioInputNode.AddOutgoingConnection(_audioFrameOutputNode);

                // Update sample rate/channels from the OUTPUT node (what we'll actually receive)
                _audioSampleRate = _audioFrameOutputNode.EncodingProperties.SampleRate;
                _sourceAudioChannels = _audioFrameOutputNode.EncodingProperties.ChannelCount;

                // Normalize recording to stereo for player compatibility
                _audioChannels = 2;

                // Log all formats for debugging
                System.Diagnostics.Debug.WriteLine($"Graph format: {_audioGraph.EncodingProperties.SampleRate}Hz, {_audioGraph.EncodingProperties.ChannelCount}ch");
                System.Diagnostics.Debug.WriteLine($"Input node format: {_audioInputNode.EncodingProperties.SampleRate}Hz, {_audioInputNode.EncodingProperties.ChannelCount}ch");
                System.Diagnostics.Debug.WriteLine($"Output node format: {_audioFrameOutputNode.EncodingProperties.SampleRate}Hz, {_audioFrameOutputNode.EncodingProperties.ChannelCount}ch");
                System.Diagnostics.Debug.WriteLine($"Recording audio format: {_audioSampleRate}Hz, {_audioChannels}ch (source {_sourceAudioChannels}ch)");

                // Boost audio gain for voice recording (5.0 = 500% = +14dB)
                _audioInputNode.OutgoingGain = 5.0;

                _audioGraph.QuantumStarted += OnAudioQuantumStarted;
                _audioGraph.Start();

                System.Diagnostics.Debug.WriteLine($"Audio capture started with gain={_audioInputNode.OutgoingGain}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR initializing audio: {ex.Message}");
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

        // Timer callback that writes frames at exactly 30 FPS (VIDEO ONLY)
        private void WriteFrameTimerCallback(object? state)
        {
            if (!_isRecording || _isShuttingDown || _videoStream == null)
                return;

            try
            {
                // Write video frame
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

                // Write audio to AVI - must be interleaved with video for proper playback
                if (_audioStream != null)
                {
                    // Use pre-calculated audio block size (per SharpAvi guidelines)
                    var audioBuffer = new byte[_audioBlockSize];
                    int offset = 0;

                    lock (_audioLock)
                    {
                        // Start with the remainder from the last run
                        if (_audioRemainder != null)
                        {
                            int copySize = Math.Min(_audioRemainder.Length, _audioBlockSize);
                            Array.Copy(_audioRemainder, 0, audioBuffer, 0, copySize);
                            offset += copySize;

                            if (copySize < _audioRemainder.Length)
                            {
                                // Not all of the remainder was used, save the rest for next time
                                var newRemainder = new byte[_audioRemainder.Length - copySize];
                                Array.Copy(_audioRemainder, copySize, newRemainder, 0, newRemainder.Length);
                                _audioRemainder = newRemainder;
                            }
                            else
                            {
                                _audioRemainder = null;
                            }
                        }

                        // Now fill the rest from the buffer
                        while (offset < _audioBlockSize && _audioBuffer.TryDequeue(out var chunk))
                        {
                            _audioBufferSize -= chunk.Length; // Decrement full chunk size

                            int copySize = Math.Min(chunk.Length, _audioBlockSize - offset);
                            Array.Copy(chunk, 0, audioBuffer, offset, copySize);
                            offset += copySize;

                            // If we didn't use the whole chunk, save the remainder for the next callback
                            if (copySize < chunk.Length)
                            {
                                _audioRemainder = new byte[chunk.Length - copySize];
                                Array.Copy(chunk, copySize, _audioRemainder, 0, _audioRemainder.Length);
                                // Break because we have filled the audioBuffer
                                break;
                            }
                        }
                    }

                    // Write audio block to AVI (even if partially filled with silence)
                    _audioStream.WriteBlock(audioBuffer, 0, audioBuffer.Length);
                }

                _actualFrameCount++;

                if (_actualFrameCount % 30 == 0)
                {
                    var elapsed = (DateTime.Now - _recordingStartTime).TotalSeconds;
                    System.Diagnostics.Debug.WriteLine($"Recording: {_actualFrameCount} frames ({_actualFrameCount / (double)TARGET_FPS:F1}s video, {elapsed:F1}s real), audio buffer={_audioBufferSize}");
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
                using var frame = _audioFrameOutputNode.GetFrame();
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

                        var validLength = (int)buffer.Length;
                        if (validLength <= 0)
                            return;

                        // AudioGraph returns 32-bit float samples.
                        // Use AudioBuffer.Length (valid data), not IMemoryBuffer capacity.
                        int floatSampleCount = validLength / sizeof(float);
                        if (floatSampleCount <= 0)
                            return;

                        int sourceChannels = (int)Math.Max(1, _sourceAudioChannels);
                        int targetChannels = (int)Math.Max(1, _audioChannels);
                        int frameCount = floatSampleCount / sourceChannels;
                        if (frameCount <= 0)
                            return;

                        float* floatData = (float*)data;

                        var pcmData = new byte[frameCount * targetChannels * sizeof(short)];
                        int nonZeroSamples = 0;
                        float maxSample = 0;

                        for (int frameIndex = 0; frameIndex < frameCount; frameIndex++)
                        {
                            int sourceOffset = frameIndex * sourceChannels;

                            // Downmix source channels to mono to handle devices exposing 8+ channels.
                            float mono = 0f;
                            for (int c = 0; c < sourceChannels; c++)
                            {
                                mono += floatData[sourceOffset + c];
                            }
                            mono /= sourceChannels;

                            float absSample = Math.Abs(mono);
                            if (absSample > maxSample) maxSample = absSample;

                            mono = Math.Max(-1.0f, Math.Min(1.0f, mono));
                            short pcmSample = (short)(mono * 32767.0f);

                            // Write mono to all target channels (stereo duplicate)
                            for (int tc = 0; tc < targetChannels; tc++)
                            {
                                int outputIndex = (frameIndex * targetChannels + tc) * 2;
                                pcmData[outputIndex] = (byte)(pcmSample & 0xFF);
                                pcmData[outputIndex + 1] = (byte)((pcmSample >> 8) & 0xFF);
                                if (pcmSample != 0) nonZeroSamples++;
                            }
                        }

                        lock (_audioLock)
                        {
                            _audioBuffer.Enqueue(pcmData);
                            _audioBufferSize += pcmData.Length;

                            if (_actualFrameCount % 30 == 0 && nonZeroSamples > 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"Audio: {pcmData.Length}B, {nonZeroSamples}/{frameCount * targetChannels} non-zero, max={maxSample:F4}, buffer={_audioBufferSize}B");
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
                string aviPath = _outputFile?.Path ?? "";
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
                    _audioRemainder = null;
                }

                System.Diagnostics.Debug.WriteLine($"File recording stopped - saved to {aviPath}");

                return aviPath;
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

        // Mux video and audio files using ffmpeg
        private async Task<string> MuxVideoAndAudioAsync(string videoPath, string audioPath)
        {
            try
            {
                // Output path: same as video but with _final suffix
                string outputPath = Path.Combine(
                    Path.GetDirectoryName(videoPath) ?? "",
                    Path.GetFileNameWithoutExtension(videoPath) + "_final.avi"
                );

                // Try to find ffmpeg
                string? ffmpegPath = FindFfmpeg();
                if (ffmpegPath == null)
                {
                    System.Diagnostics.Debug.WriteLine("ffmpeg not found - keeping separate video and audio files");
                    System.Diagnostics.Debug.WriteLine($"To combine manually, run: ffmpeg -i \"{videoPath}\" -i \"{audioPath}\" -c:v copy -c:a pcm_s16le \"{outputPath}\"");
                    return videoPath;
                }

                System.Diagnostics.Debug.WriteLine($"Found ffmpeg at: {ffmpegPath}");
                System.Diagnostics.Debug.WriteLine($"Muxing video and audio...");

                // Run ffmpeg to combine video and audio
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = ffmpegPath,
                    Arguments = $"-y -i \"{videoPath}\" -i \"{audioPath}\" -c:v copy -c:a pcm_s16le \"{outputPath}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using var process = System.Diagnostics.Process.Start(startInfo);
                if (process != null)
                {
                    string stderr = await process.StandardError.ReadToEndAsync();
                    await process.WaitForExitAsync();

                    if (process.ExitCode == 0 && File.Exists(outputPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"Successfully muxed to: {outputPath}");

                        // Delete the separate video and audio files
                        try
                        {
                            File.Delete(videoPath);
                            File.Delete(audioPath);
                            System.Diagnostics.Debug.WriteLine("Cleaned up temporary files");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Could not delete temp files: {ex.Message}");
                        }

                        return outputPath;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"ffmpeg failed with exit code {process.ExitCode}");
                        System.Diagnostics.Debug.WriteLine($"ffmpeg stderr: {stderr}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error muxing video and audio: {ex.Message}");
            }

            return videoPath;
        }

        // Try to find ffmpeg in common locations
        private string? FindFfmpeg()
        {
            // Check if ffmpeg is in PATH
            var pathDirs = Environment.GetEnvironmentVariable("PATH")?.Split(Path.PathSeparator) ?? Array.Empty<string>();
            foreach (var dir in pathDirs)
            {
                var ffmpegPath = Path.Combine(dir, "ffmpeg.exe");
                if (File.Exists(ffmpegPath))
                    return ffmpegPath;
            }

            // Check common installation locations
            string[] commonPaths = new[]
            {
                @"C:\ffmpeg\bin\ffmpeg.exe",
                @"C:\Program Files\ffmpeg\bin\ffmpeg.exe",
                @"C:\Program Files (x86)\ffmpeg\bin\ffmpeg.exe",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"ffmpeg\bin\ffmpeg.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @"ffmpeg\bin\ffmpeg.exe"),
            };

            foreach (var path in commonPaths)
            {
                if (File.Exists(path))
                    return path;
            }

            return null;
        }
    }
}
