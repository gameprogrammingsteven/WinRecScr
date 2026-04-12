using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Windows.Devices.Enumeration;
using Windows.Media.Capture;
using Windows.Media.MediaProperties;
using Windows.Storage;

namespace A_Simple_Recorder.Services
{
    public class CameraService
    {
        private MediaCapture? _mediaCapture;
        private bool _isRecording;

        public async Task<List<CameraInfo>> GetAvailableCamerasAsync()
        {
            var cameras = new List<CameraInfo>();

            try
            {
                // Find all video capture devices
                var devices = await DeviceInformation.FindAllAsync(DeviceClass.VideoCapture);

                foreach (var device in devices)
                {
                    cameras.Add(new CameraInfo
                    {
                        Id = device.Id,
                        Name = device.Name,
                        DeviceInformation = device
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error enumerating cameras: {ex.Message}");
            }

            return cameras;
        }

        public async Task StartRecordingAsync(CameraInfo camera, StorageFolder? saveFolder, string fileName)
        {
            if (_isRecording)
                return;

            try
            {
                _mediaCapture = new MediaCapture();

                var settings = new MediaCaptureInitializationSettings
                {
                    VideoDeviceId = camera.Id,
                    StreamingCaptureMode = StreamingCaptureMode.Video
                };

                await _mediaCapture.InitializeAsync(settings);

                // Create a file for recording
                var folder = saveFolder ?? ApplicationData.Current.LocalFolder;
                var file = await folder.CreateFileAsync(
                    fileName,
                    CreationCollisionOption.GenerateUniqueName);

                // Start recording
                var encodingProfile = MediaEncodingProfile.CreateMp4(VideoEncodingQuality.Auto);
                await _mediaCapture.StartRecordToStorageFileAsync(encodingProfile, file);

                _isRecording = true;
                System.Diagnostics.Debug.WriteLine($"Recording to: {file.Path}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error starting camera recording: {ex.Message}");
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
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping camera recording: {ex.Message}");
                throw;
            }
        }
    }
}
