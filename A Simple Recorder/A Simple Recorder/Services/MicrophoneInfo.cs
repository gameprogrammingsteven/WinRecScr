using Windows.Devices.Enumeration;

namespace A_Simple_Recorder.Services
{
    public class MicrophoneInfo
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public DeviceInformation DeviceInformation { get; set; } = null!;
    }
}
