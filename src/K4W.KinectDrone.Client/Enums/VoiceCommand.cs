using System.ComponentModel;

namespace K4W.KinectDrone.Client.Enums
{
    public enum VoiceCommand
    {
        [Description("Unknown")]
        Unknown = 0,
        [Description("Takeoff")]
        TakeOff = 1,
        [Description("Land")]
        Land = 2,
        [Description("Emergency Landing")]
        EmergencyLanding = 4
    }
}
