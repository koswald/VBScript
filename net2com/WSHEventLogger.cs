using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;

[assembly:AssemblyKeyFileAttribute("WSHEventLogger.snk")]

namespace EventLogging
{
    [Guid("26508E95-8A27-4ae6-B6DE-0542A0FC7039")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _EventLogger
    {
        [DispId(1)]
        void log(string message);
    }

    [Guid("265032AD-4BF8-495f-AB4D-6C61BD463EA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("WSHEventLogger")]
    public class EventLogger : _EventLogger
    {
        public void log(string message)
        {
            EventLog logger = new EventLog("Application", ".", "WSH");
            logger.WriteEntry(message);
        }
    }
}