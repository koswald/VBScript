using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;

// remove or comment the following line before compiling if Visual Studio isn't installed
[assembly:AssemblyKeyFileAttribute("WSHEventLogger.snk")]

namespace EventLogging
{
    [Guid("2650C2AB-1BF8-495F-AB4D-6C61BD463EA4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IEventLogger
    {
        [DispId(1)]
        void log(string message);
    }

    [Guid("2650C2AB-1AF8-495F-AB4D-6C61BD463EA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("WSHEventLogger")]
    public class EventLogger : IEventLogger
    {
        public void log(string message)
        {
            if (!string.IsNullOrWhiteSpace(message))
            {
                EventLog logger = new EventLog("Application", ".", "WSH");
                logger.WriteEntry(message);
            }
        }
    }
}