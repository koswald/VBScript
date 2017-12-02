// fixture for DotNetCompiler.spec.elev.vbs

using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;

namespace EventLogging
{
    [Guid("2650C2AD-1BF8-495F-AB4D-6C61BD463EA4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IEventLogger
    {
        [DispId(1)]
        void log(string message);
    }

    [Guid("2650C2AD-1AF8-495F-AB4D-6C61BD463EA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("SourceCode1")]
    public class EventLogger : IEventLogger
    {
        public void log(string message)
        {
            EventLog logger = new EventLog("Application", ".", "WSH");
            logger.WriteEntry(message);
        }
    }
}