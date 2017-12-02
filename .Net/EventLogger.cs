
// Wraps one method of the EventLog class for VBScript,
// for illustration purposes.

// The WScript.Shell object has a native LogEvent method
// that logs to the Windows event logs
// { Log: Application; source: WSH }

using System.Runtime.InteropServices;
using System.Diagnostics;

namespace VBScripting
{
    /// <summary> A COM Interface for <see cref="EventLogger"/> </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-1BF8-495F-AB4D-6C61BD463EA4")]
    public interface IEventLogger
    {
        /// <summary> COM interface member for <see cref="EventLogger.log(string)"/>. </summary>
        [DispId(1)]
        void log(string message);
    }

    /// <summary> Provides system logging for VBScript. </summary>
    [ProgId("VBScripting.EventLogger"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-1AF8-495F-AB4D-6C61BD463EA4")]
    public class EventLogger : IEventLogger
    {
        /// <summary> Writes a message to the system log (Application/WSH). </summary>
        /// <param name="message"> The message to be logged. </param>
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
