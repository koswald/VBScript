// Wraps System.Diagnostics.EventLog class' WriteEntry method for VBScript.

// The WScript.Shell object has a native LogEvent method that logs to the Windows event logs { Log: Application; source: WSH }

using System.Runtime.InteropServices;
using System.Diagnostics;

namespace VBScripting
{
    /// <summary> Provides system logging for VBScript. </summary>
    [ProgId( "VBScripting.EventLogger" ),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-1AF8-495F-AB4D-6C61BD463EA4")]
    public class EventLogger : IEventLogger
    {
        /// <summary> Writes the specified message to the Application event log (source=VBScripting). </summary>
        /// <parameters> message </parameters>
        public void log(string message)
        {
            if (!string.IsNullOrWhiteSpace(message))
            {
                EventLog logger = new EventLog("Application", ".", "VBScripting");
                logger.WriteEntry(message);
            }
        }
    }

    /// <summary> A COM Interface for VBScripting.EventLogger </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-1BF8-495F-AB4D-6C61BD463EA4")]
    public interface IEventLogger
    {
        /// <summary> </summary>
        [DispId(1)]
        void log(string message);
    }
}
