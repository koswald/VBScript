
// Miscellaneous admin functions
// Not all may be available to VBScript

using System.Diagnostics;
using System.Windows.Forms;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System;
using System.Linq;
using System.Collections.Generic;

namespace VBScripting
{
    /// <summary> Provide sys admin features for C# and VBScript. </summary>
    [ProgId("VBScripting.Admin"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-8AF8-495F-AB4D-6C61BD463EA4")]
    public class Admin : IAdmin
    {
        /// <summary> Gets whether the current process has elevated privileges. </summary>
        public static bool PrivilegesAreElevated
        {
            get
            {
                bool areElevated;
                using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
                {
                    WindowsPrincipal principal = new WindowsPrincipal(identity);
                    areElevated = principal.IsInRole(WindowsBuiltInRole.Administrator);
                }
                return areElevated;
            }
        }

        /// <summary> VBScript wrapper for the static <see cref="PrivilegesAreElevated"/> property. </summary>
        public bool privilegesAreElevated
        {
            get
            {
                return PrivilegesAreElevated;
            }
        }

        # region EventLogs

        /// <summary> Name of the preferred event source for this namespace. </summary>
        private const string desiredEventSource = "VBScripting";
        /// <summary> Name of the default event source provided for WScript. </summary>
        private const string alternateEventSource = "WSH";
        /// <summary> Name of the log to which events will be logged. </summary>
        private const string logName = "Application";

        /// <summary> Logs an event to the Application event log. </summary>
        public static void Log(string msg)
        {
            try
            {
                new EventLog(logName, ".", desiredEventSource).WriteEntry(msg);
            }
            catch(Exception e)
            {
                string ex = e.Message;
                new EventLog(logName, ".", alternateEventSource).WriteEntry(msg);
            }
        }

        /// <summary> VBScript wrapper for the static <see cref="Log(string)"/> method. </summary>
        public void log(string msg)
        {
            Log(msg);
        }

        /// <summary> Get an array of logs entries from the Application log. </summary>
        /// <param name="source"> Event source </param>
        /// <param name="message"> A substring of the event message. </param>
        public object GetLogs(string source, string message)
        {
            EventLog log = new EventLog(logName);
            var entries = log.Entries.Cast<EventLogEntry>()
                                     .Where(x => x.Source == source & x.Message.Contains(message))
                                     .Select(x => new
                                     {
                                         x.Source,
                                         x.Message
                                     })
                                     .ToList();
            // convert entries to List<string>
            List<string> messages = new List<string>();
            foreach (var entry in entries)
            {
                messages.Add(entry.Message);
            }
            // convert messages to VBScript string array
            return messages.Cast<object>().ToArray();
        }

        /// <summary> Gets whether the EventLog source exists. </summary>
        /// <param name="source"></param>
        public bool SourceExists(string source)
        {
            try
            {
                return EventLog.SourceExists(source);
            }
            catch (Exception e)
            {
                Log(string.Format(
                    "Failed to determine whether the source \"{0}\" exists.\n{1}\n\n" +
                    "Privileges {2} elevated.\nElevated privileges are required.",
                    source, e.Message, PrivilegesAreElevated? "are" : "are not"
                ));
                throw;
            }
        }

        /// <summary> Create an EventLog source. </summary>
        /// <param name="source"></param>
        public object CreateEventSource(string source)
        {
            string msg;

            // check if event source exists
            if (this.SourceExists(source))
            {
                msg = string.Format(
                    "The source \"{0}\" already exists.",
                    source
                );
                return VBScriptArrayResult
                (
                    msg, 
                    this.Result.SourceAlreadyExists
                );
            }
            // privileges already checked on this.SourceExists call
            // create the source
            else
            {
                try
                {
                    EventLog.CreateEventSource(source, logName);
                    msg = string.Format(
                        "The EventLog source \"{0}\" has been created.",
                        source
                    );
                    return VBScriptArrayResult
                    (
                        msg,
                        this.Result.SourceCreated
                    );
                }
                catch (Exception e)
                {
                    Log(string.Format(
                        "Failed to create source \"{0}\".\n\n{1}",
                        source, e.ToString()
                    ));
                    throw;
                }
            }
        }
        // convert two result strings into a VBScript array
        private object VBScriptArrayResult(string msg, string result)
        {
            List<string> list = new List<string>();
            list.Add(msg);
            list.Add(result);
            return list.Cast<object>().ToArray();
        }
        /// <summary> Delete an EventLog source and all of its logs. </summary>
        /// <param name="source"></param>
        public object DeleteEventSource(string source)
        {
            string expectedLogName = logName;
            string msg;

            // check if event source exists
            // this also checks privileges
            if (!this.SourceExists(source))
            {
                msg = string.Format(
                    "The source \"{0}\" does not exist.",
                    source
                );
                return VBScriptArrayResult
                (
                    msg,
                    this.Result.SourceDoesNotExist
                );
            }

            // check that the source is in the expected log
            if (EventLog.LogNameFromSourceName(source, ".") != expectedLogName)
            {
                msg = string.Format(
                    "The source \"{0}\" exists " +
                    "but not in the expected log, \"{1}\"",
                    source, expectedLogName
                );
                return VBScriptArrayResult
                (
                    msg,
                    this.Result.SourceFoundInAnotherLog
                );
            }
            // delete the source
            else
            {
                try
                {
                    EventLog.DeleteEventSource(source);
                    msg = string.Format(
                        "The EventLog source \"{0}\" has been deleted.",
                        source
                    );
                    return VBScriptArrayResult
                    (
                        msg,
                        this.Result.SourceDeleted
                    );
                }
                catch (Exception e)
                {
                    Log(string.Format(
                        "Failed to delete source \"{0}\".\n" +
                        "Privileges {1} elevated.\nElevated privileges are required.\n\n{2}",
                        source,
                        PrivilegesAreElevated? "are" : "are not",
                        e.Message
                    ));
                    throw;
                }
            }
        }
        /// <summary> Returns the behavior enum object to VBScript as a property. </summary>
        public EventLogResultT Result
        {
            get { return new EventLogResultT(); }
        }
        # endregion EventLogs
    }
    /// <summary> Provides an enum of sorts for VBScript for reporting class behaviors. </summary>
    public class EventLogResultT // T for Type
    {
        /// <summary>  </summary>
        public string SourceFound { get { return "Source was found."; } } // when checking existence
        /// <summary>  </summary>
        public string SourceNotFoundLowPrivileges { get { return "Source not found; privileges are not elevated."; } } // Can't determine whether source exists
        /// <summary>  </summary>
        public string SourceNotFoundHighPrivileges { get { return "Source not found; privileges are elevated."; } } // source doesn't exist

        /// <summary>  </summary>
        public string SourceAlreadyExists { get { return "Source already exists."; } } // when attempting to create
        /// <summary>  </summary>
        public string SourceCreated { get { return "Source was created."; } }
        /// <summary>  </summary>
        public string SourceNotCreated { get { return "Source was not created."; } }

        /// <summary>  </summary>
        public string SourceDoesNotExist { get { return "Source does not exist."; } } // attempting to delete
        /// <summary>  </summary>
        public string SourceDeleted { get { return "Source was deleted successfully."; } }
        /// <summary>  </summary>
        public string SourceFoundInAnotherLog { get { return "Source was found in another log."; } }
        /// <summary>  </summary>
        public string SourceNotDeleted { get { return "Source was not deleted."; } }
    }

    /// <summary> COM interface for <see cref="Admin"/> </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-8BF8-495F-AB4D-6C61BD463EA4")]
    public interface IAdmin
    {
        /// <summary> COM interface member for <see cref="Admin.CreateEventSource(string)"/> </summary>
        [DispId(1)]
        object CreateEventSource(string source);

        /// <summary> COM interface member for <see cref="Admin.DeleteEventSource(string)"/> </summary>
        [DispId(2)]
        object DeleteEventSource(string source);

        /// <summary> ComInterface member for <see cref="Admin.SourceExists(string)"/> </summary>
        [DispId(3)]
        bool SourceExists(string source);

        /// <summary> COM interface member for <see cref="Admin.GetLogs(string, string)"/> </summary>
        [DispId(4)]
        object GetLogs(string source, string message);

        /// <summary> COM interface for <see cref="Admin.privilegesAreElevated"/> </summary>
        [DispId(5)]
        bool privilegesAreElevated { get; }

        /// <summary> COM interface for <see cref="Admin.log(string)"/> </summary>
        [DispId(6)]
        void log(string msg);

        /// <summary> COM interface for <see cref="Admin.Result"/> </summary>
        [DispId(7)]
        EventLogResultT Result { get; }
    }
}
