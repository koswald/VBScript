
// Miscellaneous admin functions

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
            get { return PrivilegesAreElevated; }
        }

        # region EventLogs

        /// <summary> Name of the event log source for this namespace. </summary>
        private const string eventSource = "VBScripting";
        /// <summary> Name of the log to which events will be logged. </summary>
        private const string logName = "Application";

        /// <summary> VBScript wrapper for <see cref="eventSource"/> </summary>
        public string EventSource { get { return eventSource; } }
        /// <summary> VBScript wrapper for <see cref="logName"/> </summary>
        public string LogName { get { return logName; } }

        /// <summary> Logs an event to the event log. </summary>
        public static void Log(string msg)
        {
            try
            {
                new EventLog(logName, ".", eventSource).WriteEntry(msg);
            }
            catch { }
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
        public EventLogSourceResult CreateEventSource(string source)
        {
            string msg;

            // check if event source exists
            if (this.SourceExists(source))
            {
                msg = string.Format(
                    "The source \"{0}\" already exists.",
                    source
                );
                return new EventLogSourceResult
                {
                    SourceExists = true,
                    Message = msg,
                    Result = this.Result.SourceAlreadyExists
                };
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
                    return new EventLogSourceResult
                    {
                        SourceExists = true,
                        Message = msg,
                        Result = this.Result.SourceCreated
                    };
                }
                catch (Exception e)
                {
                    msg = string.Format(
                        "Failed to create source \"{0}\".\n\n{1}",
                        source, e.ToString()
                    );
                    Log(msg);
                    return new EventLogSourceResult
                    {
                        SourceExists = false,
                        Message = msg,
                        Result = this.Result.SourceCreationException
                    };
                }
            }
        }
        /// <summary> Delete an EventLog source and all of its logs. </summary>
        /// <param name="source"></param>
        public EventLogSourceResult DeleteEventSource(string source)
        {
            string msg;

            // check if event source exists, and check privileges
            if (!this.SourceExists(source))
            {
                msg = string.Format(
                    "The source \"{0}\" does not exist.",
                    source
                );
                return new EventLogSourceResult
                {
                    SourceExists = false,
                    Message = msg,
                    Result = this.Result.SourceDoesNotExist
                };
            }
            // delete the source
            try
            {
                EventLog.DeleteEventSource(source);
                msg = string.Format(
                    "The EventLog source \"{0}\" has been deleted.",
                    source
                );
                return new EventLogSourceResult
                {
                    SourceExists = false,
                    Message = msg,
                    Result = this.Result.SourceDeleted
                };
            }
            catch (Exception e)
            {
                msg = string.Format(
                    "Failed to delete source \"{0}\".\n" +
                    "Privileges {1} elevated.\n" + 
                    "Elevated privileges are required.\n\n{2}",
                    source,
                    PrivilegesAreElevated? "are" : "are not",
                    e.Message
                );
                Log(msg);
                return new EventLogSourceResult
                {
                    SourceExists = true,
                    Message = msg,
                    Result = this.Result.SourceDeletionException
                };
            }
        }
        /// <summary> Returns the behavior description collection object to VBScript as a property. </summary>
        public EventLogResultT Result
        {
            get { return new EventLogResultT(); }
        }
        # endregion EventLogs
    }
    /// <summary> A set of terse behavior/result descriptions suitable for VBScript MsgBox captions. </summary>
    public class EventLogResultT
    {
        /// <summary>  </summary>
        public string SourceFound {
            get { return "Source found"; } } // when checking existence
        /// <summary>  </summary>
        public string SourceNotFoundLowPrivileges {
            get { return "Source not found; priv. not elevated"; } } // Can't determine whether source exists
        /// <summary>  </summary>
        public string SourceNotFoundHighPrivileges {
            get { return "Source not found; priv. elevated"; } } // source doesn't exist
        /// <summary>  </summary>
        public string SourceAlreadyExists {
            get { return "Source already exists"; } } // when attempting to create
        /// <summary>  </summary>
        public string SourceCreated {
            get { return "Source created"; } }
        /// <summary>  </summary>
        public string SourceNotCreated {
            get { return "Source not created"; } }
        /// <summary>  </summary>
        public string SourceCreationException {
            get { return "Source creation error"; } }
        /// <summary>  </summary>
        public string SourceDoesNotExist {
            get { return "Source does not exist"; } } // attempting to delete
        /// <summary>  </summary>
        public string SourceDeleted {
            get { return "Source deleted"; } }
        /// <summary>  </summary>
        public string SourceFoundInAnotherLog {
            get { return "Source found in another log"; } }
        /// <summary>  </summary>
        public string SourceNotDeleted {
            get { return "Source not deleted"; } }
        /// <summary>  </summary>
        public string SourceDeletionException {
            get { return "Source deletion error"; }
        }
    }
    /// <summary> Result for returning to VBScript. </summary>
    public class EventLogSourceResult
    {
        /// <summary>  </summary>
        public bool SourceExists { get; set; }
        /// <summary>  </summary>
        public string Message { get; set; }
        /// <summary>  </summary>
        public string Result { get; set; }
    }

    /// <summary> COM interface for <see cref="Admin"/> </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-8BF8-495F-AB4D-6C61BD463EA4")]
    public interface IAdmin
    {
        /// <summary> COM interface member for <see cref="Admin.CreateEventSource(string)"/> </summary>
        [DispId(1)]
        EventLogSourceResult CreateEventSource(string source);

        /// <summary> COM interface member for <see cref="Admin.DeleteEventSource(string)"/> </summary>
        [DispId(2)]
        EventLogSourceResult DeleteEventSource(string source);

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

        /// <summary> COM interface for <see cref="Admin.EventSource"/> </summary>
        [DispId(8)]
        string EventSource { get; }

        /// <summary> COM interface for <see cref="Admin.LogName"/> </summary>
        [DispId(9)]
        string LogName { get; }

    }
}
