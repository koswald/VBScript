using System.Diagnostics;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System;
using System.Linq;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;

namespace VBScripting
{
    /// <summary> Provide miscellaneous system admin. features. </summary>
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

        /// <summary> </summary>
        // VBScript wrapper for the static PrivilegesAreElevated property. </summary>
        public bool privilegesAreElevated
        {
            get { return PrivilegesAreElevated; }
        }

        /// <summary> Gets whether the current user is in the Administrator group (on the current machine). Does not necessarily mean that privileges are elevated. Adapted from a <a href="https://stackoverflow.com/questions/44507149/how-to-check-if-current-user-is-in-admin-group-c-sharp#answer-47564106" title="stackoverflow.com" target="_blank"> stackoverflow.com post</a>.</summary>
        // requires using statement: System.DirectoryServices.AccountManagement
        // requires assembly reference: C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.DirectoryServices.AccountManagement.dll
        public static bool IsAdministrator()
        {
            using (var pc = new PrincipalContext(ContextType.Machine, Environment.UserDomainName))
            {
                using (var up = UserPrincipal.FindByIdentity(pc, WindowsIdentity.GetCurrent().Name))
                {
                    return up.GetAuthorizationGroups().Any(group => group.Name == "Administrators");
                }
            }
        }
        /// <summary> </summary>
        // VBScript wrapper for the static IsAdministrator
        public bool isAdministrator()
        {
            return IsAdministrator();
        }

        # region EventLogs

        private const string eventSource = "VBScripting";
        private const string logName = "Application";

        /// <summary> Gets the name of the event log source for this namespace (VBScripting). </summary>
        /// <returns> a string </returns>
        public string EventSource { get { return eventSource; } }
        /// <summary> Gets the name of the log to which events will be logged. </summary>
        /// <returns> a string </returns>
        public string LogName { get { return logName; } }

        /// <summary> Logs the specified message to the event log. </summary>
        /// <parameters> message </parameters>
        public static void Log(string msg)
        {
            try
            {
                new EventLog(logName, ".", eventSource).WriteEntry(msg);
            }
            catch { }
        }

        /// <summary> </summary>
        // <summary> VBScript wrapper for the static Log method. </summary>
        public void log(string msg)
        {
            Log(msg);
        }

        /// <summary> Get an array of logs entries from the Application log. </summary>
        /// <parameters> source, message </parameters>
        /// <returns> an array </returns>
        /// <remarks> Returns an array of logs (strings) from the specified event source that contain the specified message string. Searches the Application log only. </remarks>
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

        /// <summary> Gets whether the specified EventLog source exists. </summary>
        /// <returns> a boolean </returns>
        /// <parameters> source </parameters>
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

        /// <summary> Creates the specified EventLog source. </summary>
        /// <parameters> source </parameters>
        /// <returns> an EventLogSourceResult </returns>
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
        /// <summary> Deletes the specified EventLog source and all of its logs. </summary>
        /// <parameters> source </parameters>
        /// <returns> an EventLogSourceResult </returns>
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
        /// <summary> Gets an EventLogResultT object. </summary>
        /// <returns> an EventLogResultT </returns>
        /// <remarks> VBScript example: <pre> Set returnValue = adm.CreateEventSource <br /> If returnValue.Result = adm.Result.SourceCreationException Then <br />     MsgBox returnValue.Message <br /> End If</pre></remarks>
        public EventLogResultT Result
        {
            get { return new EventLogResultT(); }
        }
        # endregion EventLogs

        /// <summary> </summary>
        // VBScript wrapper for the static MonitorOff()
        public void monitorOff()
        {
            MonitorOff();
        }
        /// <summary> Turn off the monitor(s) </summary>
        public static void MonitorOff()
        {
            int HWND = -1;
            int WM_SYSCOMMAND = 0x0112;
            int SC_MONITORPOWER = 0xF170;
            //int MONITOR_ON = -1;
            int MONITOR_OFF = 2;
            SendMessage(HWND, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF);
        }

        [DllImport("user32.dll")]
        private static extern int SendMessage(int hWnd, int hMsg, int wParam, int lParam);
    }
    /// <summary> Provides a set of terse behavior/result descriptions suitable for VBScript comparisons and MsgBox captions. </summary>
    /// <remarks> Not directly available to VBScript. See <tt>Admin.Result</tt>. </remarks>
    [ProgId("VBScripting.EventLogResultT"),
        Guid("2650C2AB-8CF8-495F-AB4D-6C61BD463EA4")]
    public class EventLogResultT
    {
        /// <returns> "Source already exists" </returns>
        public string SourceAlreadyExists {
            get { return "Source already exists"; } } // attempting to create
        /// <returns> "Source created" </returns>
        public string SourceCreated {
            get { return "Source created"; } }
        /// <returns> "Source creation error" </returns>
        public string SourceCreationException {
            get { return "Source creation error"; } }
        /// <returns> "Source does not exist" </returns>
        public string SourceDoesNotExist {
            get { return "Source does not exist"; } } // attempting to delete
        /// <returns> "Source deleted" </returns>
        public string SourceDeleted {
            get { return "Source deleted"; } }
        /// <returns> "Source deletion error" </returns>
        public string SourceDeletionException {
            get { return "Source deletion error"; } }
    }
    /// <summary> Type returned by CreateEventSource and DeleteEventSource. </summary>
    [ProgId("VBScripting.EventLogSourceResult"),
        Guid("2650C2AB-8DF8-495F-AB4D-6C61BD463EA4")]
    public class EventLogSourceResult
    {
        /// <summary> Returns True if the source exists after the attempted operation has completed. </summary>
        /// <returns> a boolean </returns>
        public bool SourceExists { get; set; }
        /// <summary> Returns a message descriptive of the outcome of the operation. </summary>
        /// <returns> a string </returns>
        public string Message { get; set; }
        /// <summary> Returns a string: one of the EventLogResultT strings. </summary>
        /// <returns> a string </returns>
        public string Result { get; set; }
    }

    /// <summary> COM interface for VBScripting.Admin </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-8BF8-495F-AB4D-6C61BD463EA4")]
    public interface IAdmin
    {
        /// <summary> </summary>
        [DispId(1)]
        EventLogSourceResult CreateEventSource(string source);

        /// <summary> </summary>
        [DispId(2)]
        EventLogSourceResult DeleteEventSource(string source);

        /// <summary> </summary>
        [DispId(3)]
        bool SourceExists(string source);

        /// <summary> </summary>
        [DispId(4)]
        object GetLogs(string source, string message);

        /// <summary> </summary>
        [DispId(5)]
        bool privilegesAreElevated { get; }

        /// <summary> </summary>
        [DispId(6)]
        void log(string msg);

        /// <summary> </summary>
        [DispId(7)]
        EventLogResultT Result { get; }

        /// <summary> </summary>
        [DispId(8)]
        string EventSource { get; }

        /// <summary> </summary>
        [DispId(9)]
        string LogName { get; }

        /// <summary> </summary>
        [DispId(10)]
        void monitorOff();

        /// <summary> </summary>
        [DispId(11)]
        bool isAdministrator();
    }
}
