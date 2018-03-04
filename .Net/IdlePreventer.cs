using System.Runtime.InteropServices;
using System;
using System.Timers;

namespace VBScripting
{
    /// <summary> Provides something like presentation mode for non-Pro Windows systems, which don't have presentation.exe. </summary>
    /// <remarks> Adapted from a <a href="https://stackoverflow.com/questions/6302185/how-to-prevent-windows-from-entering-idle-state"> stackoverflow post</a> and a <a href="http://www.pinvoke.net/default.aspx/kernel32.setthreadexecutionstate"> pinvoke.net post</a>. See the <a href="https://msdn.microsoft.com/en-us/library/aa373208(v=vs.85).aspx"> SetThreadExecutionState docs</a>. </remarks>
    [ProgId("VBScripting.IdlePreventer"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-C000-495F-AB4D-6C61BD463EA4")]
    public class IdlePreventer : IIdlePreventer
    {
        private uint _refreshPeriod;
        private static System.Timers.Timer timer;

        /// <summary> Constructor </summary>
        public IdlePreventer()
        {
            LogOps = false;
            RefreshPeriod = 30000;
        }

        /// <summary> Prevents the system from going into an idle state. </summary>
        /// <remarks> Also prevents the monitor from powering down. </remarks>
        public void PreventIdle()
        {
            timer.Start();
            if (LogOps)
                Admin.Log("VBScripting.IdlePreventer.PreventIdle: started the timer.");
        }

        /// <summary> Allows the computer to go into an idle state after having prevented it with the PreventIdle method. </summary>
        public void AllowIdle()
        {
            timer.Stop();
            SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS);
            if (LogOps)
                Admin.Log("VBScripting.IdlePreventer.AllowIdle: stopped the timer.");
        }

        /// <summary> Gets or sets whether operations are logged to the event log. </summary>
        /// <remarks> Default is False. </remarks>
        public static bool LogOps { get; set; }

        /// <summary> </summary>
        // VBScript wrapper for the static LogOps
        public bool logOps
        {
            get { return LogOps; }
            set { LogOps = value; }
        }

        private static void PreventIdleRefresh(Object source, ElapsedEventArgs e)
        {
            SetThreadExecutionState(EXECUTION_STATE.ES_DISPLAY_REQUIRED | EXECUTION_STATE.ES_CONTINUOUS | EXECUTION_STATE.ES_SYSTEM_REQUIRED);
            if (LogOps)
                Admin.Log("VBScripting.IdlePreventer.PreventIdleRefresh: Refreshed the execution state.");
        }

        /// <summary> Disposes of the object's resources. </summary>
        public void Dispose()
        {
            timer.Dispose();
        }

        /// <summary> Gets or sets the time in milliseconds between refreshes of the PreventIdle setting. Default is 30,000. </summary>
        public uint RefreshPeriod
        {
            get { return _refreshPeriod; }
            set
            {
                _refreshPeriod = value;
                timer = new System.Timers.Timer(value);
                timer.Elapsed += PreventIdleRefresh;
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
        static extern EXECUTION_STATE SetThreadExecutionState(EXECUTION_STATE esFlags);

        /// <summary> </summary>
        [FlagsAttribute]
        public enum EXECUTION_STATE : uint
        {
            /// <summary> </summary>
            ES_AWAYMODE_REQUIRED = 0x00000040,
            /// <summary> </summary>
            ES_CONTINUOUS = 0x80000000,
            /// <summary> </summary>
            ES_DISPLAY_REQUIRED = 0x00000002,
            /// <summary> </summary>
            ES_SYSTEM_REQUIRED = 0x00000001
            // Legacy flag, should not be used.
            // ES_USER_PRESENT = 0x00000004
        }
    }

    /// <summary> The COM interface for IdlePreventer. </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-C001-495F-AB4D-6C61BD463EA4")]
    public interface IIdlePreventer
    {

        /// <summary> </summary>
        [DispId(0)]
        void AllowIdle();

        /// <summary> </summary>
        [DispId(1)]
        void PreventIdle();

        /// <summary> </summary>
        [DispId(2)]
        void Dispose();

        /// <summary> </summary>
        [DispId(3)]
        bool logOps { get; set; }

        /// <summary> </summary>
        [DispId(4)]
        uint RefreshPeriod { get; set; }
    }

}

