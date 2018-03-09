using System.Runtime.InteropServices;
using System;
using System.Timers;

namespace VBScripting
{
    /// <summary> Provides something like presentation mode for Windows systems that don't have presentation.exe. </summary>
    /// <remarks> Uses <a href="https://msdn.microsoft.com/en-us/library/aa373208(v=vs.85).aspx"> SetThreadExecutionState</a>. Adapted from <a href="https://stackoverflow.com/questions/6302185/how-to-prevent-windows-from-entering-idle-state"> stackoverflow.com</a> and <a href="http://www.pinvoke.net/default.aspx/kernel32.setthreadexecutionstate"> pinvoke.net</a> posts. </remarks>
    [ProgId("VBScripting.IdleTimer"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-C000-495F-AB4D-6C61BD463EA4")]
    public class IdleTimer : IIdleTimer
    {
        private static uint _resetPeriod;
        private static System.Timers.Timer timer;

        /// <summary> Constructor </summary>
        public IdleTimer()
        {
            LogOps = false;
            ResetPeriod = 30000;
            InitialState = SetThreadExecutionState(ES_CONTINUOUS);
            SetThreadExecutionState(InitialState);
            PreventSleepState = ES_DISPLAY_REQUIRED | ES_CONTINUOUS | ES_SYSTEM_REQUIRED;
            AllowSleepState = ES_CONTINUOUS;
        }

        /// <summary> Tends to prevent the system from entering a suspend (sleep) state or hibernation. </summary>
        /// <remarks> Other applications or direct user action may still cause the computer to sleep or hibernate. Uses a private <em> reset</em> timer to periodically reset the system idle timer. By default, also prevents the monitor from powering down; this can be changed by setting PreventSleepState to &amp;h80000001 before calling PreventSleep. </remarks>
        public void PreventSleep()
        {
            InitializeTimer();
            timer.Start();
            if (LogOps)
                Admin.Log("VBScripting.IdleTimer.PreventSleep: started the reset timer.");
        }
        /// <summary> Allows the computer to go into a sleep state. Reverses the effect of the PreventSleep method. </summary>
        public void AllowSleep()
        {
            timer.Stop();
            SetThreadExecutionState(AllowSleepState);
            if (LogOps)
                Admin.Log("VBScripting.IdleTimer.AllowSleep: stopped the reset timer.");
        }
        /// <summary> Gets or sets whether operations are logged to the Application event log. </summary>
        /// <remarks> Default is False. </remarks>
        public static bool LogOps { get; set; }

        /// <summary> </summary>
        // VBScript wrapper for the static LogOps
        public bool logOps
        {
            get { return LogOps; }
            set { LogOps = value; }
        }
        private static void ResetIdleTimer(Object source, ElapsedEventArgs e)
        {
            SetThreadExecutionState(PreventSleepState);
            if (LogOps)
                Admin.Log("VBScripting.IdleTimer.ResetIdleTimer: Refreshed the execution state.");
        }
        /// <summary> Disposes of the object's resources. </summary>
        public void Dispose()
        {
            timer.Dispose();
            SetThreadExecutionState(InitialState);
        }
        /// <summary> Gets or sets the time in milliseconds between idle-timer resets. Optional. Default is 30,000. </summary>
        public uint ResetPeriod
        {
            get { return _resetPeriod; }
            set
            {
                _resetPeriod = value;
                InitializeTimer();
            }
        }
        private void InitializeTimer()
        {
            timer = new System.Timers.Timer(ResetPeriod);
            timer.Elapsed += ResetIdleTimer;
        }
        /// <summary> Gets the initial state. </summary>
        public static uint InitialState { get; private set; }
        /// <summary> </summary>
        // VBScript wrapper for static property
        public uint initialState { get { return InitialState; } }
        /// <summary> Gets or sets the state for preventing sleep. Default is &amp;h80000003. </summary>
        public static uint PreventSleepState { get; set; }
        /// <summary> </summary>
        // VBScript wrapper for static property
        public uint preventSleepState { get { return PreventSleepState; } set { PreventSleepState = value; } }
        /// <summary> Gets or sets the state for allowing sleep. Default is &amp;h80000000. </summary>
        public static uint AllowSleepState { get; set; }
        /// <summary> </summary>
        // VBScript wrapper for static property
        public uint allowSleepState { get { return AllowSleepState; } set { AllowSleepState = value; } }

        /// <summary> Typically not required or recommended. See <a href="https://msdn.microsoft.com/en-us/library/aa373208(v=vs.85).aspx"> SetThreadExecutionState</a>. </summary>
        /// <returns> &amp;h00000040 </returns>
        public uint ES_AWAYMODE_REQUIRED { get { return 0x00000040; } }
        /// <returns> &amp;h80000000 </returns>
        public uint ES_CONTINUOUS { get { return 0x80000000; } }
        /// <returns> &amp;h00000002 </returns>
        public uint ES_DISPLAY_REQUIRED { get { return 0x00000002; } }
        /// <returns> &amp;h00000001 </returns>
        public uint ES_SYSTEM_REQUIRED { get { return 0x00000001; } }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern uint SetThreadExecutionState(uint esFlags);
    }
   
    /// <summary> The COM interface for VBScripting.IdleTimer. </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-C001-495F-AB4D-6C61BD463EA4")]
    public interface IIdleTimer
    {
        /// <summary> </summary>
        [DispId(0)]
        void AllowSleep();

        /// <summary> </summary>
        [DispId(1)]
        void PreventSleep();

        /// <summary> </summary>
        [DispId(2)]
        void Dispose();

        /// <summary> </summary>
        [DispId(3)]
        bool logOps { get; set; }

        /// <summary> </summary>
        [DispId(4)]
        uint ResetPeriod { get; set; }

        /// <summary> </summary>
        [DispId(5)]
        uint initialState { get; }

        /// <summary> </summary>
        [DispId(6)]
        uint preventSleepState { get; set; }
        /// <summary> </summary>
        [DispId(7)]
        uint allowSleepState { get; set; }
        /// <summary> </summary>
        [DispId(8)]
        uint ES_AWAYMODE_REQUIRED { get; }
        /// <summary> </summary>
        [DispId(9)]
        uint ES_CONTINUOUS { get; }
        /// <summary> </summary>
        [DispId(10)]
        uint ES_DISPLAY_REQUIRED { get; }
        /// <summary> </summary>
        [DispId(11)]
        uint ES_SYSTEM_REQUIRED { get; }
    }
}

