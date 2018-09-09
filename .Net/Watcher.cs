using System.Runtime.InteropServices;
using System;
using System.Timers;

namespace VBScripting
{
    /// <summary> Provides something like presentation mode for Windows systems that don't have presentation.exe. </summary>
    /// <remarks> Uses <a href="https://msdn.microsoft.com/en-us/library/aa373208(v=vs.85).aspx"> SetThreadExecutionState</a>. Adapted from <a href="https://stackoverflow.com/questions/6302185/how-to-prevent-windows-from-entering-idle-state"> stackoverflow.com</a> and <a href="http://www.pinvoke.net/default.aspx/kernel32.setthreadexecutionstate"> pinvoke.net</a> posts. </remarks>
    [ProgId("VBScripting.Watcher"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-C000-495F-AB4D-6C61BD463EA4")]
    public class Watcher : IWatcher
    {
        private int _resetPeriod;
        private byte _currentState;
        private bool _watch;
        private System.Timers.Timer timer;

        /// <summary> Constructor. Starts a private timer that periodically resets the system idle timer with the desired state. </summary>
        public Watcher()
        {
            ResetPeriod = 30000;
            Watch = false;
        }
        /// <summary> Gets or sets whether the system and monitor(s) should be kept from going into a suspend (sleep) state. The computer may still be put to sleep by other applications or by user actions such as closing a laptop lid or pressing a sleep button or power button. Default is False. </summary>
        public bool Watch
        {
            get { return _watch; }
            set
            {
                _watch = value;
                if (value)
                {
                    CurrentState = 3;
                }
                else
                {
                    CurrentState = 0;
                }
            }
        }
        /// <summary> Gets or sets an integer describing the current thread execution state. Intended for internal use and testing only.</summary>
        public byte CurrentState
        {
            get { return _currentState; }
            set
            {
                _currentState = value;
                SetThreadExecutionState((uint)(value + 0x80000000));
            }
        }
        // Called periodically to refresh the current state
        private void ResetIdleTimer(Object source, ElapsedEventArgs e)
        {
            CurrentState = CurrentState;
        }
        /// <summary> Disposes of the object's resources. </summary>
        public void Dispose()
        {
            if (timer != null)
            {
                timer.Stop();
                timer.Dispose();
            }
        }
        /// <summary> Gets or sets the time in milliseconds between idle-timer resets. Optional. Default is 30000. Max 2147483647.</summary>
        // Set also initializes/resets the internal timer.
        public int ResetPeriod
        {
            get { return _resetPeriod; }
            set
            {
                _resetPeriod = value;
                if (timer != null)
                {
                    timer.Dispose();
                }
                timer = new System.Timers.Timer(value);
                timer.Elapsed += ResetIdleTimer;
                timer.Start();
            }
        }
        /// <summary> Turn off the monitor(s). </summary>
        public void MonitorOff()
        {
            Admin.MonitorOff();
        }
        /// <summary> Gets a boolean indicating whether privileges are elevated. </summary>
        public bool Privileged
        {
            get { return Admin.PrivilegesAreElevated; }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        extern static uint SetThreadExecutionState(uint esFlags);
    }
   
    /// <summary> The COM interface for VBScripting.Watcher. </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-C001-495F-AB4D-6C61BD463EA4")]
    public interface IWatcher
    {
        /// <summary> </summary>
        [DispId(0)]
        int ResetPeriod { get; set; }
        /// <summary> </summary>
        [DispId(1)]
        bool Watch { get; set; }
        /// <summary> </summary>
        [DispId(3)]
        byte CurrentState { get; set; }
        /// <summary> </summary>
        [DispId(5)]
        void Dispose();
        /// <summary> </summary>
        [DispId(6)]
        void MonitorOff();
        /// <summary> </summary>
        [DispId(7)]
        bool Privileged { get; }
    }
}

