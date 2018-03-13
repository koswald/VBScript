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
        private double _resetPeriod;
        private uint _currentState;
        private const double T32 = 0x100000000; // converting large hex values to & from VBScript
        private System.Timers.Timer timer;

        /// <summary> Constructor. Starts a private timer that periodically resets the system idle timer with the desired state. </summary>
        public IdleTimer()
        {
            DesiredState = 0x80000000;
            CurrentState = DesiredState;
            SystemRequired = false;
            DisplayRequired = false;
            ResetPeriod = 30000;
        }
        /// <summary> Gets or sets whether the system should be kept from going into a suspend (sleep) state or hibernate. Default is False. </summary>
        public bool SystemRequired
        {
            get { return (DesiredState & 0x0000001) > 0; }
            set
            {
                if (value && DisplayRequired)
                {
                    DesiredState = 0x80000003;
                }
                else if (value)
                {
                    DesiredState = 0x80000001;
                }
                else if (DisplayRequired)
                {
                    DesiredState = 0x80000002;
                }
                else
                {
                    DesiredState = 0x80000000;
                }
                CurrentState = DesiredState;
            }
        }
        /// <summary> Gets or sets whether the monitor should be kept awake. Default is False. </summary>
        public bool DisplayRequired
        {
            get { return (DesiredState & 0x00000002) > 0; }
            set
            {
                if (value && SystemRequired)
                {
                    DesiredState = 0x80000003;
                }
                else if (value)
                {
                    DesiredState = 0x80000002;
                }
                else if (SystemRequired)
                {
                    DesiredState = 0x80000001;
                }
                else
                {
                    DesiredState = 0x80000000;
                }
                CurrentState = DesiredState;
            }
        }
        /// <summary> Gets or sets a double describing the current thread execution state. </summary>
        public double currentState
        {
            get { return (double)(CurrentState - T32); }
            set { CurrentState = (uint)(value + T32); }
        }
        private uint CurrentState
        {
            get { return _currentState; }
            set
            {
                SetThreadExecutionState(value);
                _currentState = value;
            }
        }
        /// <summary> </summary>
        // wraps DesiredState; returns a double to VBScript
        // undocumented property made public for testability
        public double desiredState
        {
            get { return (double)(DesiredState - T32); }
            set { DesiredState = (uint)(value + T32); }
        }
        private uint DesiredState { get; set; }

        // called periodically to refresh the current state
        private void ResetIdleTimer(Object source, ElapsedEventArgs e)
        {
            CurrentState = DesiredState;
        }
        /// <summary> Disposes of the object's resources. </summary>
        public void Dispose()
        {
            timer.Stop();
            timer.Dispose();
        }
        /// <summary> Gets or sets the time in milliseconds between idle-timer resets. Optional. Default is 30,000. </summary>
        public double resetPeriod
        {
            get { return ResetPeriod; }
            set { ResetPeriod = value; }
        }
        private double ResetPeriod
        {
            get { return _resetPeriod; }
            set
            {
                _resetPeriod = value;
                timer = new System.Timers.Timer(ResetPeriod);
                timer.Elapsed += ResetIdleTimer;
                timer.Start();
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        extern static uint SetThreadExecutionState(uint esFlags);
    }
   
    /// <summary> The COM interface for VBScripting.IdleTimer. </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-C001-495F-AB4D-6C61BD463EA4")]
    public interface IIdleTimer
    {
        /// <summary> </summary>
        [DispId(0)]
        double resetPeriod { get; set; }
        /// <summary> </summary>
        [DispId(1)]
        bool SystemRequired { get; set; }
        /// <summary> </summary>
        [DispId(2)]
        bool DisplayRequired { get; set; }
        /// <summary> </summary>
        [DispId(3)]
        double currentState { get; set; }
        /// <summary> </summary>
        [DispId(4)]
        double desiredState { get; set; }
        /// <summary> </summary>
        [DispId(5)]
        void Dispose();
    }
}

