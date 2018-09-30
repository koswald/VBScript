using System.Runtime.InteropServices;
using System.Timers;
using System;

namespace VBScripting
{
    /// <summary> Wraps the <a href="https://docs.microsoft.com/en-us/dotnet/api/system.timers.timer?view=netframework-4.7.1" title="docs.microsoft.com"> System.Timers.Timer class</a> for VBScript. </summary>
    [ProgId("VBScripting.Timer"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-C020-495F-AB4D-6C61BD463EA4")]
    public class Timer : ITimer
    {
        private System.Timers.Timer timer;

        /// <summary> Gets or sets the number of milliseconds between when the Start method is called and when the callback is invoked. Default is 100. Max is 2,147,483,647 milliseconds, or 24 days 20 hours 31 minutes 23.647 seconds.</summary>
        public int Interval { get; set; }
        /// <summary> Gets or sets a reference to the VBScript Sub that is called when the interval has elapsed. </summary>
        public object Callback { get; set; }
        /// <summary> Gets or sets a boolean determining whether to repeatedly invoke the callback. Default is False. If False, the callback is invoked only once, until the timer is restarted with the Start method. </summary>
        public bool AutoReset { get; set; }

        /// <summary> Constructor </summary>
        public Timer()
        {
            Interval = 60000;
        }
        /// <summary> Starts or restarts the timer. </summary>
        public void Start()
        {
            if (timer != null)
                timer.Stop();
            
            timer = new System.Timers.Timer();
            timer.Interval = Interval;
            timer.AutoReset = AutoReset;
            timer.Elapsed += Elapsed;
            timer.Start();
        }
        private void Elapsed(object sender, ElapsedEventArgs e)
        {
            if (Callback != null)
                ComEvent.InvokeComCallback(Callback);
        }
        /// <summary> Stops the timer. </summary>
        public void Stop()
        {
            Dispose();
        }
        /// <summary> Disposes of the timer's resources. </summary>
        public void Dispose()
        {
            if (timer != null)
            {
                timer.Stop();
                timer.Dispose();
            }
        }
        /// <summary> Gets or sets the interval in hours. </summary>
        public float IntervalInHours
        {
            get { return (float) Interval/3600000; }
            set { Interval = Convert.ToInt32(value*3600000); }
        }
    }
    /// <summary> COM interface for VBScripting.Timer. </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-C021-495F-AB4D-6C61BD463EA4")]
    public interface ITimer
    {
        /// <summary> </summary>
        [DispId(0)]
        int Interval { get; set; }
        /// <summary> </summary>
        [DispId(1)]
        object Callback { get; set; }
        /// <summary> </summary>
        [DispId(2)]
        bool AutoReset { get; set; }
        /// <summary> </summary>
        [DispId(3)]
        void Start();
        /// <summary> </summary>
        [DispId(4)]
        void Stop();
        /// <summary> </summary>
        [DispId(5)]
        void Dispose ();
        /// <summary> </summary>
        [DispId(6)]
        float IntervalInHours { get; set; }
    }
}