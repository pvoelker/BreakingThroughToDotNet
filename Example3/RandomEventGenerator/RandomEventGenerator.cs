// Breaking Through to .NET - Example 3 - COM interop event mechanism
// March 2014
// Paul Voelker

using System;
using System.Text;
using System.Timers;

namespace RandomEventGenerator
{
    public delegate void RandomTimerEventHandler(object sender);

    // .NET class with delegate based events
    public class RandomEventGenerator
    {
        public event RandomTimerEventHandler TimerFired;

        private Random random = new Random();
        private System.Timers.Timer timer = new System.Timers.Timer();

        protected virtual void OnTimerEvent(bool fireEvent)
        {
            if (fireEvent == true)
            {
                if (TimerFired != null)
                    TimerFired(this);
            }

            // Randomize timer (maximum 2 seconds)
            timer.Interval = random.Next(2000);
            timer.Enabled = true;
        }

        public RandomEventGenerator()
        {
            timer.AutoReset = false;
            timer.Elapsed += timer_Elapsed;
            OnTimerEvent(false);
        }

        private void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            OnTimerEvent(true);
        }
    }
}
