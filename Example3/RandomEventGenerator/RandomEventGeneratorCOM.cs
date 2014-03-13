// Breaking Through to .NET - Example 3 - COM interop event mechanism
// March 2014
// Paul Voelker

// NOTES:
// --- This example demonstrates a way to convert .NET events to COM
//     callback-based events.
// --- 'Register for COM interop' is enabled under build settings.  This will
//     create a type library (TLB) file in the output folder.
// --- There is a post build event to copy the type library (TLB) file
//     to a common location so the test application can use it readily.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace RandomEventGenerator
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("1EBE2441-23AE-47D2-B820-6CE9DFFD4C47")]
    public interface IRandomEventGeneratorCOM
    {
        void Subscribe(IRandomEventCOMEvents callback);
        void Unsubscribe();
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("44228556-77BF-4316-B14C-E8477CBADCC9")]
    public interface IRandomEventCOMEvents
    {
        // Definition for COM callback interface.  Interface needs to be implemented
        // by whoever is using IRandomEventGeneratorCOM.
        void TimerFired();
    }

    // This is a wrapper class to convert .NET events to COM callbacks
    [ComVisible(true)]
    [ComDefaultInterface(typeof(IRandomEventGeneratorCOM))]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("DD7DA936-7355-453F-8FEC-EDDDB334AD80")]
    public class RandomEventGeneratorCOM : IRandomEventGeneratorCOM
    {
        private RandomEventGenerator m_generator = new RandomEventGenerator();

        private IRandomEventCOMEvents m_callback;

        public RandomEventGeneratorCOM()
        {
            m_generator.TimerFired += generator_TimerFired;
        }

        public void Subscribe(IRandomEventCOMEvents callback)
        {
            if (callback == null)
                throw new ArgumentNullException("callback");

            // Note: Make sure multi-threaded concurrency is set.  Because timer
            //       events can execute on different threads, the COM callback
            //       calls will freeze if multi-threaded concurrency is NOT set.
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.MTA)
                throw new Exception("Multithreaded object concurrency required");

            if (m_callback == null)
            {
                m_callback = callback;
            }
            else
            {
                throw new Exception("Callback is already specified");
            }
        }

        public void Unsubscribe()
        {
            m_callback = null;
        }

        private void generator_TimerFired(object sender)
        {
            // Note: Timers can fire on a different thread.  COM callbacks
            //       must be done on thread the parent COM object was created.

            if (m_callback != null)
            {
                m_callback.TimerFired();
            }
        }
    }
}
