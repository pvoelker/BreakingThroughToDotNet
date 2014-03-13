// Breaking Through to .NET - Example 4 - COM interop ActiveX wrapper for WinForms user control
// March 2014
// Paul Voelker

// NOTES:
// --- For ActiveX controls, IDispatch must be used instead of IUnknown.
// --- 'Register for COM interop' is enabled under build settings.  This will
//     create a type library (TLB) file in the output folder.
// --- There is a post build event to copy the type library (TLB) file
//     to a common location so the test application can use it readily.
// --- 'KeypadControl.rc' is used to include a Win32 bitmap resource for the
//      ActiveX 'ToolboxBitmap32' registry entry.
// --- The 'rc' is compiled to a 'res' file using a pre-build event.
// --- The 'res' file is included using the 'Application' project settings.
// --- Because we are using our own 'res' file during the build, the version
//     information for the compiled DLL will come from the 'rc' file.
//     However, the assembly version information still needs to be maintained.
//     Make sure the version information in the assembly information and 'rc'
//     file are synced up with each other.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace KeypadWinFormLib
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)] // ActiveX uses IDispatch interfaces vs IUnknown
    [Guid("D7FD181F-B1DE-471F-9DA4-AE93CA96E8D6")]
    public interface IKeypadControl
    {
        [DispId(1)] // Optional, but helps with backwards compatability.  DispId are used to indentify functions with IDispatch interfaces
        String CurrentValue { get; }

        [DispId(2)] // Optional, but helps with backwards compatability.  DispId are used to indentify functions with IDispatch interfaces
        void ClearValue();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("62887F0D-8AAF-4594-BE0D-438F25320BA2")]
    public partial class KeypadControl: UserControl, IKeypadControl
    {
        public KeypadControl()
        {
            InitializeComponent();
        }

        private String m_CurrentValue;

        public String CurrentValue
        {
            get { return m_CurrentValue; }
        }

        public void ClearValue()
        {
            m_CurrentValue = "";
            Debug.WriteLine("Keypad Value Cleared");
        }

        [ComRegisterFunction()]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public static void RegisterClass(Type t)
        {
            ActiveXHelpers.ActiveXHelpers.RegisterClass(t, 101, 1, 0);
        }

        [ComUnregisterFunction()]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public static void UnregisterClass(Type t)
        {
            ActiveXHelpers.ActiveXHelpers.UnregisterClass(t);
        }

        private void btnNum1_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 1 Pressed");
            m_CurrentValue += "1";
        }

        private void btnNum2_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 2 Pressed");
            m_CurrentValue += "2";
        }

        private void btnNum3_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 3 Pressed");
            m_CurrentValue += "3";
        }

        private void btnNum4_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 4 Pressed");
            m_CurrentValue += "4";
        }

        private void btnNum5_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 5 Pressed");
            m_CurrentValue += "5";
        }

        private void btnNum6_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 6 Pressed");
            m_CurrentValue += "6";
        }

        private void btnNum7_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 7 Pressed");
            m_CurrentValue += "7";
        }

        private void btnNum8_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 8 Pressed");
            m_CurrentValue += "8";
        }

        private void btnNum9_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 9 Pressed");
            m_CurrentValue += "9";
        }

        private void btnNum0_Click(object sender, EventArgs e)
        {
            Debug.WriteLine("Key 0 Pressed");
            m_CurrentValue += "0";
        }
    }
}
