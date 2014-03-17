// Re-usable ActiveX helper functions
// March 2014
// Paul Voelker

using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace ActiveXHelpers
{
    internal static class ActiveXHelpers
    {
        public static int OLEMISC_RECOMPOSEONRESIZE             = 0x1;
        public static int OLEMISC_ONLYICONIC                    = 0x2;
        public static int OLEMISC_INSERTNOTREPLACE              = 0x4;
        public static int OLEMISC_STATIC                        = 0x8;
        public static int OLEMISC_CANTLINKINSIDE                = 0x10;
        public static int OLEMISC_CANLINKBYOLE1                 = 0x20;
        public static int OLEMISC_ISLINKOBJECT                  = 0x40;
        public static int OLEMISC_INSIDEOUT                     = 0x80;
        public static int OLEMISC_ACTIVATEWHENVISIBLE           = 0x100;
        public static int OLEMISC_RENDERINGISDEVICEINDEPENDENT  = 0x200;
        public static int OLEMISC_INVISIBLEATRUNTIME            = 0x400;
        public static int OLEMISC_ALWAYSRUN                     = 0x800;
        public static int OLEMISC_ACTSLIKEBUTTON                = 0x1000;
        public static int OLEMISC_ACTSLIKELABEL                 = 0x2000;
        public static int OLEMISC_NOUIACTIVATE                  = 0x4000;
        public static int OLEMISC_ALIGNABLE                     = 0x8000;
        public static int OLEMISC_SIMPLEFRAME                   = 0x10000;
        public static int OLEMISC_SETCLIENTSITEFIRST            = 0x20000;
        public static int OLEMISC_IMEMODE                       = 0x40000;
        public static int OLEMISC_IGNOREACTIVATEWHENVISIBLE     = 0x80000;
        public static int OLEMISC_WANTSTOMENUMERGE              = 0x100000;
        public static int OLEMISC_SUPPORTSMULTILEVELUNDO        = 0x200000;

        public static void RegisterClass(Type t, int bitmapId, int versionMajor,
            int versionMinor)
        {
            string keyName = @"CLSID\" + t.GUID.ToString("B");

            // ActiveX registry info: http://msdn.microsoft.com/en-us/library/windows/desktop/ms695220%28v=vs.85%29.aspx
            //                        http://blogs.msdn.com/b/andreww/archive/2008/11/24/using-managed-controls-as-activex-controls.aspx
            //                        http://msdn.microsoft.com/en-us/library/tzat5yw6(v=vs.80).aspx

            using (RegistryKey k = Registry.ClassesRoot.OpenSubKey(keyName, true))
            {
                using (RegistryKey ctrl = k.CreateSubKey("Control")) { }

                using (RegistryKey inprocServer32Key = k.OpenSubKey("InprocServer32",
                    true))
                {
                    inprocServer32Key.SetValue("CodeBase",
                        Assembly.GetExecutingAssembly().CodeBase);
                }

                using (RegistryKey miscStatusKey = k.CreateSubKey("MiscStatus"))
                {
                    int nMiscStatus = OLEMISC_RECOMPOSEONRESIZE | OLEMISC_CANTLINKINSIDE |
                        OLEMISC_INSIDEOUT | OLEMISC_ACTIVATEWHENVISIBLE |
                        OLEMISC_SETCLIENTSITEFIRST;
                    miscStatusKey.SetValue("", nMiscStatus.ToString(),
                        RegistryValueKind.String);
                }

                using (RegistryKey toolBoxBitmapKey = k.CreateSubKey("ToolboxBitmap32"))
                {
                    toolBoxBitmapKey.SetValue("",
                        Assembly.GetExecutingAssembly().Location + ", " +
                        bitmapId.ToString(), RegistryValueKind.String);
                }

                using (RegistryKey typeLibKey = k.CreateSubKey("TypeLib"))
                {
                    Guid libId = Marshal.GetTypeLibGuidForAssembly(t.Assembly);
                    typeLibKey.SetValue("", libId.ToString("B"),
                        RegistryValueKind.String);
                }

                using (RegistryKey versionKey = k.CreateSubKey("Version"))
                {
                    versionKey.SetValue("", string.Format("{0}.{1}", versionMajor,
                        versionMinor));
                }
            }
        }

        public static void UnregisterClass(Type t)
        {
            string keyName = @"CLSID\" + t.GUID.ToString("B");
            Registry.ClassesRoot.DeleteSubKeyTree(keyName);
        }
    }
}
