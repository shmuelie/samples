using System;
using System.Runtime.InteropServices;

namespace SelfElevation
{
    public class Program
    {
        public static void Main()
        {
            object obj;
            int hr = CoCreateInstanceAsAdmin(IntPtr.Zero, Contract.Constants.ServerClassGuid, typeof(IServer).GUID, out obj);
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }

            var server = (IServer)obj;
            double pi = server.ComputePi();
            Console.WriteLine($"\u03C0 = {pi}");
        }

        private static int CoCreateInstanceAsAdmin(IntPtr hwnd, Guid rclsid, Guid riid, out object ppv)
        {
            string monikerName = "Elevation:Administrator!new:" + rclsid.ToString("B").ToUpper();
            Ole32.BIND_OPTS3 bo = default;
            bo.cbStruct = Marshal.SizeOf<Ole32.BIND_OPTS3>();
            bo.hwnd = hwnd;
            bo.dwClassContext = Ole32.CLSCTX_LOCAL_SERVER;
            return Ole32.CoGetObject(monikerName, ref bo, riid, out ppv);
        }

        private class Ole32
        {
            // https://docs.microsoft.com/windows/win32/api/wtypesbase/ne-wtypesbase-clsctx
            public const int CLSCTX_LOCAL_SERVER = 0x4;

            public struct BIND_OPTS3
            {
                public int cbStruct;
                public uint grfGlags;
                public uint grfMode;
                public uint dwTickCountDeadline;
                public uint dwTrackFlags;
                public uint dwClassContext;
                public uint locale;
                public IntPtr pServerInfo;
                public IntPtr hwnd;
            }

            [DllImport(nameof(Ole32), CharSet = CharSet.Unicode)]
            public static extern int CoGetObject(
                [In] string pszName,
                [In, Optional] ref BIND_OPTS3 pBindOptions,
                [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid,
                [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        }
    }
}
