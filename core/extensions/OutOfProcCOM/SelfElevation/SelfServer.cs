using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace SelfElevation
{
    [ComVisible(true)]
    [Guid(Contract.Constants.ServerClass)]
    [ComDefaultInterface(typeof(IServer))]
    public class SelfServer : IServer
    {
        double IServer.ComputePi()
        {
            Trace.WriteLine($"Running {nameof(SelfServer)}.{nameof(IServer.ComputePi)}");
            double sum = 0.0;
            int sign = 1;
            for (int i = 0; i < 1024; ++i)
            {
                sum += sign / (2.0 * i + 1.0);
                sign *= -1;
            }

            return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator) ? (4.0 * sum) : (-4.0 * sum);
        }

#if EMBEDDED_TYPE_LIBRARY
        private static readonly string tlbPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), $"{nameof(SelfServer)}.comhost.dll");
#else
        private static readonly string tlbPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Contract.Constants.TypeLibraryName);
#endif

        [ComRegisterFunction]
        internal static void RegisterFunction(Type t)
        {
            if (t != typeof(SelfServer))
                return;

            // Register DLL surrogate and type library
            COMRegistration.DllSurrogate.Register(Contract.Constants.ServerClassGuid, tlbPath);
            COMRegistration.DllSurrogate.AllowElevation(Contract.Constants.ServerClassGuid, Assembly.GetExecutingAssembly().Location, 102, 103);
        }

        [ComUnregisterFunction]
        internal static void UnregisterFunction(Type t)
        {
            if (t != typeof(SelfServer))
                return;

            // Unregister DLL surrogate and type library
            COMRegistration.DllSurrogate.Unregister(Contract.Constants.ServerClassGuid, tlbPath);
            COMRegistration.DllSurrogate.DisabllowElevation(Contract.Constants.ServerClassGuid);
        }
    }
}
