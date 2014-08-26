using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ProfMan
{
    public static class ProfManLoader
    {
        #region public methods

        //64 bit dll location - defaults to <assemblydir>\ProfMan64.dll
        public static string DllLocation64Bit;
        //32 bit dll location - defaults to <assemblydir>\ProfMan.dll
        public static string DllLocation32Bit;
        

        public static Profiles new_Profiles()
        {
            return (Profiles)NewProfManObject(new Guid("EBC7A7B5-C614-47B3-A579-27A2C2C98A13"));
        }

        public static PropertyBag new_PropertyBag()
        {
            return (PropertyBag)NewProfManObject(new Guid("FC583D50-A2F5-4656-8B1D-360488B183D3"));
        }
                
        #endregion


        #region private methods

        static ProfManLoader()
        {
            //default locations of the dlls
            var vUri = new UriBuilder(Assembly.GetExecutingAssembly().CodeBase);
            string vPath = Uri.UnescapeDataString(vUri.Path + vUri.Fragment);
            string directory = Path.GetDirectoryName(vPath);
            DllLocation64Bit = Path.Combine(directory, "ProfMan64.dll");
            DllLocation32Bit = Path.Combine(directory, "ProfMan.dll");
        }


        [ComVisible(false)]
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000001-0000-0000-C000-000000000046")]
        private interface IClassFactory
        {
            void CreateInstance([MarshalAs(UnmanagedType.Interface)] object pUnkOuter, ref Guid refiid, [MarshalAs(UnmanagedType.Interface)] out object ppunk);
            void LockServer(bool fLock);
        }

        [ComVisible(false)]
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000000-0000-0000-C000-000000000046")]
        private interface IUnknown
        {
        }

        private delegate int DllGetClassObject(ref Guid ClassId, ref Guid InterfaceId, [Out, MarshalAs(UnmanagedType.Interface)] out object ppunk);
        private delegate int DllCanUnloadNow();

        //COM GUIDs
        private static Guid IID_IClassFactory = new Guid("00000001-0000-0000-C000-000000000046");
        private static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        //win32 functions to load\unload dlls and get a function pointer 
        private class Win32NativeMethods
        {
            [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
            public static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool FreeLibrary(IntPtr hModule);

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern IntPtr LoadLibraryW(string lpFileName);
        }

        //private variables
        private static IntPtr _profmanDllHandle = IntPtr.Zero;
        private static IntPtr _dllGetClassObjectPtr = IntPtr.Zero;
        private static DllGetClassObject _dllGetClassObject;
        private static readonly object _criticalSection = new object();

        private static IUnknown NewProfManObject(Guid guid)
        {
            object res = null;
            lock (_criticalSection)
            {
                IClassFactory ClassFactory;
                if (_profmanDllHandle.Equals(IntPtr.Zero))
                {
                    string dllPath;
                    if (IntPtr.Size == 8) dllPath = DllLocation64Bit;
                    else dllPath = DllLocation32Bit;
                    _profmanDllHandle = Win32NativeMethods.LoadLibraryW(dllPath);
                    if (_profmanDllHandle.Equals(IntPtr.Zero))
                        //throw new Exception(string.Format("Could not load '{0}'\nMake sure the dll exists.", dllPath));
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    _dllGetClassObjectPtr = Win32NativeMethods.GetProcAddress(_profmanDllHandle, "DllGetClassObject");
                    if (_dllGetClassObjectPtr.Equals(IntPtr.Zero))
                        //throw new Exception("Could not retrieve a pointer to the 'DllGetClassObject' function exported by the dll");
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    _dllGetClassObject =
                        (DllGetClassObject)
                        Marshal.GetDelegateForFunctionPointer(_dllGetClassObjectPtr, typeof (DllGetClassObject));
                }


                Object unk;
                int hr = _dllGetClassObject(ref guid, ref IID_IClassFactory, out unk);
                if (hr != 0) throw new Exception("DllGetClassObject failed with error code 0x" + hr.ToString("x8"));
                ClassFactory = unk as IClassFactory;
                ClassFactory.CreateInstance(null, ref IID_IUnknown, out res);

                //If the same class factory is returned as the one still
                //referenced by .Net, the call will be marshalled to the original thread
                //where that class factory was retrieved first.
                //Make .Net forget these objects
                Marshal.ReleaseComObject(unk);
                Marshal.ReleaseComObject(ClassFactory);
            } //lock

            return (res as IUnknown);
        }

        #endregion

    }

}
