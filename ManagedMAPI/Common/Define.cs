using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace ManagedMAPI
{
    #region public MAPI definitions
  
    /// <summary>
    /// Describes a row from a table containing selected properties for a specific object.
    /// </summary>
    public struct SRow
    {
        /// <summary>
        /// An array of SPropValue structures that describe the property values for the columns in the row. 
        /// </summary>
        public IPropValue[] propVals;
    }
   
   /// <summary>
    /// Bitmask of flags that controls the mapi character set.
    /// </summary>
    public enum CharacterSet : uint
    {
        /// <summary>
        /// Ansi
        /// </summary>
        ANSI = 0,
        /// <summary>
        /// Unicode
        /// </summary>
        UNICODE = 0x80000000
    }

    /// <summary>
    /// Generic MAPI Bitmask
    /// </summary>
    public enum MAPIFlag : uint
    {
        /// <summary>
        /// An attempt should be made to create a new MAPI session instead of acquiring the shared session. 
        /// </summary>
        NEW_SESSION = 0x00000002,
        /// <summary>
        /// An attempt should be made to download all of the user's messages before returning. If it is not set, messages can be downloaded in the background after the call to MAPILogonEx returns.
        /// </summary>
        FORCE_DOWNLOAD = 0x00001000,
        /// <summary>
        /// A dialog box should be displayed to prompt the user for logon information if required.
        /// </summary>
        MAPI_LOGON_UI = 0x00000001,
        /// <summary>
        /// The shared session should be returned, which allows later clients to obtain the session without providing any user credentials.
        /// </summary>
        ALLOW_OTHERS = 0x00000008,
        /// <summary>
        /// The default profile should not be used and the user should be required to supply a profile.
        /// </summary>
        EXPLICIT_PROFILE = 0x00000010,
        /// <summary>
        /// Log on with extended capabilities. This flag should always be set.
        /// </summary>
        EXTENDED = 0x00000020,
        /// <summary>
        /// MAPILogonEx should display a configuration dialog box for each message service in the profile. 
        /// </summary>
        SERVICE_UI_ALWAYS = 0x00002000,
        /// <summary>
        /// MAPI should not inform the MAPI spooler of the session's existence.
        /// </summary>
        NO_MAIL = 0x00008000,
        /// <summary>
        /// The messaging subsystem should substitute the profile name of the default profile for the Profile Name parameter. T
        /// </summary>
        USE_DEFAULT = 0x00000040,
        /// <summary>
        /// Requests that the object be opened by using the maximum network permissions allowed for the user and the maximum client application access. 
        /// </summary>
        BEST_ACCESS = 0x00000010,
        /// <summary>
        /// MAPI spooler is a process that is responsible for sending messages to and receiving messages from a messaging system.
        /// </summary>
        SPOOLER = 37,
        /// <summary>
        /// Suppresses the display of dialog boxes. If the AB_NO_DIALOG flag is not set, the address book providers that contribute to the integrated address book can prompt the user for any necessary information.
        /// </summary>
        AB_NO_DIALOG = 0x00000001,
        /// <summary>
        /// The ulUIParam parameter is ignored unless the MAPI_DIALOG flag is set.
        /// </summary>
        MAPI_DIALOG = 0x00000008,
    }
    
    #endregion

    #region internal definitions
    [StructLayout(LayoutKind.Explicit)]
    internal struct _PV
    {
        [FieldOffset(0)]
        public Int16 i;
        [FieldOffset(0)]
        public int l;
        [FieldOffset(0)]
        public uint ul;
        [FieldOffset(0)]
        public float flt;
        [FieldOffset(0)]
        public double dbl;
        [FieldOffset(0)]
        public UInt16 b;
        [FieldOffset(0)]
        public double at;
        [FieldOffset(0)]
        public IntPtr lpszA;
        [FieldOffset(0)]
        public IntPtr lpszW;
        [FieldOffset(0)]
        public IntPtr lpguid;
        /*[FieldOffset(0)]
        public IntPtr bin;*/
        [FieldOffset(0)]
        public UInt64 li;
        [FieldOffset(0)]
        public SBinary bin;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SPropValue
    {
        public UInt32 ulPropTag;
        public UInt32 dwAlignPad;
        public _PV Value;
    }

    #endregion
}
