using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace ManagedMAPI
{

    /// <summary>
    /// IMAPISession .Net wrapper object
    /// </summary>
    public class MAPISession : IDisposable
    {
        /// <summary>
        /// New mail event
        /// </summary>
        IMAPISession session_ = null;
     

        /// <summary>
        /// Initializes a new instance of the MAPISession class. 
        /// </summary>
        public MAPISession()
        {
            Initialize();
        }

        #region Public Properties
        
      
        #endregion


        #region Public Methods
                
        public Dictionary<string, bool> GetMessageStores()
        {
            Dictionary<string, bool> stores = new Dictionary<string, bool>();
            if (session_ == null)
                return stores;
            IntPtr pTable = IntPtr.Zero;
            session_.GetMsgStoresTable(0, out pTable);
            if (pTable == IntPtr.Zero)
                return stores;
            object tableObj = null;
            MAPITable mb = null;
            try
            {
                tableObj = Marshal.GetObjectForIUnknown(pTable);
                mb = new MAPITable(tableObj as IMAPITable);
                if (mb != null)
                {
                    if (mb.SetColumns(new PropTags[] { PropTags.PR_DISPLAY_NAME, PropTags.PR_DEFAULT_STORE }))
                    {
                        SRow[] sRows;

                        while (mb.QueryRows(1, out sRows))
                        {
                            if (sRows.Length != 1)
                                break;
                            stores[sRows[0].propVals[0].AsString] = sRows[0].propVals[1].AsBool;
                        }
                    }
                    mb.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
            }
            return stores;
        }
      
        #endregion

        #region Implement IDisposable Interface

        public void Dispose()
        {
            Dispose(true);
        }

        #endregion 

        #region private methods

        /// <summary>
        /// MAPI session intialization and logon.
        /// </summary>
        /// <returns>true if successful; otherwise, false</returns>
        private bool Initialize()
        {
            IntPtr pSession = IntPtr.Zero;
            if (MAPINative.MAPIInitialize(IntPtr.Zero) == HRESULT.S_OK)
            {
                MAPINative.MAPILogonEx(0, null, null, (uint)(MAPIFlag.EXTENDED | MAPIFlag.USE_DEFAULT), out pSession);
                if (pSession == IntPtr.Zero)
                    MAPINative.MAPILogonEx(0, null, null, (uint)(MAPIFlag.EXTENDED | MAPIFlag.NEW_SESSION | MAPIFlag.USE_DEFAULT), out pSession);
            }

            if (pSession != IntPtr.Zero)
            {
                object sessionObj = null;
                try
                {
                    sessionObj = Marshal.GetObjectForIUnknown(pSession);
                    session_ = sessionObj as IMAPISession;
                }
                catch { }
            }

            return session_ != null;
        }
        
        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (session_ != null)
                {
                    Marshal.ReleaseComObject(session_);
                    session_ = null;
                    MAPINative.MAPIUninitialize();
                }
            }
        }
        #endregion

    }
}

