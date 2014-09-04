using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using utils;
using Helpers;

using Microsoft.Office;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Outlook.Extensions;

namespace OutlookMail
{
    public class OutlookMail:IMailHost
    {
        Microsoft.Office.Interop.Outlook.Application MyApp = null;
        Microsoft.Office.Interop.Outlook.NameSpace MailNS = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder MyInbox = null;
        public UserData _userData {get;set;}

        public OutlookMail()
        {
        }
        public void start()
        {
            try
            {
                //start outlook
                MyApp = new Application();
                MailNS = MyApp.GetNamespace("MAPI");
            }
            catch (System.Exception ex)
            {
                helpers.addLog("Exception: " + ex.Message);
            }
        }
        public bool logon(string sProfile, string sUser, string sPassword, bool bShowDialog)
        {
            bool bRet = false;
            bool bNewSession = false;
            try
            {
                MailNS.Logon(sProfile, sPassword, bShowDialog, bNewSession);
                MyInbox = MailNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                bRet = true;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Exception: " + ex.Message);
            }
            return bRet;
        }
        
        public void getMailsAsync()
        {
            //_thread = new Thread(_getMails);
            //_thread.Name = "getMail thread";
            //_thread.Start(this);

            return;
        }

        bool _bHandleEvents = true;
        public event Helpers.StateChangedEventHandler StateChanged;
        protected virtual void OnStateChanged(StatusEventArgs args)
        {
            if (!_bHandleEvents)
                return;
            System.Diagnostics.Debug.WriteLine("onStateChanged: " + args.eStatus.ToString() + ":" + args.strMessage);
            StateChangedEventHandler handler = StateChanged;
            if (handler != null)
            {
                handler(this, args);
            }
        }
        public void Dispose()
        {
            MailNS.Logoff();            
        }

    }
}
