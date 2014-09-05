using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Threading;

using utils;
using Helpers;

using Microsoft.Office;
//using Microsoft.Office.Interop.Outlook;
//using Microsoft.Office.Interop;
//using Microsoft.Office.Interop.Outlook.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMail
{
    public class OutlookMail:IMailHost
    {
        Outlook.Application MyApp = null;
        Outlook.NameSpace MailNS = null;
        Outlook.MAPIFolder MyInbox = null;
        Outlook.Items MyItems = null;

        public UserData _userData {get;set;}
        Thread _thread;
        volatile bool _bRunThread = true;
        Helpers.LicenseMail _licenseMail = null;
        const string sMailHasAlreadyProcessed = "[processed]";

        public OutlookMail(ref LicenseMail licenseMail)
        {
            _licenseMail = licenseMail;
            _userData = new UserData();
        }
        public void start()
        {
            try
            {
                //start outlook
                MyApp = new Outlook.Application();
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
            bool bNewSession = true;
            _userData.sUser = sUser;
            try
            {
                MailNS.Logon(sProfile, sPassword, bShowDialog, bNewSession);
                MyInbox = MailNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                OnStateChanged(new StatusEventArgs(StatusType.ews_started, "logon done"));

                getMailsAsync();

                MyItems = MyInbox.Items;
                MyItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(newItem);

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
            _thread = new Thread(_getMails);
            _thread.Name = "getMail thread";
            _thread.Start(this);

            return;
        }
        void _getMails(object param)
        {
            helpers.addLog("_getMails() started...");
            OnStateChanged(new StatusEventArgs(StatusType.busy, "looking thru eMails"));
            OutlookMail _olk = (OutlookMail)param;  //need an instance
            _olk.OnStateChanged(new StatusEventArgs(StatusType.busy, "_getMails() started..."));
            try
            {
                _olk.OnStateChanged(new StatusEventArgs(StatusType.ews_pulse, "looking thru emails"));
                Outlook.Items items = MyInbox.Items;
                items.Sort("[ReceivedTime]");
                foreach (object o in items)// MyInbox.Items.Sort("[ReceivedTime]"))
                {
                    processMail(o);
                    /*
                    if (o == null)
                        continue;
                    int foundAttachements=0;
                    if (o == null)
                        continue;
                    Outlook.MailItem mitem = o as Outlook.MailItem;
                    if (mitem == null)
                        continue;
                    if(mitem.Subject!=null)
                        helpers.addLog(string.Format("looking at: {0}\n", mitem.Subject));
                    else
                        helpers.addLog(string.Format("looking at: {0}\n", mitem.EntryID));

                    if (mitem.Attachments.Count == 0)
                        continue;
                    else
                    {
                        //look for .xml
                        foreach (Outlook.Attachment att in mitem.Attachments)
                        {
                            if (att.FileName.EndsWith("xml"))
                                foundAttachements++;
                        }
                    }
                    if (foundAttachements == 0)
                        continue;   //no attachement ending in xml
                    if(mitem.Subject.IndexOf(helpers.filterSubject)==-1)
                        continue;
                    if (mitem.Subject.IndexOf(sMailHasAlreadyProcessed) > 0)  //do not process mail again
                        continue;

                    //process mail
                    helpers.addLog(string.Format("### mail found: {0},\n{1}\n\n", mitem.Subject, mitem.Body));
                    _olk.OnStateChanged(new StatusEventArgs(StatusType.ews_pulse, "getMail: processing eMail " + mitem.Subject));
                    MailMsg myMailMsg = new MailMsg(mitem, _olk._userData.sUser);
                    int iRet = _olk._licenseMail.processMail(myMailMsg);
                    //change subject
                    mitem.Subject += sMailHasAlreadyProcessed;
                    mitem.Close(Outlook.OlInspectorClose.olSave);
                    Thread.Sleep(100);
                    Thread.Yield();
                    */
                }
            }
            catch (System.Exception ex)
            {
                helpers.addExceptionLog(ex.StackTrace);
            }
            _olk.OnStateChanged(new StatusEventArgs(StatusType.ews_pulse, "read mail done"));
        }

        int processMail(object oMsg)
        {
            if (oMsg == null)
                return -1;
            int foundAttachements = 0;
            if (oMsg == null)
                return -2;
            Outlook.MailItem mitem = oMsg as Outlook.MailItem;
            if (mitem == null)
                return -3;
            if (mitem.Subject != null)
                helpers.addLog(string.Format("looking at: {0}\n", mitem.Subject));
            else
                helpers.addLog(string.Format("looking at: {0}\n", mitem.EntryID));

            if (mitem.Attachments.Count == 0)
                return -4;
            else
            {
                //look for .xml
                foreach (Outlook.Attachment att in mitem.Attachments)
                {
                    if (att.FileName.EndsWith("xml"))
                        foundAttachements++;
                }
                if (foundAttachements == 0)
                    return -5;   //no attachement ending in xml
            }
            if (mitem.Subject.IndexOf(helpers.filterSubject) == -1)
                return -6;
            if (mitem.Subject.IndexOf(sMailHasAlreadyProcessed) >= 0)  //do not process mail again
                return -7;

            //process mail
            helpers.addLog(string.Format("### mail found: {0},\n{1}\n\n", mitem.Subject, mitem.Body));
            OnStateChanged(new StatusEventArgs(StatusType.ews_pulse, "getMail: processing eMail " + mitem.Subject));
            MailMsg myMailMsg = new MailMsg(mitem, _userData.sUser);
            int iRet = _licenseMail.processMail(myMailMsg);
            
            //change subject
            if (mitem.Subject.IndexOf(sMailHasAlreadyProcessed) == -1)
            {
                mitem.Subject += sMailHasAlreadyProcessed;
                mitem.Close(Outlook.OlInspectorClose.olSave);
            }
            
            OnStateChanged(new StatusEventArgs(StatusType.license_mail, "processed " + iRet.ToString()));
            Thread.Sleep(100);
            Thread.Yield();
            return iRet;
        }

        void newItem(object item)
        {
            OnStateChanged(new StatusEventArgs(StatusType.ews_pulse, "new mail"));
            try
            {
                processMail(item);
            }
            catch (Exception ex)
            {
                helpers.addExceptionLog(ex.StackTrace);
            }
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
            OnStateChanged(new StatusEventArgs(StatusType.ews_stopped, "stopped"));
            _bRunThread = false;
            if (_thread != null)
            {
                //try to join thread
                bool b = _thread.Join(1000);
                if (!b)
                    _thread.Abort();
            } 
            MailNS.Logoff();
            OnStateChanged(new StatusEventArgs(StatusType.ews_stopped, "outlook session ended"));
        }

    }
}
