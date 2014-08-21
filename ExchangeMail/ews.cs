using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;

using System.Net;
using Microsoft.Exchange.WebServices.Data;
using utils;
using Helpers;

using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

using System.Threading;

namespace ExchangeMail
{
    //see interfaces //public delegate void StateChangedEventHandler(object sender, StatusEventArgs args);

    public class ews : IMailHost
    {
        Microsoft.Exchange.WebServices.Data.ExchangeService service = null;

        const string hsm_service_url = "https://az18-cas-01.global.ds.honeywell.com/EWS/Exchange.asmx";
        static string serviceURL = "https://az18-cas-01.global.ds.honeywell.com/EWS/Exchange.asmx";
        // EMEA-CAS-01.global.ds.honeywell.com
        const string webProxy = "fr44proxy.honeywell.com";
        const int webProxyPort = 8080;

        Thread _thread;
        volatile bool _bRunThread = true;

        public bool _enableTrace = false;

        bool bRunPullThread = true;
        bool bHandleEvents = true;

        public userData UserData { get; set; }
        
        /// <summary>
        /// initialize a new ExchangeWebService object
        /// </summary>
        public ews()
        {
            
        }

        public void start()
        {
            if (service != null)
            {
                OnStateChanged(new StatusEventArgs(StatusType.error, "Already connected"));
                return;
            }
            try
            {
                serviceURL = hsm_service_url;
                // Hook up the cert callback.
                System.Net.ServicePointManager.ServerCertificateValidationCallback =
                    delegate(
                        Object obj,
                        X509Certificate certificate,
                        X509Chain chain,
                        SslPolicyErrors errors)
                    {
                        // Validate the certificate and return true or false as appropriate.
                        // Note that it not a good practice to always return true because not
                        // all certificates should be trusted.
                        helpers.addLog("certificate validated");
                        OnStateChanged(new StatusEventArgs(StatusType.validating));
                        return true;
                    };

                service = new ExchangeService(Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2010_SP2);

                if (_enableTrace)
                {
                    service.TraceListener = new TraceListener();// ITraceListenerInstance;
                    // Optional flags to indicate the requests and responses to trace.
                    service.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
                    service.TraceEnabled = true;
                }
            }
            catch (Exception ex)
            {
                helpers.addLog("Exception: " + ex.Message);
            }

        }

        /// <summary>
        /// logon using domain, user and password
        /// </summary>
        /// <param name="sDomain">the domain for login</param>
        /// <param name="sUser">user name</param>
        /// <param name="sPassword">password</param>
        /// <returns>true if success</returns>

        public bool logon(string sDomain, string sUser, string sPassword, bool bProxy)
        {
            OnStateChanged(new StatusEventArgs(StatusType.busy, "logon..."));
            bool bRet = false;
            UserData = new userData(sDomain, sUser, sPassword, bProxy);
            try
            {
                // Or use NetworkCredential directly (WebCredentials is a wrapper
                // around NetworkCredential).
                service.Credentials = new NetworkCredential(sUser, sPassword, sDomain);
                if(UserData.bUseProxy)
                    service.WebProxy = new WebProxy(webProxy, webProxyPort);
                service.Url = new Uri(hsm_service_url);
                helpers.addLog("connected to: " + service.Url.ToString());
                OnStateChanged(new StatusEventArgs(StatusType.idle, "connected to: " + service.Url.ToString()));

                bRet = true;
            }
            catch (Exception ex)
            {
                helpers.addLog("Exception: " + ex.Message);
                OnStateChanged(new StatusEventArgs(StatusType.error, "logon failed: " + ex.Message));
            }
            return bRet;
        }

        /// <summary>
        /// url autodiscover validation callback
        /// </summary>
        /// <param name="url"></param>
        /// <returns>true if URL is valid</returns>
        private bool ValidateRedirectionUrlCallback(string url)
        {
            // Validate the URL and return true to allow the redirection or false to prevent it.
            helpers.addLog("Autodiscovery changed to " + url);
            serviceURL = url;
            OnStateChanged(new StatusEventArgs(StatusType.url_changed, "new URL = " + url));
            return true;
        }

        public void Dispose()
        {
            _bRunThread = false;
            if (_thread != null)
            {
                //try to join thread
                bool b = _thread.Join(1000);
                if (!b)
                    _thread.Abort();
            }

            bRunPullThread = false;
            if (_pullMails != null)
            {
                bool b = _pullMails.Join(1000);
                if (!b)
                    _pullMails.Abort();
            }

        }

        #region Threading_stuff
        // see interfaces //public event StateChangedEventHandler stateChangedEvent;  
        public event Helpers.StateChangedEventHandler StateChanged;
        protected virtual void OnStateChanged(StatusEventArgs args)
        {
            if (!bHandleEvents)
                return;
            System.Diagnostics.Debug.WriteLine("onStateChanged: " + args.eStatus.ToString() + ":" + args.strMessage);
            StateChangedEventHandler handler = StateChanged;
            if (handler != null)
            {
                handler(this, args);
            }
        }

        public void getMailsAsync()
        {            
            _thread = new Thread(_getMails);
            _thread.Name = "getMail thread";
            _thread.Start(this);

            return;
        }

        //thread to get all emails
        public static void _getMails(object param)
        {
            helpers.addLog("_getMails() started...");
            ews _ews = (ews)param;  //need an instance
            _ews.OnStateChanged(new StatusEventArgs(StatusType.busy, "_getMails() started..."));
            const int chunkSize = 10;
            try
            {
                //blocking call
                ItemView view = new ItemView(chunkSize);
                view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
                /*
                private static void GetAttachments(ExchangeService service)
                {
                    // Return a single item.
                    ItemView view = new ItemView(1);

                    string querystring = "HasAttachments:true Subject:'Message with Attachments' Kind:email";

                    // Find the first email message in the Inbox that has attachments. This results in a FindItem operation call to EWS.
                    FindItemsResults<Item> results = service.FindItems(WellKnownFolderName.Inbox, querystring, view);

                    if (results.TotalCount > 0)
                    {
                        EmailMessage email = results.Items[0] as EmailMessage;
                */
                FindItemsResults<Item> findResults;

                do
                {
                    //findResults = service.FindItems(WellKnownFolderName.Inbox, view);
                    findResults = _ews.service.FindItems(
                        WellKnownFolderName.Inbox,
                        new SearchFilter.SearchFilterCollection(
                            LogicalOperator.And,
                            new SearchFilter.ContainsSubstring(ItemSchema.Subject, helpers.filterSubject),
                            new SearchFilter.ContainsSubstring(ItemSchema.Attachments, helpers.filterAttachement)),
                            view);
                    foreach (Item item in findResults.Items)
                    {
                        helpers.addLog("found item...");
                        if (item is EmailMessage)
                        {
                            helpers.addLog("\t is email ...");

                            // If the item is an e-mail message, write the sender's name.
                            helpers.addLog((item as EmailMessage).Sender.Name + ": " + (item as EmailMessage).Subject);

                            MailMsg myMailMsg = new MailMsg(item as EmailMessage, _ews.UserData.sUser);

                            // Bind to an existing message using its unique identifier.
                            //EmailMessage message = EmailMessage.Bind(service, new ItemId(item.Id.UniqueId));
                            _ews.OnStateChanged(new StatusEventArgs(StatusType.license_mail, myMailMsg));
                        }
                    }
                    view.Offset += chunkSize;
                } while (findResults.MoreAvailable && _ews._bRunThread);
                _ews.OnStateChanged(new StatusEventArgs(StatusType.idle, "readmail done"));
            }
            catch (ThreadAbortException ex)
            {
                helpers.addLog("ThreadAbortException: " + ex.Message);
            }
            catch (Exception ex)
            {
                helpers.addLog("Exception: " + ex.Message);
                _ews.OnStateChanged(new StatusEventArgs(StatusType.error, "readmail exception: " + ex.Message));
            }
            helpers.addLog("_getMails() ended");
            _ews.startPull();
        }
        #endregion

        #region TESTING
        //public string getMails()
        //{
        //    helpers.addLog("getMails started...");
        //    StringBuilder sb = new StringBuilder();

        //    //blocking call
        //    FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, new ItemView(20));

        //    helpers.addLog("getMails: found " + findResults.Items.Count.ToString() + " items");
        //    foreach (Item item in findResults.Items)
        //    {
        //        if (item is EmailMessage)
        //        {
        //            helpers.addLog("found email item...");
        //            // If the item is an e-mail message, write the sender's name.
        //            Console.WriteLine((item as EmailMessage).Sender.Name);

        //            // Bind to an existing message using its unique identifier.
        //            //EmailMessage message = EmailMessage.Bind(service, new ItemId(item.Id.UniqueId));

        //            //cast item to EmailMessage
        //            EmailMessage message = (EmailMessage)item;

        //            sb.Append("\r\n" + message.From.Name);
        //            sb.Append("\r\n" + message.Subject);
        //            if (message.Attachments.Count > 0)
        //            {
        //                sb.Append("\r\n\tAttachements\r\n");
        //                foreach (Attachment a in message.Attachments)
        //                    sb.Append("\r\n\t" + a.Name);
        //            }
        //            sb.Append("\r\n===================\r\n");
        //        }

        //        //Console.WriteLine(item.Subject);
        //    }
        //    helpers.addLog(sb.ToString());
        //    helpers.addLog("getMails done");
        //    return sb.ToString();
        //}

        ////Retrieve all the items in the Inbox by groups of 50 items
        //public void PageThroughEntireInbox()
        //{
        //    ItemView view = new ItemView(50);
        //    FindItemsResults<Item> findResults;

        //    do
        //    {
        //        findResults = service.FindItems(WellKnownFolderName.Inbox, view);

        //        foreach (Item item in findResults.Items)
        //        {
        //            // Do something with the item.
        //        }

        //        view.Offset += 50;
        //    } while (findResults.MoreAvailable);
        //}

        ////Find the first 10 messages in the Inbox that have a subject that contains the words "EWS" or "API", order by date received, and only return the Subject and DateTimeReceived properties
        //public void FindItems()
        //{
        //    ItemView view = new ItemView(10);
        //    view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
        //    view.PropertySet = new PropertySet(
        //        BasePropertySet.IdOnly,
        //        ItemSchema.Subject,
        //        ItemSchema.DateTimeReceived);

        //    FindItemsResults<Item> findResults = service.FindItems(
        //        WellKnownFolderName.Inbox,
        //        new SearchFilter.SearchFilterCollection(
        //            LogicalOperator.Or,
        //            new SearchFilter.ContainsSubstring(ItemSchema.Subject, "EWS"),
        //            new SearchFilter.ContainsSubstring(ItemSchema.Subject, "API")),
        //        view);

        //    Console.WriteLine("Total number of items found: " + findResults.TotalCount.ToString());

        //    foreach (Item item in findResults)
        //    {
        //        // Do something with the item.
        //    }
        //}
        #endregion

        Thread _pullMails = null;
        void startPull()
        {
            bRunPullThread = true;
            _pullMails = new Thread(startPullNotification);
            _pullMails.Name = "Pull thread";
            _pullMails.Start(this);
        }

        //pull notifications
        static void startPullNotification(object param)
        {
            ews _ews = (ews)param;  //need an instance
            _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "Pull started"));
#if DEBUG
            int iTimeoutMinutes = 2;
#else
            int iTimeoutMinutes = 5;
#endif

            // Subscribe to pull notifications in the Inbox folder, and get notified when
            // a new mail is received, when an item or folder is created, or when an item
            // or folder is deleted.
            PullSubscription subscription=null;
            try
            {
                subscription = _ews.service.SubscribeToPullNotifications(new FolderId[] { WellKnownFolderName.Inbox },
                iTimeoutMinutes /* timeOut: the subscription will end if the server is not polled within 5 minutes. */,
                null /* watermark: null to start a new subscription. */,
                EventType.NewMail);//, EventType.Created, EventType.Deleted);
            }
            catch (Exception ex)
            {
                _ews.StateChanged(null, new StatusEventArgs(StatusType.error, "PullSubscription FAILED: "+ex.Message));
            }
            if (subscription == null)
                return;
            try
            {
                while (_ews.bRunPullThread)
                {
                    Thread.Sleep((iTimeoutMinutes - 1) * 60000);
                    // Wait a couple minutes, then poll the server for new events.   
                    GetEventsResults events = subscription.GetEvents();

                    // Loop through all item-related events.
                    foreach (ItemEvent itemEvent in events.ItemEvents)
                    {
                        switch (itemEvent.EventType)
                        {
                            case EventType.NewMail:
                                // A new mail has been received. Bind to it
                                EmailMessage message = EmailMessage.Bind(_ews.service, itemEvent.ItemId);
                                bool bDoProcessMail=false;
                                if(message.Subject.Contains(helpers.filterSubject)){
                                    if(message.HasAttachments){
                                        foreach (Attachment att in message.Attachments)
                                        {
                                            if (att.Name.EndsWith(helpers.filterAttachement))
                                            {
                                                bDoProcessMail = true;
                                                continue;
                                            }
                                        }
                                    }
                                    if (bDoProcessMail)
                                    {
                                        //create a new IMailMessage from the EmailMessage
                                        MailMsg myMailMsg = new MailMsg(message, _ews.UserData.sUser);
                                        _ews.OnStateChanged(new StatusEventArgs(StatusType.license_mail, myMailMsg));
                                    }
                                    else
                                        _ews.OnStateChanged(new StatusEventArgs(StatusType.other_mail, "mismatched mail received"));
                                }
                                break;
                            case EventType.Created:
                                // An item was created in the folder. Bind to it.
                                Item item = Item.Bind(_ews.service, itemEvent.ItemId);
                                break;
                            case EventType.Deleted:
                                // An item has been deleted. Output its ID to the console.
                                Console.WriteLine("Item deleted: " + itemEvent.ItemId.UniqueId);
                                break;
                        }
                    }

                    // Loop through all folder-related events.
                    foreach (FolderEvent folderEvent in events.FolderEvents)
                    {
                        switch (folderEvent.EventType)
                        {
                            case EventType.Created:
                                // An folder was created. Bind to it.
                                Folder folder = Folder.Bind(_ews.service, folderEvent.FolderId);
                                break;
                            case EventType.Deleted:
                                // A folder has been deleted. Output its Id to the console.
                                Console.WriteLine("Folder deleted:" + folderEvent.FolderId.UniqueId);
                                break;
                        }
                    }

                    /*
                    // As an alternative, you can also loop through all the events.
                    foreach (NotificationEvent notificationEvent in events.AllEvents)
                    {
                        if (notificationEvent is ItemEvent)
                        {
                            ItemEvent itemEvent = notificationEvent as ItemEvent;

                            switch (itemEvent.EventType)
                            {
                                ...
                            }
                        }
                        else
                        {
                            FolderEvent folderEvent = notificationEvent as FolderEvent;

                            switch (folderEvent.EventType)
                            {
                                ...
                            }
                        }
                    }
                    */
                }//end while
            }
            catch (ThreadAbortException)
            {
                _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "Pull Thread ThreadAbortException"));
                _ews.bRunPullThread = false;
            }
            catch (Exception ex)
            {
                _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "Pull Thread Exception:" + ex.Message));
                _ews.bRunPullThread = false;
            }
            subscription.Unsubscribe();
            subscription = null;
            _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "Pull ended"));
        }

        class TraceListener : ITraceListener
        {
            #region ITraceListener Members
            public void Trace(string traceType, string traceMessage)
            {
                CreateXMLTextFile(traceType, traceMessage.ToString());
            }
            #endregion

            private void CreateXMLTextFile(string fileName, string traceContent)
            {
                // Create a new XML file for the trace information.
                try
                {
                    // If the trace data is valid XML, create an XmlDocument object and save.
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(traceContent);
                    xmlDoc.Save(fileName + ".xml");
                }
                catch
                {
                    // If the trace data is not valid XML, save it as a text document.
                    System.IO.File.WriteAllText(fileName + ".txt", traceContent);
                }
            }
        }
    }
}
