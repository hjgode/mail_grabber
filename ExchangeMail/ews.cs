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
        Microsoft.Exchange.WebServices.Data.ExchangeService _service = null;
        Helpers.LicenseMail _licenseMail=null;

        string ExchangeWebServiceURL = "https://az18-cas-01.global.ds.honeywell.com/EWS/Exchange.asmx";
        string _ExchangeWebServiceURL
        {
            get { return _ExchangeWebServiceURL; }
        }
        static string serviceURL = "https://az18-cas-01.global.ds.honeywell.com/EWS/Exchange.asmx";
        // EMEA-CAS-01.global.ds.honeywell.com
        const string sMailHasAlreadyProcessed = "[processed]";
        string webProxy = "fr44proxy.honeywell.com";
        string _webProxy
        {
            get { return webProxy; }
        }
        int webProxyPort = 8080;
        int _webProxyPort
        {
            get { return webProxyPort; }
        }

        string _Username
        {
            get;
            set;
        }

        Thread _thread;
        volatile bool _bRunThread = true;

        public bool _enableTrace = false;

        bool _bRunPullThread = true;
        bool _bHandleEvents = true;

        public UserData _userData { get; set; }
        
        /// <summary>
        /// initialize a new ExchangeWebService object
        /// </summary>
        public ews(ref LicenseMail lm)
        {
            _licenseMail = lm;
        }
        [Obsolete]
        public ews(string sServiceURL, string sWebProxy, int iWebProxyPort, bool bUseProxy)
        {
            ExchangeWebServiceURL = sServiceURL;
            _userData.WebProxy = sWebProxy;
            _userData.WebProxyPort = iWebProxyPort;
            _userData.bUseProxy = bUseProxy;
        }
        [Obsolete]
        public ews(string sServiceURL, string sUsername, string sPassword, string sWebProxy, int iWebProxyPort)
        {
            ExchangeWebServiceURL = sServiceURL;
            if (sWebProxy.Length > 0)
            {
                _userData.bUseProxy = true;
                _userData.WebProxy = sWebProxy;
                _userData.WebProxyPort = iWebProxyPort;
            }
            else{
                _userData.bUseProxy = false;
                _userData.WebProxy = "";
                _userData.WebProxyPort = 8080;
            }
        }

        public void start()
        {
            if (_service != null)
            {
                OnStateChanged(new StatusEventArgs(StatusType.error, "Already connected"));
                return;
            }
            try
            {
                serviceURL = ExchangeWebServiceURL;
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
                        
                        // If the certificate is a valid, signed certificate, return true.
                        if (errors == System.Net.Security.SslPolicyErrors.None)
                        {
                            OnStateChanged(new StatusEventArgs(StatusType.validating, "No error"));
                            return true;
                        }
                        // If there are errors in the certificate chain, look at each error to determine the cause.
                        if ((errors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
                        {
                            OnStateChanged(new StatusEventArgs(StatusType.validating, "validating certificate"));
                            if (chain != null && chain.ChainStatus != null)
                            {
                                foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                                {
                                    if ((certificate.Subject == certificate.Issuer) &&
                                       (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                                    {
                                        // Self-signed certificates with an untrusted root are valid. 
                                        continue;
                                    }
                                    else
                                    {
                                        if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                                        {
                                            // If there are any other errors in the certificate chain, the certificate is invalid,
                                            // so the method returns false.
                                            helpers.addLog("certificate not validated");
                                            OnStateChanged(new StatusEventArgs(StatusType.validating, "certificate not validated"));
                                            return false;
                                        }
                                    }
                                }
                            }

                            // When processing reaches this line, the only errors in the certificate chain are 
                            // untrusted root errors for self-signed certificates. These certificates are valid
                            // for default Exchange server installations, so return true.
                            helpers.addLog("certificate validated");
                            OnStateChanged(new StatusEventArgs(StatusType.validating, "certificate validated"));
                            return true;
                        }
                        else
                        {
                            // In all other cases, return false.
                            helpers.addLog("certificate not validated");
                            OnStateChanged(new StatusEventArgs(StatusType.validating, "certificate not validated"));
                            return false;
                        }
                        
                    };

                _service = new ExchangeService(Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2010_SP2);

                if (_enableTrace)
                {
                    _service.TraceListener = new TraceListener();// ITraceListenerInstance;
                    // Optional flags to indicate the requests and responses to trace.
                    _service.TraceFlags = TraceFlags.EwsRequest | TraceFlags.EwsResponse;
                    _service.TraceEnabled = true;
                }
            }
            catch (Exception ex)
            {
                helpers.addLog("Exception: " + ex.Message);
            }

        }

        public bool logon(string sDomain, string sUser, string sPassword, bool bProxy)
        {
            _userData.bUseProxy = bProxy;
            _userData.sDomain = sDomain;
            _userData.sUser = sUser;
            _userData.sPassword = sPassword;
            _userData.WebProxy = webProxy;
            _userData.WebProxyPort = webProxyPort;
            return logon(_userData.sDomain, _userData.sUser, _userData.sPassword, _userData.bUseProxy, _userData.WebProxy, _userData.WebProxyPort);
        }

        /// <summary>
        /// logon using domain, user and password
        /// </summary>
        /// <param name="sDomain">the domain for login</param>
        /// <param name="sUser">user name</param>
        /// <param name="sPassword">password</param>
        /// <returns>true if success</returns>
        public bool logon(string sDomain, string sUser, string sPassword, bool bProxy, string sWebProxy, int iWebProxyPort)
        {
            OnStateChanged(new StatusEventArgs(StatusType.busy, "logon..."));
            bool bRet = false;
            _userData = new UserData(sDomain, sUser, sPassword, bProxy, sWebProxy, iWebProxyPort);
            try
            {
                // Or use NetworkCredential directly (WebCredentials is a wrapper
                // around NetworkCredential).
                _service.Credentials = new NetworkCredential(sUser, sPassword, sDomain);
                if(_userData.bUseProxy)
                    _service.WebProxy = new WebProxy(_userData.WebProxy, _userData.WebProxyPort);
                _service.Url = new Uri(ExchangeWebServiceURL);
                helpers.addLog("connected to: " + _service.Url.ToString());
                OnStateChanged(new StatusEventArgs(StatusType.idle, "connected to: " + _service.Url.ToString()));
                
                OnStateChanged(new StatusEventArgs(StatusType.ews_started, "ews started"));

                bRet = true;
            }
            catch (Exception ex)
            {
                helpers.addLog("Exception: " + ex.Message);
                OnStateChanged(new StatusEventArgs(StatusType.error, "logon failed: " + ex.Message));
                OnStateChanged(new StatusEventArgs(StatusType.ews_stopped, "ews logon failed"));
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
            serviceURL = url; //temporary store the new service URL
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

            _bRunPullThread = false;
            if (_pullMails != null)
            {
                bool b = _pullMails.Join(1000);
                if (!b)
                    _pullMails.Abort();
            }
            OnStateChanged(new StatusEventArgs(StatusType.ews_stopped, "ews ended"));
        }

        #region Threading_stuff
        // see interfaces //public event StateChangedEventHandler stateChangedEvent;  
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
            const int chunkSize = 50;
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
                    SearchFilter.SearchFilterCollection filterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
                    filterCollection.Add(new SearchFilter.Not(new SearchFilter.ContainsSubstring(ItemSchema.Subject, sMailHasAlreadyProcessed)));
                    filterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, helpers.filterSubject));
                    filterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Attachments, helpers.filterAttachement));
                    findResults = _ews._service.FindItems(WellKnownFolderName.Inbox, filterCollection, view);

                    _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "getMail: found "+ findResults.Items.Count + " items in inbox"));
                    foreach (Item item in findResults.Items)
                    {
                        helpers.addLog("found item...");
                        if (item is EmailMessage)
                        {
                            EmailMessage mailmessage = item as EmailMessage;
                            mailmessage.Load(); //load data from server

                            helpers.addLog("\t is email ...");

                            // If the item is an e-mail message, write the sender's name.
                            helpers.addLog(mailmessage.Sender.Name + ": " + mailmessage.Subject);
                            _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "getMail: processing eMail " + mailmessage.Subject));

                            MailMsg myMailMsg = new MailMsg(mailmessage, _ews._userData.sUser);

                            // Bind to an existing message using its unique identifier.
                            //EmailMessage message = EmailMessage.Bind(service, new ItemId(item.Id.UniqueId));
                            _ews._licenseMail.processMail(myMailMsg);

                            //change subject?
                            // Bind to the existing item, using the ItemId. This method call results in a GetItem call to EWS.
                            Item myItem = Item.Bind(_ews._service, item.Id as ItemId);
                            myItem.Load();
                            // Update the Subject of the email.
                            myItem.Subject += "[processed]";
                            // Save the updated email. This method call results in an UpdateItem call to EWS.
                            myItem.Update(ConflictResolutionMode.AlwaysOverwrite);

                            _ews.OnStateChanged(new StatusEventArgs(StatusType.license_mail, "processMail"));
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
            _bRunPullThread = true;
            _pullMails = new Thread(startPullNotification);
            _pullMails.Name = "Pull thread";
            _pullMails.Start(this);
        }

        //pull notifications
        static void startPullNotification(object param)
        {
            ews _ews = (ews)param;  //need an instance
            _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "Pullnotification: started"));
            int iTimeoutMinutes = 5;

            // Subscribe to pull notifications in the Inbox folder, and get notified when
            // a new mail is received, when an item or folder is created, or when an item
            // or folder is deleted.
            PullSubscription subscription=null;
            try
            {
                subscription = _ews._service.SubscribeToPullNotifications(new FolderId[] { WellKnownFolderName.Inbox },
                iTimeoutMinutes /* timeOut: the subscription will end if the server is not polled within 5 minutes. */,
                null /* watermark: null to start a new subscription. */,
                EventType.NewMail);//, EventType.Created, EventType.Deleted);
            }
            catch (Exception ex)
            {
                _ews.OnStateChanged(new StatusEventArgs(StatusType.error, "Pullnotification: PullSubscription FAILED: " + ex.Message));
            }
            if (subscription == null)
            {
                _ews.OnStateChanged(new StatusEventArgs(StatusType.error, "Pullnotification: END with no subscription"));
                return;
            }
            try
            {
                while (_ews._bRunPullThread)
                {
                    _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "Pull sleeping "+(iTimeoutMinutes - 1).ToString()+" minutes..."));
                    int iCount=0;
                    do
                    {
                        _ews.OnStateChanged(new StatusEventArgs(StatusType.ews_pulse, "ews lives"));
                        iCount += 5000;
                        Thread.Sleep(5000);
                        //Thread.Sleep((iTimeoutMinutes - 1) * 60000);
                    } while (iCount < 60000);//do not sleep longer than iTimeout or you loose the subscription

                    // Wait a couple minutes, then poll the server for new events.   
                    _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "Pull looking for new mails"));
                    GetEventsResults events = subscription.GetEvents();
                    _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "Pull processing"));
                    // Loop through all item-related events.
                    foreach (ItemEvent itemEvent in events.ItemEvents)
                    {
                        switch (itemEvent.EventType)
                        {
                            case EventType.NewMail:
                                utils.helpers.addLog("PullNotification: " + "EventType.NewMail");
                                _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "check eMail ..."));
                                // A new mail has been received. Bind to it
                                EmailMessage message = EmailMessage.Bind(_ews._service, itemEvent.ItemId);
                                bool bDoProcessMail=false;
                                if (message.Subject.Contains(helpers.filterSubject) && (!message.Subject.Contains("[processed]")))
                                {
                                    if(message.HasAttachments){
                                        //check for endings of attachements
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
                                        _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "PullNotification: processMail"));
                                        //create a new IMailMessage from the EmailMessage
                                        MailMsg myMailMsg = new MailMsg(message, _ews._userData.sUser);
                                        _ews._licenseMail.processMail(myMailMsg);
                                        
                                        //change subject?
                                        // Bind to the existing item, using the ItemId. This method call results in a GetItem call to EWS.
                                        Item myItem = Item.Bind(_ews._service, itemEvent.ItemId);
                                        myItem.Load();
                                        // Update the Subject of the email.
                                        myItem.Subject += "[processed]";
                                        // Save the updated email. This method call results in an UpdateItem call to EWS.
                                        myItem.Update(ConflictResolutionMode.AlwaysOverwrite);
                                        
                                        _ews.OnStateChanged(new StatusEventArgs(StatusType.license_mail, "Pullnotification: email marked"));
                                    }
                                    else
                                        _ews.OnStateChanged(new StatusEventArgs(StatusType.other_mail, "PullNotification: mail does not match"));
                                }
                                break;
                            case EventType.Created:
                                // An item was created in the folder. Bind to it.
                                Item item = Item.Bind(_ews._service, itemEvent.ItemId);
                                utils.helpers.addLog("PullNotification: " + "EventType.Created " + item.ToString());
                                break;
                            case EventType.Deleted:
                                // An item has been deleted. Output its ID to the console.
                                utils.helpers.addLog("Item deleted: " + itemEvent.ItemId.UniqueId);
                                break;
                            default:
                                utils.helpers.addLog("PullNotification: " + itemEvent.EventType.ToString());
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
                                Folder folder = Folder.Bind(_ews._service, folderEvent.FolderId);
                                break;
                            case EventType.Deleted:
                                // A folder has been deleted. Output its Id to the console.
                                utils.helpers.addLog("PullNotification: " + folderEvent.FolderId.UniqueId);
                                break;
                            default:
                                utils.helpers.addLog("PullNotification: " + folderEvent.EventType.ToString());
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
                _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "PullNotification: ThreadAbortException"));
                _ews._bRunPullThread = false;
            }
            catch (Exception ex)
            {
                _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "PullNotification: Pull Thread Exception:" + ex.Message));
                _ews._bRunPullThread = false;
            }
            try
            {
                subscription.Unsubscribe();
            }
            catch (Exception ex)
            {
                utils.helpers.addExceptionLog("subscriptions unsubscribe exception");
                _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "PullNotification: subscriptions unsubscribe exception " + ex.Message));
            }
            subscription = null;
            _ews.OnStateChanged(new StatusEventArgs(StatusType.ews_stopped, "ews pull ended"));
            _ews.OnStateChanged(new StatusEventArgs(StatusType.none, "PullNotification: Pull ended"));
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
