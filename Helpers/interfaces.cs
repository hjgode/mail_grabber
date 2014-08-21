using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Helpers
{
    public delegate void StateChangedEventHandler(object sender, StatusEventArgs args);
    /// <summary>
    /// interface to be used for a mail 'client'
    /// </summary>
    public interface IMailHost:IDisposable
    {
        event StateChangedEventHandler StateChanged;
        bool logon(string sDomain, string sUser, string sPassword);
        void start();        
        void getMailsAsync();
        utils.userData UserData { get; set; }
    }

    /// <summary>
    /// interface for a mail message
    /// </summary>
    public interface IMailMessage
    {
        /// <summary>
        /// the user that processed this mail
        /// </summary>
        string User{get;set;}
        string Body { get; set; }
        string Subject { get; set; }
        Attachement[] Attachements { get; set; }
        void addAttachement(Attachement att);
        DateTime timestamp { get; set; }
    }

    /// <summary>
    /// interface for attachement store
    /// </summary>
    public interface IAttachement
    {
        System.IO.Stream data { get; set; }
        long size { get; set; }
        string name{ get; set; }
    }

    /// <summary>
    /// class that implements an attachement
    /// </summary>
    public class Attachement : IAttachement
    {
        public long size
        {
            get;
            set;
        }
        public System.IO.Stream data
        {
            get { return _stream; }
            set { _stream = value; }
        }
        System.IO.Stream _stream;
        public string name
        {
            get { return _name; }
            set { _name = value; }
        }
        string _name;
        public Attachement(System.IO.Stream s, string n)
        {
            this._stream = s;
            this._name = n;
            size = s.Length;
        }
    }
    public class MailMessage : IMailMessage
    {
        public string User { get; set; }
        public string Body { get; set; }
        public string Subject { get; set; }
        public Attachement[] Attachements { get; set; }
        List<Attachement> attList = new List<Attachement>();
        public DateTime timestamp{ get; set; }
        public MailMessage(string user, string body, string subject, Attachement[] atts, DateTime dt)
        {
            this.User = user;
            this.Body = body;
            this.Subject = subject;
            this.Attachements = atts;
            attList.AddRange(atts);
            timestamp = dt;
        }
        public void addAttachement(Attachement att)
        {
            attList.Add(att);
        }
    }

}
