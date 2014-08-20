using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Helpers
{
    public interface IMailMessage
    {
        string Sender{get;set;}
        string Body { get; set; }
        string Subject { get; set; }
        excAttachement[] Attachements { get; set; }
        //object blob { get; set; }
        string id { get; set; }
    }
    public interface IAttachement
    {
        object blob { get; set; }
        string name{ get; set; }
    }
    public interface IMailHost:IDisposable
    {
        //public delegate void StateChangedEventHandler(object sender, StatusEventArgs args);
        //public event StateChangedEventHandler stateChangedEvent;
        bool logon(string sDomain, string sUser, string sPassword);
        void start();
        void OnStateChanged(StatusEventArgs args);
        void getMailsAsync();
    }
    public class excAttachement : IAttachement
    {
        public object blob
        {
            get { return _blob; }
            set { _blob = value; }
        }
        object _blob;
        public string name
        {
            get { return _name; }
            set { _name = value; }
        }
        string _name;
        public excAttachement(object o, string n)
        {
            this._blob = o;
            this._name = n;
        }
    }

}
