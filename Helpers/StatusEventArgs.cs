using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Helpers
{
    public enum StatusType
    {
        none,
        idle,
        busy,
        success,
        error,
        url_changed,
        validating,
        license_mail,
        other_mail
    }

    public class StatusEventArgs:EventArgs
    {
        public StatusType eStatus;
        public string strMessage;
        public IMailMessage mailmsg { 
            get { return myMailMsg; }
            private set { myMailMsg = value; }
        }
        IMailMessage myMailMsg;
        public StatusEventArgs(StatusType state)
        {
            eStatus = state;
        }
        public StatusEventArgs(StatusType state, string msg)
        {
            eStatus = state;
            strMessage = msg;
        }
        public StatusEventArgs(StatusType state, IMailMessage msg)
        {
            eStatus = state;
            strMessage = msg.Subject;
            myMailMsg = msg;
        }
    }
}
