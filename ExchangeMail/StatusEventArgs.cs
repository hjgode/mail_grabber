using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeMail
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
        license_mail
    }

    public class StatusEventArgs:EventArgs
    {
        public StatusType eStatus;
        public string message;
        public StatusEventArgs(StatusType state)
        {
            eStatus = state;
        }
        public Microsoft.Exchange.WebServices.Data.EmailMessage excMail;
        public StatusEventArgs(StatusType state, string msg)
        {
            eStatus = state;
            message = msg;
            excMail = null;
        }
        public StatusEventArgs(StatusType state, Microsoft.Exchange.WebServices.Data.EmailMessage email)
        {
            eStatus = state;
            message = "new mail";
            excMail = email;
        }

    }
}
