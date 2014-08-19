using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ews_grabber
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

    public class StatusEventArgs
    {
        public StatusType eStatus;
        public string message;
        public StatusEventArgs(StatusType state)
        {
            eStatus = state;
        }
        public StatusEventArgs(StatusType state, string msg)
        {
            eStatus = state;
            message = msg;
        }

    }
}
