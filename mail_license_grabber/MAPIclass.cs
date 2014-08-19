using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ManagedMAPI;
using System.Runtime.InteropServices;

namespace mail_license_grabber
{
    class MAPIclass
    {
        ManagedMAPI.MAPISession session=null;

        public MAPIclass()
        {
            try
            {
                session = new MAPISession();
                Dictionary<string, bool> msgStores = session.GetMessageStores();
                foreach (KeyValuePair<string, bool> e in msgStores)
                {
                    helpers.addLog(e.Key + "/" + e.Value.ToString());
                }
            }
            catch (Exception ex)
            {
                helpers.addLog("Exception: " + ex.Message);
            }
        }
    }

    public class helpers
    {
        public static void addLog(string s)
        {
            System.Diagnostics.Debug.WriteLine(s);
        }
    }
}
