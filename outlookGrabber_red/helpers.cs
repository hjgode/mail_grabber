using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Diagnostics;

namespace utils
{
    public class helpers
    {
        public static void addLog(string s)
        {
            System.Diagnostics.Debug.WriteLine(s);
        }
        public static void addExceptionLog(string s)
        {
            // Get call stack
            StackTrace stackTrace = new StackTrace();
            System.Diagnostics.Debug.WriteLine("Exception in '"+ stackTrace.GetFrame(1).GetMethod().Name +": " + s);
        }
    }
}
