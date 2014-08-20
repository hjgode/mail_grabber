using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace outlookGrabber_red
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            OutlookRedemptionClass oc = new OutlookRedemptionClass();
            if (oc.logon())
            {
                textBox1.Text = oc.getMails();
            }
            
        }
    }
}
