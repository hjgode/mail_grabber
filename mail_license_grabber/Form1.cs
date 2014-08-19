using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace mail_license_grabber
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            MAPIclass mc = new MAPIclass();
        }
    }
}
