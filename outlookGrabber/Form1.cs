using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace outlookGrabber
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            OutlookClass oc = new OutlookClass();
            textBox1.Text += oc.getMails();            
            textBox1.SelectionStart = 0;
            textBox1.SelectionLength = 0;
        }
    }
}
