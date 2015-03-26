using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ews_grabber
{
    public partial class ManualDataInput : Form
    {
        Helpers.LicenseDataBase _licenseDataBase;
        MySettings _settings;

        public ManualDataInput(ref Helpers.LicenseDataBase licData, ref MySettings Settings)
        {
            InitializeComponent();

            txtProduct.Text = "[NAU-1504] CETerm for Windows CE 6.0 / 5.0 / CE .NET";
            _licenseDataBase = licData;
            attachErrorProvider();
            btnSave.Enabled = false;

            // Create the list to use as the custom source. 
            var source = new AutoCompleteStringCollection();
            source.AddRange(_licenseDataBase.productList);
            txtProduct.AutoCompleteCustomSource = source;
            txtProduct.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtProduct.AutoCompleteSource = AutoCompleteSource.CustomSource;

            //prefill
            _settings = Settings;
            //settings.load();
            txtReceivedBy.Text = _settings.ExchangeUsername;
        }

        void attachErrorProvider()
        {
            foreach (Control c in this.Controls)
            {
                if(c.GetType()==typeof(TextBox)){
                    ((TextBox)c).Validated += new EventHandler(textBox_Validated);
                }
            }
        }
        
        public void textBox_Validated(object sender, EventArgs e)
        {
            var tb = (TextBox)sender;
            if (tb.Name == "txtID")
                return;

            btnSave.Enabled=true;
            if (string.IsNullOrEmpty(tb.Text))
            {
                errorProvider1.SetError(tb, "Please fill valid data");
                btnSave.Enabled = false;
            }
            else
                errorProvider1.SetError(tb, "");                
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //chceck for errors first

            string deviceid=txtDeviceID.Text;
            string customer=txtCustomer.Text;
            string key=txtLicenseKey.Text;
            string ordernumber=txtOrderNumber.Text;
            
            DateTime orderdate=dtPickerOrderDate.Value;
            
            string ponumber=txtPONumber.Text;
            string endcustomer=txtEndCustomer.Text;
            string product=txtProduct.Text;

            int i,quantity;
            if(int.TryParse(txtQuantity.Text, out i))
                quantity=int.Parse(txtQuantity.Text);
            else{
                errorProvider1.SetError(txtQuantity, "Not a number");
                return;
            }
            string receivedby=txtReceivedBy.Text;

            DateTime sendat = dtPickerSendAt.Value;

            int iRes = _licenseDataBase.add(deviceid, customer, key, ordernumber, orderdate, ponumber, endcustomer, product, quantity, receivedby, sendat);
            if (iRes == -1)
            {
                MessageBox.Show("Data already exists for this key and device");
                lblLog.Text="Data already exists for this key and device:\r\nDevice ID: "+deviceid+"\r\nLicense Key: "+key;
            }
            else
            {
                MessageBox.Show("Data saved with id=" + iRes);
                Helpers.LicenseData licData = new Helpers.LicenseData(deviceid, customer, key, ordernumber, orderdate, ponumber, endcustomer, product, quantity, receivedby, sendat);
                lblLog.Text = "SAVED\r\n"+licData.ToString();
            }
            //clear all textboxes?
            
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            foreach (Control c in this.Controls)
            {
                if (c.GetType() == typeof(TextBox))
                {
                    ((TextBox)c).Text = "";
                }
            }

        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            txtDeviceID.Text = "";
            txtID.Text = "";
            txtLicenseKey.Text = "";
            txtReceivedBy.Text = _settings.ExchangeUsername;
            txtProduct.Text = "[NAU-1504] CETerm for Windows CE 6.0 / 5.0 / CE .NET";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
    }
}
