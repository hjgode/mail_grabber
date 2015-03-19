using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Helpers
{
    public class LicenseData
    {
        #region FIELDS
        /// <summary>
        /// the deviceID of this license
        /// </summary>
        public string _deviceid;
        /// <summary>
        /// the customer for the license
        /// </summary>
        public string _customer;
        /// <summary>
        /// the key for this device
        /// </summary>
        public string _key;
        /// <summary>
        /// the order number
        /// </summary>
        public string _ordernumber;
        /// <summary>
        /// the order date of this purchase
        /// </summary>
        public DateTime _orderdate;
        /// <summary>
        /// the purchase order number
        /// </summary>
        public string _ponumber;
        /// <summary>
        /// the end customer of this license
        /// </summary>
        public string _endcustomer;
        /// <summary>
        /// the product valid for the license
        /// </summary>
        public string _product;
        /// <summary>
        /// number of licenses send as bundle
        /// </summary>
        public int _quantity;
        /// <summary>
        /// who grabbed this license data off exchange/outlook
        /// </summary>
        public string _receivedby;
        /// <summary>
        /// when was this license email sent
        /// </summary>
        public DateTime _sendat;
        #endregion

        public LicenseData(
                        string deviceid,
            string customer,
            string key,
            string ordernumber,
            DateTime orderdate,
            string ponumber,
            string endcustomer,
            string product,
            int quantity,
            string receivedby,
            DateTime sendat)
        {
            _deviceid = deviceid;
            _customer = customer;
            _key = key;
            _ordernumber = ordernumber;
            _orderdate = orderdate;
            _ponumber = ponumber;
            _endcustomer = endcustomer;
            _product = product;
            _quantity = quantity;
            _receivedby = receivedby;
            _sendat = sendat;
        }
        public LicenseData()
        {
        }
        public override string ToString()
        {
            string s = string.Format("deviceID: {0}\r\nCustomer: {1}\r\nLicense Key:{2}\r\nOrder Number: {3}\r\nOrder Date: {4}\r\n" +
                "PO Number: {5}\r\nEnd Customer: {6}\r\nProduct: {7}\r\nQuantity: {8}\r\nReceived by: {9}\r\nSend at: {10}",
                _deviceid,
                _customer,
                _key, _ordernumber, _orderdate, _ponumber, _endcustomer, _product, _quantity, _receivedby, _sendat);
            return s;
        }
    }
}
