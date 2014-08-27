namespace ews_grabber
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuSettings = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exchangeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuConnect = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuDisconnect = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuRefresh = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuAdmin = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuTest_xml = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuTest_DB = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuProcess_Mail = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuClearData = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabBrowse = new System.Windows.Forms.TabPage();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnAll = new System.Windows.Forms.Button();
            this.btnSearch = new System.Windows.Forms.Button();
            this.txtFilter = new System.Windows.Forms.TextBox();
            this.radioDeviceID = new System.Windows.Forms.RadioButton();
            this.radioKeyNumber = new System.Windows.Forms.RadioButton();
            this.radioPOnumber = new System.Windows.Forms.RadioButton();
            this.radioOrderNumber = new System.Windows.Forms.RadioButton();
            this.tabLog = new System.Windows.Forms.TabPage();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lblLED = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dataPopup = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuExport = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabBrowse.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.tabLog.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.dataPopup.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.exchangeToolStripMenuItem,
            this.mnuRefresh,
            this.mnuAdmin});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(603, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuSettings,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // mnuSettings
            // 
            this.mnuSettings.Name = "mnuSettings";
            this.mnuSettings.Size = new System.Drawing.Size(116, 22);
            this.mnuSettings.Text = "Settings";
            this.mnuSettings.Click += new System.EventHandler(this.mnuSettings_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // exchangeToolStripMenuItem
            // 
            this.exchangeToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuConnect,
            this.mnuDisconnect});
            this.exchangeToolStripMenuItem.Name = "exchangeToolStripMenuItem";
            this.exchangeToolStripMenuItem.Size = new System.Drawing.Size(69, 20);
            this.exchangeToolStripMenuItem.Text = "Exchange";
            // 
            // mnuConnect
            // 
            this.mnuConnect.Name = "mnuConnect";
            this.mnuConnect.Size = new System.Drawing.Size(133, 22);
            this.mnuConnect.Text = "Connect";
            this.mnuConnect.Click += new System.EventHandler(this.mnuConnect_Click);
            // 
            // mnuDisconnect
            // 
            this.mnuDisconnect.Enabled = false;
            this.mnuDisconnect.Name = "mnuDisconnect";
            this.mnuDisconnect.Size = new System.Drawing.Size(133, 22);
            this.mnuDisconnect.Text = "Disconnect";
            this.mnuDisconnect.Click += new System.EventHandler(this.mnuDisconnect_Click);
            // 
            // mnuRefresh
            // 
            this.mnuRefresh.Name = "mnuRefresh";
            this.mnuRefresh.Size = new System.Drawing.Size(58, 20);
            this.mnuRefresh.Text = "Refresh";
            this.mnuRefresh.Click += new System.EventHandler(this.mnuRefresh_Click);
            // 
            // mnuAdmin
            // 
            this.mnuAdmin.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuTest_xml,
            this.mnuTest_DB,
            this.mnuProcess_Mail,
            this.mnuClearData});
            this.mnuAdmin.Name = "mnuAdmin";
            this.mnuAdmin.Size = new System.Drawing.Size(55, 20);
            this.mnuAdmin.Text = "Admin";
            // 
            // mnuTest_xml
            // 
            this.mnuTest_xml.Name = "mnuTest_xml";
            this.mnuTest_xml.Size = new System.Drawing.Size(162, 22);
            this.mnuTest_xml.Text = "Test xml";
            this.mnuTest_xml.Click += new System.EventHandler(this.mnuTest_xml_Click);
            // 
            // mnuTest_DB
            // 
            this.mnuTest_DB.Name = "mnuTest_DB";
            this.mnuTest_DB.Size = new System.Drawing.Size(162, 22);
            this.mnuTest_DB.Text = "Test db";
            this.mnuTest_DB.Click += new System.EventHandler(this.mnuTest_DB_Click);
            // 
            // mnuProcess_Mail
            // 
            this.mnuProcess_Mail.Name = "mnuProcess_Mail";
            this.mnuProcess_Mail.Size = new System.Drawing.Size(162, 22);
            this.mnuProcess_Mail.Text = "Test processMail";
            this.mnuProcess_Mail.Click += new System.EventHandler(this.mnuProcess_Mail_Click);
            // 
            // mnuClearData
            // 
            this.mnuClearData.Name = "mnuClearData";
            this.mnuClearData.Size = new System.Drawing.Size(162, 22);
            this.mnuClearData.Text = "Clear data";
            this.mnuClearData.Click += new System.EventHandler(this.mnuClearData_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabBrowse);
            this.tabControl1.Controls.Add(this.tabLog);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 24);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(603, 435);
            this.tabControl1.TabIndex = 2;
            // 
            // tabBrowse
            // 
            this.tabBrowse.Controls.Add(this.dataGridView1);
            this.tabBrowse.Controls.Add(this.panel1);
            this.tabBrowse.Location = new System.Drawing.Point(4, 22);
            this.tabBrowse.Name = "tabBrowse";
            this.tabBrowse.Size = new System.Drawing.Size(595, 409);
            this.tabBrowse.TabIndex = 2;
            this.tabBrowse.Text = "Browse";
            this.tabBrowse.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.ContextMenuStrip = this.dataPopup;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 67);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(595, 342);
            this.dataGridView1.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnExport);
            this.panel1.Controls.Add(this.btnAll);
            this.panel1.Controls.Add(this.btnSearch);
            this.panel1.Controls.Add(this.txtFilter);
            this.panel1.Controls.Add(this.radioDeviceID);
            this.panel1.Controls.Add(this.radioKeyNumber);
            this.panel1.Controls.Add(this.radioPOnumber);
            this.panel1.Controls.Add(this.radioOrderNumber);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(595, 67);
            this.panel1.TabIndex = 1;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(510, 7);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(77, 24);
            this.btnExport.TabIndex = 2;
            this.btnExport.Text = "EXPORT";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnAll
            // 
            this.btnAll.Location = new System.Drawing.Point(395, 37);
            this.btnAll.Name = "btnAll";
            this.btnAll.Size = new System.Drawing.Size(77, 24);
            this.btnAll.TabIndex = 2;
            this.btnAll.Text = "No Filter";
            this.btnAll.UseVisualStyleBackColor = true;
            this.btnAll.Click += new System.EventHandler(this.btnAll_Click);
            // 
            // btnSearch
            // 
            this.btnSearch.Location = new System.Drawing.Point(395, 6);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(77, 25);
            this.btnSearch.TabIndex = 2;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // txtFilter
            // 
            this.txtFilter.Location = new System.Drawing.Point(8, 9);
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.Size = new System.Drawing.Size(187, 20);
            this.txtFilter.TabIndex = 1;
            // 
            // radioDeviceID
            // 
            this.radioDeviceID.AutoSize = true;
            this.radioDeviceID.Location = new System.Drawing.Point(304, 10);
            this.radioDeviceID.Name = "radioDeviceID";
            this.radioDeviceID.Size = new System.Drawing.Size(73, 17);
            this.radioDeviceID.TabIndex = 0;
            this.radioDeviceID.TabStop = true;
            this.radioDeviceID.Text = "Device ID";
            this.radioDeviceID.UseVisualStyleBackColor = true;
            // 
            // radioKeyNumber
            // 
            this.radioKeyNumber.AutoSize = true;
            this.radioKeyNumber.Location = new System.Drawing.Point(304, 33);
            this.radioKeyNumber.Name = "radioKeyNumber";
            this.radioKeyNumber.Size = new System.Drawing.Size(81, 17);
            this.radioKeyNumber.TabIndex = 0;
            this.radioKeyNumber.TabStop = true;
            this.radioKeyNumber.Text = "Key number";
            this.radioKeyNumber.UseVisualStyleBackColor = true;
            // 
            // radioPOnumber
            // 
            this.radioPOnumber.AutoSize = true;
            this.radioPOnumber.Location = new System.Drawing.Point(201, 33);
            this.radioPOnumber.Name = "radioPOnumber";
            this.radioPOnumber.Size = new System.Drawing.Size(97, 17);
            this.radioPOnumber.TabIndex = 0;
            this.radioPOnumber.TabStop = true;
            this.radioPOnumber.Text = "Purchase order";
            this.radioPOnumber.UseVisualStyleBackColor = true;
            // 
            // radioOrderNumber
            // 
            this.radioOrderNumber.AutoSize = true;
            this.radioOrderNumber.Location = new System.Drawing.Point(201, 10);
            this.radioOrderNumber.Name = "radioOrderNumber";
            this.radioOrderNumber.Size = new System.Drawing.Size(89, 17);
            this.radioOrderNumber.TabIndex = 0;
            this.radioOrderNumber.TabStop = true;
            this.radioOrderNumber.Text = "Order number";
            this.radioOrderNumber.UseVisualStyleBackColor = true;
            // 
            // tabLog
            // 
            this.tabLog.Controls.Add(this.textBox1);
            this.tabLog.Location = new System.Drawing.Point(4, 22);
            this.tabLog.Name = "tabLog";
            this.tabLog.Padding = new System.Windows.Forms.Padding(3);
            this.tabLog.Size = new System.Drawing.Size(595, 409);
            this.tabLog.TabIndex = 1;
            this.tabLog.Text = "Log";
            this.tabLog.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox1.Location = new System.Drawing.Point(3, 3);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(589, 403);
            this.textBox1.TabIndex = 1;
            // 
            // lblLED
            // 
            this.lblLED.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblLED.AutoSize = true;
            this.lblLED.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.lblLED.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.lblLED.Location = new System.Drawing.Point(3, 6);
            this.lblLED.Name = "lblLED";
            this.lblLED.Size = new System.Drawing.Size(15, 13);
            this.lblLED.TabIndex = 4;
            this.lblLED.Text = "O";
            // 
            // lblStatus
            // 
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblStatus.Location = new System.Drawing.Point(24, 3);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.ReadOnly = true;
            this.lblStatus.Size = new System.Drawing.Size(576, 20);
            this.lblStatus.TabIndex = 5;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.lblLED, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.lblStatus, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 433);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.Size = new System.Drawing.Size(603, 26);
            this.tableLayoutPanel1.TabIndex = 6;
            // 
            // dataPopup
            // 
            this.dataPopup.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuExport});
            this.dataPopup.Name = "dataPopup";
            this.dataPopup.Size = new System.Drawing.Size(108, 26);
            // 
            // mnuExport
            // 
            this.mnuExport.Name = "mnuExport";
            this.mnuExport.Size = new System.Drawing.Size(152, 22);
            this.mnuExport.Text = "Export";
            this.mnuExport.Click += new System.EventHandler(this.mnuExport_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(603, 459);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Exchange Test Mail Grabber";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabBrowse.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabLog.ResumeLayout(false);
            this.tabLog.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.dataPopup.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exchangeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mnuConnect;
        private System.Windows.Forms.ToolStripMenuItem mnuAdmin;
        private System.Windows.Forms.ToolStripMenuItem mnuTest_xml;
        private System.Windows.Forms.ToolStripMenuItem mnuTest_DB;
        private System.Windows.Forms.ToolStripMenuItem mnuProcess_Mail;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabBrowse;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TabPage tabLog;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ToolStripMenuItem mnuClearData;
        private System.Windows.Forms.ToolStripMenuItem mnuRefresh;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnAll;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.TextBox txtFilter;
        private System.Windows.Forms.RadioButton radioDeviceID;
        private System.Windows.Forms.RadioButton radioKeyNumber;
        private System.Windows.Forms.RadioButton radioPOnumber;
        private System.Windows.Forms.RadioButton radioOrderNumber;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.ToolStripMenuItem mnuSettings;
        private System.Windows.Forms.ToolStripMenuItem mnuDisconnect;
        private System.Windows.Forms.Label lblLED;
        private System.Windows.Forms.TextBox lblStatus;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ContextMenuStrip dataPopup;
        private System.Windows.Forms.ToolStripMenuItem mnuExport;
    }
}

