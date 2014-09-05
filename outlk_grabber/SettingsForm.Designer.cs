namespace outlk_grabber
{
    partial class SettingsForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.txtProfile = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkEnableDialog = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txtDatabaseFile = new System.Windows.Forms.TextBox();
            this.btnGetFile = new System.Windows.Forms.Button();
            this.chkNewSession = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "User:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Profile:";
            // 
            // txtUser
            // 
            this.txtUser.Location = new System.Drawing.Point(133, 43);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(219, 20);
            this.txtUser.TabIndex = 1;
            // 
            // txtProfile
            // 
            this.txtProfile.Location = new System.Drawing.Point(133, 15);
            this.txtProfile.Name = "txtProfile";
            this.txtProfile.Size = new System.Drawing.Size(219, 20);
            this.txtProfile.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkNewSession);
            this.groupBox1.Controls.Add(this.chkEnableDialog);
            this.groupBox1.Location = new System.Drawing.Point(22, 146);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(352, 102);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Outlook Interop";
            // 
            // chkEnableDialog
            // 
            this.chkEnableDialog.AutoSize = true;
            this.chkEnableDialog.Location = new System.Drawing.Point(111, 31);
            this.chkEnableDialog.Name = "chkEnableDialog";
            this.chkEnableDialog.Size = new System.Drawing.Size(92, 17);
            this.chkEnableDialog.TabIndex = 4;
            this.chkEnableDialog.Text = "Enable Dialog";
            this.chkEnableDialog.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(296, 345);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(79, 30);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(22, 345);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(79, 30);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 92);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Database File:";
            // 
            // txtDatabaseFile
            // 
            this.txtDatabaseFile.Location = new System.Drawing.Point(133, 89);
            this.txtDatabaseFile.Name = "txtDatabaseFile";
            this.txtDatabaseFile.Size = new System.Drawing.Size(219, 20);
            this.txtDatabaseFile.TabIndex = 1;
            // 
            // btnGetFile
            // 
            this.btnGetFile.Location = new System.Drawing.Point(360, 89);
            this.btnGetFile.Name = "btnGetFile";
            this.btnGetFile.Size = new System.Drawing.Size(38, 22);
            this.btnGetFile.TabIndex = 4;
            this.btnGetFile.Text = "...";
            this.btnGetFile.UseVisualStyleBackColor = true;
            this.btnGetFile.Click += new System.EventHandler(this.btnGetFile_Click);
            // 
            // chkNewSession
            // 
            this.chkNewSession.AutoSize = true;
            this.chkNewSession.Location = new System.Drawing.Point(111, 65);
            this.chkNewSession.Name = "chkNewSession";
            this.chkNewSession.Size = new System.Drawing.Size(88, 17);
            this.chkNewSession.TabIndex = 4;
            this.chkNewSession.Text = "New Session";
            this.chkNewSession.UseVisualStyleBackColor = true;
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(404, 399);
            this.Controls.Add(this.btnGetFile);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.txtProfile);
            this.Controls.Add(this.txtDatabaseFile);
            this.Controls.Add(this.txtUser);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.Text = "SettingsForm";
            this.Activated += new System.EventHandler(this.SettingsForm_Activated);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtUser;
        private System.Windows.Forms.TextBox txtProfile;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkEnableDialog;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtDatabaseFile;
        private System.Windows.Forms.Button btnGetFile;
        private System.Windows.Forms.CheckBox chkNewSession;
    }
}