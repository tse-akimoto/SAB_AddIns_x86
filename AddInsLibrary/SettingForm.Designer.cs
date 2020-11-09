namespace AddInsLibrary
{
    partial class SettingForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingForm));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnRegist = new System.Windows.Forms.Button();
            this.pnlStamp = new System.Windows.Forms.Panel();
            this.lblAlpha = new System.Windows.Forms.Label();
            this.nudAlpha = new System.Windows.Forms.NumericUpDown();
            this.chkChange = new System.Windows.Forms.CheckBox();
            this.lblDisplay = new System.Windows.Forms.Label();
            this.lblStamp = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.pnlSAB = new System.Windows.Forms.Panel();
            this.lblLink = new System.Windows.Forms.LinkLabel();
            this.rdoElse = new System.Windows.Forms.RadioButton();
            this.rdoB = new System.Windows.Forms.RadioButton();
            this.rdoA = new System.Windows.Forms.RadioButton();
            this.rdoS = new System.Windows.Forms.RadioButton();
            this.lblSABSetting = new System.Windows.Forms.Label();
            this.lblSettingLabel = new System.Windows.Forms.Label();
            this.lblSecrecy = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.pnlStamp.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudAlpha)).BeginInit();
            this.pnlSAB.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.Color.Green;
            resources.ApplyResources(this.btnClose, "btnClose");
            this.btnClose.ForeColor = System.Drawing.Color.White;
            this.btnClose.Name = "btnClose";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnRegist
            // 
            this.btnRegist.BackColor = System.Drawing.Color.Green;
            resources.ApplyResources(this.btnRegist, "btnRegist");
            this.btnRegist.ForeColor = System.Drawing.Color.White;
            this.btnRegist.Name = "btnRegist";
            this.btnRegist.UseVisualStyleBackColor = false;
            this.btnRegist.Click += new System.EventHandler(this.btnRegist_Click);
            // 
            // pnlStamp
            // 
            this.pnlStamp.BackColor = System.Drawing.Color.White;
            this.pnlStamp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlStamp.Controls.Add(this.lblAlpha);
            this.pnlStamp.Controls.Add(this.nudAlpha);
            this.pnlStamp.Controls.Add(this.chkChange);
            this.pnlStamp.Controls.Add(this.lblDisplay);
            this.pnlStamp.Controls.Add(this.lblStamp);
            this.pnlStamp.Controls.Add(this.label1);
            resources.ApplyResources(this.pnlStamp, "pnlStamp");
            this.pnlStamp.Name = "pnlStamp";
            // 
            // lblAlpha
            // 
            resources.ApplyResources(this.lblAlpha, "lblAlpha");
            this.lblAlpha.Name = "lblAlpha";
            // 
            // nudAlpha
            // 
            resources.ApplyResources(this.nudAlpha, "nudAlpha");
            this.nudAlpha.Maximum = new decimal(new int[] {
            95,
            0,
            0,
            0});
            this.nudAlpha.Name = "nudAlpha";
            this.nudAlpha.Value = new decimal(new int[] {
            50,
            0,
            0,
            0});
            // 
            // chkChange
            // 
            resources.ApplyResources(this.chkChange, "chkChange");
            this.chkChange.BackColor = System.Drawing.Color.Green;
            this.chkChange.Checked = true;
            this.chkChange.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkChange.ForeColor = System.Drawing.Color.White;
            this.chkChange.Name = "chkChange";
            this.chkChange.UseVisualStyleBackColor = false;
            this.chkChange.CheckedChanged += new System.EventHandler(this.btnChange_CheckedChanged);
            // 
            // lblDisplay
            // 
            this.lblDisplay.BackColor = System.Drawing.Color.WhiteSmoke;
            this.lblDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.lblDisplay, "lblDisplay");
            this.lblDisplay.Name = "lblDisplay";
            // 
            // lblStamp
            // 
            resources.ApplyResources(this.lblStamp, "lblStamp");
            this.lblStamp.Name = "lblStamp";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // pnlSAB
            // 
            this.pnlSAB.BackColor = System.Drawing.Color.White;
            this.pnlSAB.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlSAB.Controls.Add(this.lblLink);
            this.pnlSAB.Controls.Add(this.rdoElse);
            this.pnlSAB.Controls.Add(this.rdoB);
            this.pnlSAB.Controls.Add(this.rdoA);
            this.pnlSAB.Controls.Add(this.rdoS);
            this.pnlSAB.Controls.Add(this.lblSABSetting);
            this.pnlSAB.Controls.Add(this.lblSettingLabel);
            this.pnlSAB.Controls.Add(this.lblSecrecy);
            resources.ApplyResources(this.pnlSAB, "pnlSAB");
            this.pnlSAB.Name = "pnlSAB";
            // 
            // lblLink
            // 
            resources.ApplyResources(this.lblLink, "lblLink");
            this.lblLink.Name = "lblLink";
            this.lblLink.TabStop = true;
            this.lblLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // rdoElse
            // 
            resources.ApplyResources(this.rdoElse, "rdoElse");
            this.rdoElse.BackColor = System.Drawing.Color.Gray;
            this.rdoElse.ForeColor = System.Drawing.Color.White;
            this.rdoElse.Name = "rdoElse";
            this.rdoElse.UseVisualStyleBackColor = false;
            this.rdoElse.CheckedChanged += new System.EventHandler(this.btnSAB_CheckedChanged);
            // 
            // rdoB
            // 
            resources.ApplyResources(this.rdoB, "rdoB");
            this.rdoB.BackColor = System.Drawing.Color.Gray;
            this.rdoB.ForeColor = System.Drawing.Color.White;
            this.rdoB.Name = "rdoB";
            this.rdoB.UseVisualStyleBackColor = false;
            this.rdoB.CheckedChanged += new System.EventHandler(this.btnSAB_CheckedChanged);
            // 
            // rdoA
            // 
            resources.ApplyResources(this.rdoA, "rdoA");
            this.rdoA.BackColor = System.Drawing.Color.Gray;
            this.rdoA.ForeColor = System.Drawing.Color.White;
            this.rdoA.Name = "rdoA";
            this.rdoA.UseVisualStyleBackColor = false;
            this.rdoA.CheckedChanged += new System.EventHandler(this.btnSAB_CheckedChanged);
            // 
            // rdoS
            // 
            resources.ApplyResources(this.rdoS, "rdoS");
            this.rdoS.BackColor = System.Drawing.Color.Green;
            this.rdoS.Checked = true;
            this.rdoS.ForeColor = System.Drawing.Color.White;
            this.rdoS.Name = "rdoS";
            this.rdoS.TabStop = true;
            this.rdoS.UseVisualStyleBackColor = false;
            this.rdoS.CheckedChanged += new System.EventHandler(this.btnSAB_CheckedChanged);
            // 
            // lblSABSetting
            // 
            this.lblSABSetting.BackColor = System.Drawing.Color.WhiteSmoke;
            this.lblSABSetting.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.lblSABSetting, "lblSABSetting");
            this.lblSABSetting.Name = "lblSABSetting";
            // 
            // lblSettingLabel
            // 
            resources.ApplyResources(this.lblSettingLabel, "lblSettingLabel");
            this.lblSettingLabel.Name = "lblSettingLabel";
            // 
            // lblSecrecy
            // 
            resources.ApplyResources(this.lblSecrecy, "lblSecrecy");
            this.lblSecrecy.Name = "lblSecrecy";
            // 
            // lblTitle
            // 
            this.lblTitle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.lblTitle.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            resources.ApplyResources(this.lblTitle, "lblTitle");
            this.lblTitle.Name = "lblTitle";
            // 
            // lblLanguage
            // 
            resources.ApplyResources(this.lblLanguage, "lblLanguage");
            this.lblLanguage.Name = "lblLanguage";
            // 
            // SettingForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.Controls.Add(this.lblLanguage);
            this.Controls.Add(this.pnlStamp);
            this.Controls.Add(this.pnlSAB);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnRegist);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormSetting_FormClosing);
            this.Load += new System.EventHandler(this.FormSetting_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FormSetting_KeyDown);
            this.pnlStamp.ResumeLayout(false);
            this.pnlStamp.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudAlpha)).EndInit();
            this.pnlSAB.ResumeLayout(false);
            this.pnlSAB.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Panel pnlStamp;
        private System.Windows.Forms.Label lblDisplay;
        private System.Windows.Forms.Label lblStamp;
        private System.Windows.Forms.Panel pnlSAB;
        private System.Windows.Forms.Label lblSABSetting;
        private System.Windows.Forms.Label lblSettingLabel;
        private System.Windows.Forms.Label lblSecrecy;
        private System.Windows.Forms.Label lblTitle;
        protected System.Windows.Forms.Button btnClose;
        protected System.Windows.Forms.Button btnRegist;
        protected System.Windows.Forms.CheckBox chkChange;
        protected System.Windows.Forms.RadioButton rdoElse;
        protected System.Windows.Forms.RadioButton rdoB;
        protected System.Windows.Forms.RadioButton rdoA;
        protected System.Windows.Forms.RadioButton rdoS;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblAlpha;
        protected System.Windows.Forms.NumericUpDown nudAlpha;
        private System.Windows.Forms.LinkLabel lblLink;
    }
}