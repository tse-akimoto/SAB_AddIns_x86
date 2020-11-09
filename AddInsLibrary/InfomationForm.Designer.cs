namespace AddInsLibrary
{
    partial class InfomationForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InfomationForm));
            this.lblSecrecy = new System.Windows.Forms.Label();
            this.lblSABSetting = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblSecrecy
            // 
            this.lblSecrecy.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            resources.ApplyResources(this.lblSecrecy, "lblSecrecy");
            this.lblSecrecy.Name = "lblSecrecy";
            // 
            // lblSABSetting
            // 
            this.lblSABSetting.BackColor = System.Drawing.Color.WhiteSmoke;
            this.lblSABSetting.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.lblSABSetting, "lblSABSetting");
            this.lblSABSetting.Name = "lblSABSetting";
            // 
            // btnClose
            // 
            resources.ApplyResources(this.btnClose, "btnClose");
            this.btnClose.BackColor = System.Drawing.Color.Green;
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.ForeColor = System.Drawing.Color.White;
            this.btnClose.Name = "btnClose";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnEdit
            // 
            resources.ApplyResources(this.btnEdit, "btnEdit");
            this.btnEdit.BackColor = System.Drawing.Color.Green;
            this.btnEdit.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnEdit.ForeColor = System.Drawing.Color.White;
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.UseVisualStyleBackColor = false;
            this.btnEdit.Click += new System.EventHandler(this.btnRegist_Click);
            // 
            // InfomationForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.lblSABSetting);
            this.Controls.Add(this.lblSecrecy);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InfomationForm";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label lblSecrecy;
        private System.Windows.Forms.Label lblSABSetting;
        protected System.Windows.Forms.Button btnClose;
        protected System.Windows.Forms.Button btnEdit;
    }
}