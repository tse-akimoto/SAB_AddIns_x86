namespace OutlookAddInSAB
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
            this.lviSuperiorsList = new System.Windows.Forms.ListView();
            this.columnHeaderName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderDeployment = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderTitle = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.tbxCCList = new System.Windows.Forms.TextBox();
            this.btnCC = new System.Windows.Forms.Button();
            this.btnSearch = new System.Windows.Forms.Button();
            this.tbxSearchText = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // lviSuperiorsList
            // 
            this.lviSuperiorsList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderName,
            this.columnHeaderDeployment,
            this.columnHeaderTitle});
            this.lviSuperiorsList.FullRowSelect = true;
            this.lviSuperiorsList.Location = new System.Drawing.Point(12, 36);
            this.lviSuperiorsList.Name = "lviSuperiorsList";
            this.lviSuperiorsList.Size = new System.Drawing.Size(510, 225);
            this.lviSuperiorsList.TabIndex = 2;
            this.lviSuperiorsList.UseCompatibleStateImageBehavior = false;
            this.lviSuperiorsList.View = System.Windows.Forms.View.Details;
            this.lviSuperiorsList.DoubleClick += new System.EventHandler(this.lviSuperiorsList_DoubleClick);
            // 
            // columnHeaderName
            // 
            this.columnHeaderName.Text = "名前";
            this.columnHeaderName.Width = 165;
            // 
            // columnHeaderDeployment
            // 
            this.columnHeaderDeployment.Text = "部署";
            this.columnHeaderDeployment.Width = 201;
            // 
            // columnHeaderTitle
            // 
            this.columnHeaderTitle.Text = "役職";
            this.columnHeaderTitle.Width = 140;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(432, 292);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(90, 21);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(320, 292);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(90, 21);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // tbxCCList
            // 
            this.tbxCCList.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.tbxCCList.Location = new System.Drawing.Point(108, 267);
            this.tbxCCList.Name = "tbxCCList";
            this.tbxCCList.ReadOnly = true;
            this.tbxCCList.Size = new System.Drawing.Size(414, 19);
            this.tbxCCList.TabIndex = 4;
            // 
            // btnCC
            // 
            this.btnCC.Location = new System.Drawing.Point(12, 265);
            this.btnCC.Name = "btnCC";
            this.btnCC.Size = new System.Drawing.Size(90, 21);
            this.btnCC.TabIndex = 3;
            this.btnCC.Text = "CC";
            this.btnCC.UseVisualStyleBackColor = true;
            this.btnCC.Click += new System.EventHandler(this.btnCC_Click);
            // 
            // btnSearch
            // 
            this.btnSearch.Location = new System.Drawing.Point(432, 7);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(90, 21);
            this.btnSearch.TabIndex = 1;
            this.btnSearch.Text = "検索";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // tbxSearchText
            // 
            this.tbxSearchText.Location = new System.Drawing.Point(12, 9);
            this.tbxSearchText.Name = "tbxSearchText";
            this.tbxSearchText.Size = new System.Drawing.Size(407, 19);
            this.tbxSearchText.TabIndex = 0;
            // 
            // SettingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(534, 321);
            this.Controls.Add(this.lviSuperiorsList);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.tbxCCList);
            this.Controls.Add(this.btnCC);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.tbxSearchText);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingForm";
            this.Load += new System.EventHandler(this.SettingForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView lviSuperiorsList;
        private System.Windows.Forms.ColumnHeader columnHeaderName;
        private System.Windows.Forms.ColumnHeader columnHeaderDeployment;
        private System.Windows.Forms.ColumnHeader columnHeaderTitle;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox tbxCCList;
        private System.Windows.Forms.Button btnCC;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.TextBox tbxSearchText;
    }
}