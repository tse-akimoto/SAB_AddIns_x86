namespace OutlookAddInSAB
{
    partial class Ribbon_Test : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon_Test()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon_Test));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnSendCheck = this.Factory.CreateRibbonButton();
            this.btnGetAddress = this.Factory.CreateRibbonButton();
            this.btnTempList = this.Factory.CreateRibbonButton();
            this.btnGetUserJobTitle = this.Factory.CreateRibbonButton();
            this.btnZipCreate = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnSettingFormView = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            resources.ApplyResources(this.tab1, "tab1");
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            resources.ApplyResources(this.group1, "group1");
            this.group1.Name = "group1";
            // 
            // button1
            // 
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            resources.ApplyResources(this.button2, "button2");
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.button4);
            this.group2.Items.Add(this.button3);
            resources.ApplyResources(this.group2, "group2");
            this.group2.Name = "group2";
            // 
            // button5
            // 
            resources.ApplyResources(this.button5, "button5");
            this.button5.Name = "button5";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button4
            // 
            resources.ApplyResources(this.button4, "button4");
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button3
            // 
            resources.ApplyResources(this.button3, "button3");
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnSendCheck);
            this.group3.Items.Add(this.btnGetAddress);
            this.group3.Items.Add(this.btnTempList);
            this.group3.Items.Add(this.btnGetUserJobTitle);
            this.group3.Items.Add(this.btnZipCreate);
            resources.ApplyResources(this.group3, "group3");
            this.group3.Name = "group3";
            // 
            // btnSendCheck
            // 
            resources.ApplyResources(this.btnSendCheck, "btnSendCheck");
            this.btnSendCheck.Name = "btnSendCheck";
            this.btnSendCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSendCheck_Click);
            // 
            // btnGetAddress
            // 
            resources.ApplyResources(this.btnGetAddress, "btnGetAddress");
            this.btnGetAddress.Name = "btnGetAddress";
            this.btnGetAddress.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetAddress_Click);
            // 
            // btnTempList
            // 
            resources.ApplyResources(this.btnTempList, "btnTempList");
            this.btnTempList.Name = "btnTempList";
            this.btnTempList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTempList_Click);
            // 
            // btnGetUserJobTitle
            // 
            resources.ApplyResources(this.btnGetUserJobTitle, "btnGetUserJobTitle");
            this.btnGetUserJobTitle.Name = "btnGetUserJobTitle";
            this.btnGetUserJobTitle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetUserJobTitle_Click);
            // 
            // btnZipCreate
            // 
            resources.ApplyResources(this.btnZipCreate, "btnZipCreate");
            this.btnZipCreate.Name = "btnZipCreate";
            this.btnZipCreate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnZipCreate_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnSettingFormView);
            resources.ApplyResources(this.group4, "group4");
            this.group4.Name = "group4";
            // 
            // btnSettingFormView
            // 
            this.btnSettingFormView.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.btnSettingFormView, "btnSettingFormView");
            this.btnSettingFormView.Name = "btnSettingFormView";
            this.btnSettingFormView.ShowImage = true;
            this.btnSettingFormView.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettingFormView_Click);
            // 
            // Ribbon_Test
            // 
            this.Name = "Ribbon_Test";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Close += new System.EventHandler(this.Ribbon_Test_Close);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Test_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnZipCreate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTempList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetAddress;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSendCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetUserJobTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettingFormView;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_Test Ribbon_Test
        {
            get { return this.GetRibbon<Ribbon_Test>(); }
        }
    }
}
