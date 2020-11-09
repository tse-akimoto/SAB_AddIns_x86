namespace WordAddInSAB
{
    partial class RibbonDocumentManagement : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonDocumentManagement()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonDocumentManagement));
            this.tabGCPDocumentManagement = this.Factory.CreateRibbonTab();
            this.grpDocumentManagement = this.Factory.CreateRibbonGroup();
            this.btnSAB = this.Factory.CreateRibbonButton();
            this.tabGCPDocumentManagement.SuspendLayout();
            this.grpDocumentManagement.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabGCPDocumentManagement
            // 
            this.tabGCPDocumentManagement.Groups.Add(this.grpDocumentManagement);
            resources.ApplyResources(this.tabGCPDocumentManagement, "tabGCPDocumentManagement");
            this.tabGCPDocumentManagement.Name = "tabGCPDocumentManagement";
            // 
            // grpDocumentManagement
            // 
            this.grpDocumentManagement.Items.Add(this.btnSAB);
            resources.ApplyResources(this.grpDocumentManagement, "grpDocumentManagement");
            this.grpDocumentManagement.Name = "grpDocumentManagement";
            // 
            // btnSAB
            // 
            this.btnSAB.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.btnSAB, "btnSAB");
            this.btnSAB.Name = "btnSAB";
            this.btnSAB.ShowImage = true;
            this.btnSAB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSAB_Click);
            // 
            // RibbonDocumentManagement
            // 
            this.Name = "RibbonDocumentManagement";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabGCPDocumentManagement);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonDocumentManagement_Load);
            this.tabGCPDocumentManagement.ResumeLayout(false);
            this.tabGCPDocumentManagement.PerformLayout();
            this.grpDocumentManagement.ResumeLayout(false);
            this.grpDocumentManagement.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabGCPDocumentManagement;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDocumentManagement;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSAB;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonDocumentManagement RibbonDocumentManagement
        {
            get { return this.GetRibbon<RibbonDocumentManagement>(); }
        }
    }
}
