namespace VisualPaste
{
    partial class CangjieRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CangjieRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.TabHome = this.Factory.CreateRibbonTab();
            this.grpCangjieImport = this.Factory.CreateRibbonGroup();
            this.btnCangjieImport = this.Factory.CreateRibbonMenu();
            this.btnTableTitle = this.Factory.CreateRibbonButton();
            this.btnCopyImp = this.Factory.CreateRibbonButton();
            this.TabHome.SuspendLayout();
            this.grpCangjieImport.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabHome
            // 
            this.TabHome.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabHome.ControlId.OfficeId = "TabHome";
            this.TabHome.Groups.Add(this.grpCangjieImport);
            this.TabHome.Label = "TabHome";
            this.TabHome.Name = "TabHome";
            // 
            // grpCangjieImport
            // 
            this.grpCangjieImport.Items.Add(this.btnCangjieImport);
            this.grpCangjieImport.Name = "grpCangjieImport";
            // 
            // btnCangjieImport
            // 
            this.btnCangjieImport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCangjieImport.Image = global::VSTOWordProject.Properties.Resources.logo;
            this.btnCangjieImport.Items.Add(this.btnTableTitle);
            this.btnCangjieImport.Items.Add(this.btnCopyImp);
            this.btnCangjieImport.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCangjieImport.Label = "ImpForWeb";
            this.btnCangjieImport.Name = "btnCangjieImport";
            this.btnCangjieImport.OfficeImageId = "PageMenu";
            this.btnCangjieImport.ShowImage = true;
            // 
            // btnTableTitle
            // 
            this.btnTableTitle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTableTitle.Image = global::VSTOWordProject.Properties.Resources.sheet;
            this.btnTableTitle.Label = "TableHeader";
            this.btnTableTitle.Name = "btnTableTitle";
            this.btnTableTitle.ShowImage = true;
            // 
            // btnCopyImp
            // 
            this.btnCopyImp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCopyImp.Image = global::VSTOWordProject.Properties.Resources.copy;
            this.btnCopyImp.Label = "ExportHTML";
            this.btnCopyImp.Name = "btnCopyImp";
            this.btnCopyImp.ShowImage = true;
            // 
            // CangjieRibbon
            // 
            this.Name = "CangjieRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.TabHome);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CangjieImportRibbon_Load);
            this.TabHome.ResumeLayout(false);
            this.TabHome.PerformLayout();
            this.grpCangjieImport.ResumeLayout(false);
            this.grpCangjieImport.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabHome;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCangjieImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu btnCangjieImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTableTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyImp;
    }

    partial class ThisRibbonCollection
    {
        internal CangjieRibbon CangjieRibbon
        {
            get { return this.GetRibbon<CangjieRibbon>(); }
        }
    }
}
