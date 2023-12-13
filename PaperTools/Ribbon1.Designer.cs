namespace PaperTools
{
    partial class RibbonPaperTools : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonPaperTools()
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.PaperTools = this.Factory.CreateRibbonTab();
            this.groupPandoc = this.Factory.CreateRibbonGroup();
            this.buttonPandocVersion = this.Factory.CreateRibbonButton();
            this.buttonExportLatex = this.Factory.CreateRibbonButton();
            this.groupZotero = this.Factory.CreateRibbonGroup();
            this.citationColor = this.Factory.CreateRibbonButton();
            this.PaperTools.SuspendLayout();
            this.groupPandoc.SuspendLayout();
            this.groupZotero.SuspendLayout();
            this.SuspendLayout();
            // 
            // PaperTools
            // 
            this.PaperTools.Groups.Add(this.groupPandoc);
            this.PaperTools.Groups.Add(this.groupZotero);
            this.PaperTools.Label = "PaperTools";
            this.PaperTools.Name = "PaperTools";
            // 
            // groupPandoc
            // 
            this.groupPandoc.Items.Add(this.buttonPandocVersion);
            this.groupPandoc.Items.Add(this.buttonExportLatex);
            this.groupPandoc.Label = "pandoc";
            this.groupPandoc.Name = "groupPandoc";
            // 
            // buttonPandocVersion
            // 
            this.buttonPandocVersion.Label = "Pandoc";
            this.buttonPandocVersion.Name = "buttonPandocVersion";
            // 
            // buttonExportLatex
            // 
            this.buttonExportLatex.Label = "导出latex";
            this.buttonExportLatex.Name = "buttonExportLatex";
            this.buttonExportLatex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonExportLatex_Click);
            // 
            // groupZotero
            // 
            this.groupZotero.Items.Add(this.citationColor);
            this.groupZotero.Label = "Zotero";
            this.groupZotero.Name = "groupZotero";
            // 
            // citationColor
            // 
            this.citationColor.Label = "引用颜色";
            this.citationColor.Name = "citationColor";
            this.citationColor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.citationColor_Click);
            // 
            // RibbonPaperTools
            // 
            this.Name = "RibbonPaperTools";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.PaperTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.PaperTools.ResumeLayout(false);
            this.PaperTools.PerformLayout();
            this.groupPandoc.ResumeLayout(false);
            this.groupPandoc.PerformLayout();
            this.groupZotero.ResumeLayout(false);
            this.groupZotero.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab PaperTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPandoc;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupZotero;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton citationColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPandocVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonExportLatex;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonPaperTools Ribbon1
        {
            get { return this.GetRibbon<RibbonPaperTools>(); }
        }
    }
}
