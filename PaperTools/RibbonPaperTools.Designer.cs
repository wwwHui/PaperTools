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
            this.groupColor = this.Factory.CreateRibbonGroup();
            this.zoteroCitationColor = this.Factory.CreateRibbonButton();
            this.wordCitationColor = this.Factory.CreateRibbonButton();
            this.groupDoc = this.Factory.CreateRibbonGroup();
            this.buttonReomve = this.Factory.CreateRibbonButton();
            this.buttonCNReplace = this.Factory.CreateRibbonButton();
            this.buttonENReplace = this.Factory.CreateRibbonButton();
            this.PaperTools.SuspendLayout();
            this.groupPandoc.SuspendLayout();
            this.groupColor.SuspendLayout();
            this.groupDoc.SuspendLayout();
            this.SuspendLayout();
            // 
            // PaperTools
            // 
            this.PaperTools.Groups.Add(this.groupPandoc);
            this.PaperTools.Groups.Add(this.groupColor);
            this.PaperTools.Groups.Add(this.groupDoc);
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
            // groupColor
            // 
            this.groupColor.Items.Add(this.zoteroCitationColor);
            this.groupColor.Items.Add(this.wordCitationColor);
            this.groupColor.Label = "颜色";
            this.groupColor.Name = "groupColor";
            // 
            // zoteroCitationColor
            // 
            this.zoteroCitationColor.Label = "Zotero引用";
            this.zoteroCitationColor.Name = "zoteroCitationColor";
            this.zoteroCitationColor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.zoteroCitationColor_Click);
            // 
            // wordCitationColor
            // 
            this.wordCitationColor.Label = "交叉引用";
            this.wordCitationColor.Name = "wordCitationColor";
            this.wordCitationColor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.wordCitationColor_Click);
            // 
            // groupDoc
            // 
            this.groupDoc.Items.Add(this.buttonReomve);
            this.groupDoc.Items.Add(this.buttonCNReplace);
            this.groupDoc.Items.Add(this.buttonENReplace);
            this.groupDoc.Label = "Doc";
            this.groupDoc.Name = "groupDoc";
            // 
            // buttonReomve
            // 
            this.buttonReomve.Label = "去掉空格和换行";
            this.buttonReomve.Name = "buttonReomve";
            this.buttonReomve.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonReomve_Click);
            // 
            // buttonCNReplace
            // 
            this.buttonCNReplace.Label = "英文符号转中文";
            this.buttonCNReplace.Name = "buttonCNReplace";
            this.buttonCNReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCNReplace_Click);
            // 
            // buttonENReplace
            // 
            this.buttonENReplace.Label = "中文符号转英文";
            this.buttonENReplace.Name = "buttonENReplace";
            this.buttonENReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonENReplace_Click);
            // 
            // RibbonPaperTools
            // 
            this.Name = "RibbonPaperTools";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.PaperTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonPaperTools_Load);
            this.PaperTools.ResumeLayout(false);
            this.PaperTools.PerformLayout();
            this.groupPandoc.ResumeLayout(false);
            this.groupPandoc.PerformLayout();
            this.groupColor.ResumeLayout(false);
            this.groupColor.PerformLayout();
            this.groupDoc.ResumeLayout(false);
            this.groupDoc.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab PaperTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPandoc;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton zoteroCitationColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPandocVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonExportLatex;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDoc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonReomve;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCNReplace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonENReplace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton wordCitationColor;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonPaperTools Ribbon1
        {
            get { return this.GetRibbon<RibbonPaperTools>(); }
        }
    }
}
