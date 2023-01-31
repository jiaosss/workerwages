
namespace workerwages
{
    partial class Workwages : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Workwages()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.findfile = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.splitexcel = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.mergeexcel = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "澄澄的工具箱";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.findfile);
            this.group1.Label = "工人账号信息核对";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "选择目录";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "OpenAppointment";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Label = "选择信息表";
            this.button2.Name = "button2";
            this.button2.OfficeImageId = "FileNewDocument";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button2_Click);
            // 
            // findfile
            // 
            this.findfile.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.findfile.Label = "生成汇总表";
            this.findfile.Name = "findfile";
            this.findfile.OfficeImageId = "GroupNavigate";
            this.findfile.ShowImage = true;
            this.findfile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Findfile_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.splitexcel);
            this.group3.Items.Add(this.separator1);
            this.group3.Items.Add(this.mergeexcel);
            this.group3.Label = "表格操作";
            this.group3.Name = "group3";
            // 
            // splitexcel
            // 
            this.splitexcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitexcel.Label = "拆分表格";
            this.splitexcel.Name = "splitexcel";
            this.splitexcel.OfficeImageId = "ContactCardCopy";
            this.splitexcel.ShowImage = true;
            this.splitexcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Splitexcel_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // mergeexcel
            // 
            this.mergeexcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.mergeexcel.Label = "合并表格";
            this.mergeexcel.Name = "mergeexcel";
            this.mergeexcel.OfficeImageId = "GroupAdpQueryType";
            this.mergeexcel.ShowImage = true;
            this.mergeexcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Mergeexcel_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.label1);
            this.group2.Items.Add(this.label2);
            this.group2.Name = "group2";
            // 
            // label1
            // 
            this.label1.Label = "v2.1正式版";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = "优化代码逻辑，调高效率";
            this.label2.Name = "label2";
            // 
            // workwages
            // 
            this.Name = "workwages";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Workwages_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton findfile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton splitexcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton mergeexcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal Workwages workwages
        {
            get { return this.GetRibbon<Workwages>(); }
        }
    }
}
