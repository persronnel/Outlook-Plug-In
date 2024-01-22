namespace UnityAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.create = this.Factory.CreateRibbonButton();
            this.resolve = this.Factory.CreateRibbonButton();
            this.close = this.Factory.CreateRibbonButton();
            this.search = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.edt = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Unity";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.create);
            this.group1.Items.Add(this.resolve);
            this.group1.Items.Add(this.close);
            this.group1.Items.Add(this.search);
            this.group1.Label = "Manage Cases";
            this.group1.Name = "group1";
            // 
            // create
            // 
            this.create.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.create.Label = "Create";
            this.create.Name = "create";
            this.create.ShowImage = true;
            this.create.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.create_Click);
            // 
            // resolve
            // 
            this.resolve.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.resolve.Label = "Resolve";
            this.resolve.Name = "resolve";
            this.resolve.ShowImage = true;
            this.resolve.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.resolve_Click);
            // 
            // close
            // 
            this.close.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.close.Label = "Close";
            this.close.Name = "close";
            this.close.ShowImage = true;
            this.close.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.close_Click);
            // 
            // search
            // 
            this.search.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.search.Label = "Search";
            this.search.Name = "search";
            this.search.ShowImage = true;
            this.search.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.search_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.edt);
            this.group2.Label = "High Level Function";
            this.group2.Name = "group2";
            // 
            // edt
            // 
            this.edt.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.edt.Label = "Edit";
            this.edt.Name = "edt";
            this.edt.ShowImage = true;
            this.edt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.edt_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton create;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton resolve;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton close;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton search;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton edt;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
