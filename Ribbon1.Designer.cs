namespace JobApplications_Excel_Add_in
{
    partial class JobsTracker : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public JobsTracker()
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
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnJobdetails = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnNewSheet = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "JobsTracker";
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnJobdetails);
            this.group2.Name = "group2";
            // 
            // btnJobdetails
            // 
            this.btnJobdetails.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnJobdetails.Image = global::JobApplications_Excel_Add_in.Properties.Resources.newjob;
            this.btnJobdetails.Label = "Job details";
            this.btnJobdetails.Name = "btnJobdetails";
            this.btnJobdetails.ShowImage = true;
            this.btnJobdetails.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnJobdetails_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnNewSheet);
            this.group3.Name = "group3";
            // 
            // btnNewSheet
            // 
            this.btnNewSheet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNewSheet.Label = "New Sheet";
            this.btnNewSheet.Name = "btnNewSheet";
            this.btnNewSheet.ShowImage = true;
            this.btnNewSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewSheet_Click);
            // 
            // JobsTracker
            // 
            this.Name = "JobsTracker";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnJobdetails;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewSheet;
    }

    partial class ThisRibbonCollection
    {
        internal JobsTracker Ribbon1
        {
            get { return this.GetRibbon<JobsTracker>(); }
        }
    }
}
