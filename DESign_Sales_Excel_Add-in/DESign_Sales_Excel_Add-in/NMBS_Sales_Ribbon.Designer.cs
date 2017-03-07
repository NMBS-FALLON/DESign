namespace DESign_Sales_Excel_Add_in
{
    partial class NMBS_Sales_Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public NMBS_Sales_Ribbon()
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
            this.nmbsTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.nmbsTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // nmbsTab
            // 
            this.nmbsTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.nmbsTab.Groups.Add(this.group1);
            this.nmbsTab.Label = "NMBS";
            this.nmbsTab.Name = "nmbsTab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "TAKEOFF";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Label = "Convert TO";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // NMBS_Sales_Ribbon
            // 
            this.Name = "NMBS_Sales_Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.nmbsTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.NMBS_Sales_Ribbon_Load);
            this.nmbsTab.ResumeLayout(false);
            this.nmbsTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab nmbsTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal NMBS_Sales_Ribbon NMBS_Sales_Ribbon
        {
            get { return this.GetRibbon<NMBS_Sales_Ribbon>(); }
        }
    }
}
