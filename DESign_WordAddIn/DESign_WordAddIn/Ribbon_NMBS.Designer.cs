namespace DESign_WordAddIn
{
    partial class Ribbon_NMBS : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon_NMBS()
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
            this.groupManualOps = this.Factory.CreateRibbonGroup();
            this.btnNailerBackSheet = this.Factory.CreateRibbonButton();
            this.btnHoldClear = this.Factory.CreateRibbonButton();
            this.groupPrint = this.Factory.CreateRibbonGroup();
            this.btnPrintShopCopies = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupManualOps.SuspendLayout();
            this.groupPrint.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupManualOps);
            this.tab1.Groups.Add(this.groupPrint);
            this.tab1.Label = "NMBS";
            this.tab1.Name = "tab1";
            // 
            // groupManualOps
            // 
            this.groupManualOps.Items.Add(this.btnNailerBackSheet);
            this.groupManualOps.Items.Add(this.btnHoldClear);
            this.groupManualOps.Label = "Manual Ops";
            this.groupManualOps.Name = "groupManualOps";
            // 
            // btnNailerBackSheet
            // 
            this.btnNailerBackSheet.Label = "Nailer Backsheet";
            this.btnNailerBackSheet.Name = "btnNailerBackSheet";
            this.btnNailerBackSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNailerBackSheet_Click);
            // 
            // btnHoldClear
            // 
            this.btnHoldClear.Label = "Hold Clear";
            this.btnHoldClear.Name = "btnHoldClear";
            this.btnHoldClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHoldClear_Click);
            // 
            // groupPrint
            // 
            this.groupPrint.Items.Add(this.btnPrintShopCopies);
            this.groupPrint.Label = "Print";
            this.groupPrint.Name = "groupPrint";
            // 
            // btnPrintShopCopies
            // 
            this.btnPrintShopCopies.Label = "Print Shop Copies";
            this.btnPrintShopCopies.Name = "btnPrintShopCopies";
            this.btnPrintShopCopies.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintShopCopies_Click);
            // 
            // Ribbon_NMBS
            // 
            this.Name = "Ribbon_NMBS";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_NMBS_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupManualOps.ResumeLayout(false);
            this.groupManualOps.PerformLayout();
            this.groupPrint.ResumeLayout(false);
            this.groupPrint.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupManualOps;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNailerBackSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHoldClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPrint;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintShopCopies;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_NMBS Ribbon_NMBS
        {
            get { return this.GetRibbon<Ribbon_NMBS>(); }
        }
    }
}
