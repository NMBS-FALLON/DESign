namespace DESign_WordAddIn
{
    partial class RibbonNMBS : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonNMBS()
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
            this.groupManualOperations = this.Factory.CreateRibbonGroup();
            this.btnNailBacksheet = this.Factory.CreateRibbonButton();
            this.btnHoldClear = this.Factory.CreateRibbonButton();
            this.groupPrint = this.Factory.CreateRibbonGroup();
            this.btnPrintShopCopies = this.Factory.CreateRibbonButton();
            this.btnSinglePrintShopCopy = this.Factory.CreateRibbonButton();
            this.btnBlankWorksheets = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupManualOperations.SuspendLayout();
            this.groupPrint.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupManualOperations);
            this.tab1.Groups.Add(this.groupPrint);
            this.tab1.Label = "NMBS";
            this.tab1.Name = "tab1";
            // 
            // groupManualOperations
            // 
            this.groupManualOperations.Items.Add(this.btnNailBacksheet);
            this.groupManualOperations.Items.Add(this.btnHoldClear);
            this.groupManualOperations.Items.Add(this.btnBlankWorksheets);
            this.groupManualOperations.Label = "MAN. OPERATIONS";
            this.groupManualOperations.Name = "groupManualOperations";
            // 
            // btnNailBacksheet
            // 
            this.btnNailBacksheet.Label = "Nailer Backsheet";
            this.btnNailBacksheet.Name = "btnNailBacksheet";
            this.btnNailBacksheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNailBacksheet_Click);
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
            this.groupPrint.Items.Add(this.btnSinglePrintShopCopy);
            this.groupPrint.Label = "PRINTING";
            this.groupPrint.Name = "groupPrint";
            // 
            // btnPrintShopCopies
            // 
            this.btnPrintShopCopies.Label = "Print Shop Copies";
            this.btnPrintShopCopies.Name = "btnPrintShopCopies";
            this.btnPrintShopCopies.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintShopCopies_Click);
            // 
            // btnSinglePrintShopCopy
            // 
            this.btnSinglePrintShopCopy.Label = "Single Shop Copy";
            this.btnSinglePrintShopCopy.Name = "btnSinglePrintShopCopy";
            this.btnSinglePrintShopCopy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSinglePrintShopCopy_Click);
            // 
            // btnBlankWorksheets
            // 
            this.btnBlankWorksheets.Label = "Blank Sheets";
            this.btnBlankWorksheets.Name = "btnBlankWorksheets";
            this.btnBlankWorksheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBlankWorksheets_Click);
            // 
            // RibbonNMBS
            // 
            this.Name = "RibbonNMBS";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupManualOperations.ResumeLayout(false);
            this.groupManualOperations.PerformLayout();
            this.groupPrint.ResumeLayout(false);
            this.groupPrint.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupManualOperations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNailBacksheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHoldClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPrint;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintShopCopies;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSinglePrintShopCopy;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBlankWorksheets;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonNMBS Ribbon1
        {
            get { return this.GetRibbon<RibbonNMBS>(); }
        }
    }
}
