﻿using System.Diagnostics;
using System.Deployment.Application;
using System.Reflection;

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
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                System.Version v = ApplicationDeployment.CurrentDeployment.CurrentVersion;
                tab1.Label = "NMBS (v" + v.Revision.ToString() + ")";
            }
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
            this.btnBlankWorksheets = this.Factory.CreateRibbonButton();
            this.groupPrint = this.Factory.CreateRibbonGroup();
            this.btnPrintJShopCopies = this.Factory.CreateRibbonButton();
            this.btnPrintGShopCopies = this.Factory.CreateRibbonButton();
            this.btnPrint1Master = this.Factory.CreateRibbonButton();
            this.btnPrint1Cut = this.Factory.CreateRibbonButton();
            this.groupJuarezPrint = this.Factory.CreateRibbonGroup();
            this.btnPrintJuarez = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupManualOperations.SuspendLayout();
            this.groupPrint.SuspendLayout();
            this.groupJuarezPrint.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupManualOperations);
            this.tab1.Groups.Add(this.groupPrint);
            this.tab1.Groups.Add(this.groupJuarezPrint);
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
            // btnBlankWorksheets
            // 
            this.btnBlankWorksheets.Label = "Blank Sheets";
            this.btnBlankWorksheets.Name = "btnBlankWorksheets";
            this.btnBlankWorksheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBlankWorksheets_Click);
            // 
            // groupPrint
            // 
            this.groupPrint.Items.Add(this.btnPrintJShopCopies);
            this.groupPrint.Items.Add(this.btnPrintGShopCopies);
            this.groupPrint.Items.Add(this.btnPrint1Master);
            this.groupPrint.Items.Add(this.btnPrint1Cut);
            this.groupPrint.Label = "FALLON PRINTING";
            this.groupPrint.Name = "groupPrint";
            // 
            // btnPrintJShopCopies
            // 
            this.btnPrintJShopCopies.Label = "Print J Shop Copies";
            this.btnPrintJShopCopies.Name = "btnPrintJShopCopies";
            this.btnPrintJShopCopies.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintJShopCopies_Click);
            // 
            // btnPrintGShopCopies
            // 
            this.btnPrintGShopCopies.Label = "Print G Shop Copies";
            this.btnPrintGShopCopies.Name = "btnPrintGShopCopies";
            this.btnPrintGShopCopies.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintGShopCopies_Click);
            // 
            // btnPrint1Master
            // 
            this.btnPrint1Master.Label = "Print 1 Master";
            this.btnPrint1Master.Name = "btnPrint1Master";
            this.btnPrint1Master.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrint1Master_Click);
            // 
            // btnPrint1Cut
            // 
            this.btnPrint1Cut.Label = "Print 1 Cut";
            this.btnPrint1Cut.Name = "btnPrint1Cut";
            this.btnPrint1Cut.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrint1Cut_Click);
            // 
            // groupJuarezPrint
            // 
            this.groupJuarezPrint.Items.Add(this.btnPrintJuarez);
            this.groupJuarezPrint.Label = "JUAREZ PRINTING";
            this.groupJuarezPrint.Name = "groupJuarezPrint";
            // 
            // btnPrintJuarez
            // 
            this.btnPrintJuarez.Label = "Print Shop Copies";
            this.btnPrintJuarez.Name = "btnPrintJuarez";
            this.btnPrintJuarez.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrintJuarezShopCopies_Click);
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
            this.groupJuarezPrint.ResumeLayout(false);
            this.groupJuarezPrint.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupManualOperations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNailBacksheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHoldClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPrint;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintJShopCopies;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBlankWorksheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupJuarezPrint;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintJuarez;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrintGShopCopies;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrint1Master;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrint1Cut;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonNMBS Ribbon1
        {
            get { return this.GetRibbon<RibbonNMBS>(); }
        }
    }
}
