namespace DESign_BOT
{
    partial class FormNMBSHelper
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormNMBSHelper));
            this.tabControlConvertBOM = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.labelProgramState = new System.Windows.Forms.Label();
            this.btnCreateNewBOM = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnWoodReqFromJoistDetails = new System.Windows.Forms.Button();
            this.btnQuickTCWidth = new System.Windows.Forms.Button();
            this.btnSeqSummaryFromShopOrders = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.tBoxWoodReq = new System.Windows.Forms.TextBox();
            this.btnWoodReqFromSOs = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.btnGetBomNotes = new System.Windows.Forms.Button();
            this.tabControlConvertBOM.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControlConvertBOM
            // 
            this.tabControlConvertBOM.Controls.Add(this.tabPage1);
            this.tabControlConvertBOM.Controls.Add(this.tabPage2);
            this.tabControlConvertBOM.Controls.Add(this.tabPage3);
            this.tabControlConvertBOM.Location = new System.Drawing.Point(0, 0);
            this.tabControlConvertBOM.Margin = new System.Windows.Forms.Padding(2);
            this.tabControlConvertBOM.Name = "tabControlConvertBOM";
            this.tabControlConvertBOM.SelectedIndex = 0;
            this.tabControlConvertBOM.Size = new System.Drawing.Size(331, 337);
            this.tabControlConvertBOM.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.labelProgramState);
            this.tabPage1.Controls.Add(this.btnCreateNewBOM);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage1.Size = new System.Drawing.Size(323, 311);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Convert BOM\'s";
            this.tabPage1.UseVisualStyleBackColor = true;
            this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // labelProgramState
            // 
            this.labelProgramState.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.labelProgramState.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelProgramState.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.labelProgramState.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProgramState.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.labelProgramState.Location = new System.Drawing.Point(16, 80);
            this.labelProgramState.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelProgramState.Name = "labelProgramState";
            this.labelProgramState.Size = new System.Drawing.Size(204, 110);
            this.labelProgramState.TabIndex = 1;
            this.labelProgramState.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnCreateNewBOM
            // 
            this.btnCreateNewBOM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateNewBOM.Location = new System.Drawing.Point(50, 31);
            this.btnCreateNewBOM.Margin = new System.Windows.Forms.Padding(2);
            this.btnCreateNewBOM.Name = "btnCreateNewBOM";
            this.btnCreateNewBOM.Size = new System.Drawing.Size(142, 38);
            this.btnCreateNewBOM.TabIndex = 0;
            this.btnCreateNewBOM.Text = "NUCOR TO NMBS";
            this.btnCreateNewBOM.UseVisualStyleBackColor = true;
            this.btnCreateNewBOM.Click += new System.EventHandler(this.btnCreateNewBOM_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btnWoodReqFromJoistDetails);
            this.tabPage2.Controls.Add(this.btnQuickTCWidth);
            this.tabPage2.Controls.Add(this.btnSeqSummaryFromShopOrders);
            this.tabPage2.Controls.Add(this.button3);
            this.tabPage2.Controls.Add(this.tBoxWoodReq);
            this.tabPage2.Controls.Add(this.btnWoodReqFromSOs);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(323, 311);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Wood Nailer";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnWoodReqFromJoistDetails
            // 
            this.btnWoodReqFromJoistDetails.Location = new System.Drawing.Point(7, 229);
            this.btnWoodReqFromJoistDetails.Name = "btnWoodReqFromJoistDetails";
            this.btnWoodReqFromJoistDetails.Size = new System.Drawing.Size(155, 30);
            this.btnWoodReqFromJoistDetails.TabIndex = 7;
            this.btnWoodReqFromJoistDetails.Text = "Wood Req. From Joist Details";
            this.btnWoodReqFromJoistDetails.UseVisualStyleBackColor = true;
            this.btnWoodReqFromJoistDetails.Click += new System.EventHandler(this.btnWoodReqFromJoistDetails_Click);
            // 
            // btnQuickTCWidth
            // 
            this.btnQuickTCWidth.Location = new System.Drawing.Point(8, 266);
            this.btnQuickTCWidth.Name = "btnQuickTCWidth";
            this.btnQuickTCWidth.Size = new System.Drawing.Size(155, 30);
            this.btnQuickTCWidth.TabIndex = 6;
            this.btnQuickTCWidth.Text = "TC Widths  From Joist Details";
            this.btnQuickTCWidth.UseVisualStyleBackColor = true;
            this.btnQuickTCWidth.Click += new System.EventHandler(this.btnQuickTCWidth_Click);
            // 
            // btnSeqSummaryFromShopOrders
            // 
            this.btnSeqSummaryFromShopOrders.Location = new System.Drawing.Point(8, 193);
            this.btnSeqSummaryFromShopOrders.Name = "btnSeqSummaryFromShopOrders";
            this.btnSeqSummaryFromShopOrders.Size = new System.Drawing.Size(155, 30);
            this.btnSeqSummaryFromShopOrders.TabIndex = 5;
            this.btnSeqSummaryFromShopOrders.Text = "Bolt Req. From S.O.\'s";
            this.btnSeqSummaryFromShopOrders.UseVisualStyleBackColor = true;
            this.btnSeqSummaryFromShopOrders.Click += new System.EventHandler(this.btnSeqSummaryFromShopOrders_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(8, 156);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(155, 30);
            this.button3.TabIndex = 5;
            this.button3.Text = "TC Widths From S.O.\'s";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // tBoxWoodReq
            // 
            this.tBoxWoodReq.Location = new System.Drawing.Point(169, 120);
            this.tBoxWoodReq.Multiline = true;
            this.tBoxWoodReq.Name = "tBoxWoodReq";
            this.tBoxWoodReq.Size = new System.Drawing.Size(146, 176);
            this.tBoxWoodReq.TabIndex = 4;
            // 
            // btnWoodReqFromSOs
            // 
            this.btnWoodReqFromSOs.Location = new System.Drawing.Point(8, 120);
            this.btnWoodReqFromSOs.Name = "btnWoodReqFromSOs";
            this.btnWoodReqFromSOs.Size = new System.Drawing.Size(155, 30);
            this.btnWoodReqFromSOs.TabIndex = 3;
            this.btnWoodReqFromSOs.Text = "Wood Req. From S.O.\'s";
            this.btnWoodReqFromSOs.UseVisualStyleBackColor = true;
            this.btnWoodReqFromSOs.Click += new System.EventHandler(this.button4_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button5);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Location = new System.Drawing.Point(3, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(312, 112);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "\'NOTE INFO\' EXCEL SHEETS";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(19, 80);
            this.button5.Margin = new System.Windows.Forms.Padding(2);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(172, 19);
            this.button5.TabIndex = 1;
            this.button5.Text = "MANUAL \'NOTE INFO\' SHEET";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(19, 51);
            this.button2.Margin = new System.Windows.Forms.Padding(2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(172, 19);
            this.button2.TabIndex = 1;
            this.button2.Text = "\'NOTE INFO\' FROM NUCOR BOM";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(19, 19);
            this.button1.Margin = new System.Windows.Forms.Padding(2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(172, 19);
            this.button1.TabIndex = 0;
            this.button1.Text = "\'NOTE INFO\' FROM NMBS BOM";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.btnGetBomNotes);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(323, 311);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "BOM Tools";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // btnGetBomNotes
            // 
            this.btnGetBomNotes.Location = new System.Drawing.Point(21, 13);
            this.btnGetBomNotes.Name = "btnGetBomNotes";
            this.btnGetBomNotes.Size = new System.Drawing.Size(116, 31);
            this.btnGetBomNotes.TabIndex = 0;
            this.btnGetBomNotes.Text = "Get BOM Notes";
            this.btnGetBomNotes.UseVisualStyleBackColor = true;
            this.btnGetBomNotes.Click += new System.EventHandler(this.BtnGetBomNotes_Click);
            // 
            // FormNMBSHelper
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(331, 342);
            this.Controls.Add(this.tabControlConvertBOM);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "FormNMBSHelper";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "DESign BOT";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControlConvertBOM.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControlConvertBOM;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button btnCreateNewBOM;
        private System.Windows.Forms.Label labelProgramState;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox tBoxWoodReq;
        private System.Windows.Forms.Button btnWoodReqFromSOs;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button btnQuickTCWidth;
        private System.Windows.Forms.Button btnWoodReqFromJoistDetails;
        private System.Windows.Forms.Button btnSeqSummaryFromShopOrders;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button btnGetBomNotes;
    }
}

