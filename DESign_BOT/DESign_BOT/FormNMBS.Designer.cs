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
            this.button3 = new System.Windows.Forms.Button();
            this.tBoxWoodReq = new System.Windows.Forms.TextBox();
            this.button4 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.btnQuickTCWidth = new System.Windows.Forms.Button();
            this.tabControlConvertBOM.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControlConvertBOM
            // 
            this.tabControlConvertBOM.Controls.Add(this.tabPage1);
            this.tabControlConvertBOM.Controls.Add(this.tabPage2);
            this.tabControlConvertBOM.Location = new System.Drawing.Point(0, 0);
            this.tabControlConvertBOM.Margin = new System.Windows.Forms.Padding(2);
            this.tabControlConvertBOM.Name = "tabControlConvertBOM";
            this.tabControlConvertBOM.SelectedIndex = 0;
            this.tabControlConvertBOM.Size = new System.Drawing.Size(301, 341);
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
            this.tabPage1.Size = new System.Drawing.Size(293, 256);
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
            this.tabPage2.Controls.Add(this.btnQuickTCWidth);
            this.tabPage2.Controls.Add(this.button3);
            this.tabPage2.Controls.Add(this.tBoxWoodReq);
            this.tabPage2.Controls.Add(this.button4);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(293, 315);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Wd Nailr";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(8, 176);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(52, 42);
            this.button3.TabIndex = 5;
            this.button3.Text = "TC Widths";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // tBoxWoodReq
            // 
            this.tBoxWoodReq.Location = new System.Drawing.Point(78, 120);
            this.tBoxWoodReq.Multiline = true;
            this.tBoxWoodReq.Name = "tBoxWoodReq";
            this.tBoxWoodReq.Size = new System.Drawing.Size(200, 98);
            this.tBoxWoodReq.TabIndex = 4;
            this.tBoxWoodReq.TextChanged += new System.EventHandler(this.tBoxWoodReq_TextChanged);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(8, 120);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(52, 42);
            this.button4.TabIndex = 3;
            this.button4.Text = "Wood Req.";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button5);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Location = new System.Drawing.Point(3, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(289, 112);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "AB EXCEL SHEETS";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(19, 80);
            this.button5.Margin = new System.Windows.Forms.Padding(2);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(131, 19);
            this.button5.TabIndex = 1;
            this.button5.Text = "MANUAL AB SHEET";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(19, 51);
            this.button2.Margin = new System.Windows.Forms.Padding(2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(131, 19);
            this.button2.TabIndex = 1;
            this.button2.Text = "AB FROM NUCOR BOM";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(19, 19);
            this.button1.Margin = new System.Windows.Forms.Padding(2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(131, 19);
            this.button1.TabIndex = 0;
            this.button1.Text = "A&B FROM NMBS BOM";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnQuickTCWidth
            // 
            this.btnQuickTCWidth.Location = new System.Drawing.Point(8, 251);
            this.btnQuickTCWidth.Name = "btnQuickTCWidth";
            this.btnQuickTCWidth.Size = new System.Drawing.Size(75, 23);
            this.btnQuickTCWidth.TabIndex = 6;
            this.btnQuickTCWidth.Text = "TC Widths ";
            this.btnQuickTCWidth.UseVisualStyleBackColor = true;
            this.btnQuickTCWidth.Click += new System.EventHandler(this.btnQuickTCWidth_Click);
            // 
            // FormNMBSHelper
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(298, 331);
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
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button btnQuickTCWidth;
    }
}

