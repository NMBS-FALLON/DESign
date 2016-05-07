namespace DESign_WordAddIn
{
    partial class FormNailBacksheet
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormNailBacksheet));
            this.btnCreateTable = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tboxTolerance = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tBoxDWGBY = new System.Windows.Forms.TextBox();
            this.checkBoxExcelData = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tBoxScrewSpacing = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnCreateTable
            // 
            this.btnCreateTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateTable.Location = new System.Drawing.Point(32, 272);
            this.btnCreateTable.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.btnCreateTable.Name = "btnCreateTable";
            this.btnCreateTable.Size = new System.Drawing.Size(105, 28);
            this.btnCreateTable.TabIndex = 0;
            this.btnCreateTable.Text = "CREATE";
            this.btnCreateTable.UseVisualStyleBackColor = true;
            this.btnCreateTable.Click += new System.EventHandler(this.btnCreateTable_Click);
            // 
            // label1
            // 
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label1.Location = new System.Drawing.Point(12, 186);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(106, 36);
            this.label1.TabIndex = 1;
            this.label1.Text = "Wood Length Tolerance (in.):";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tboxTolerance
            // 
            this.tboxTolerance.Location = new System.Drawing.Point(122, 193);
            this.tboxTolerance.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tboxTolerance.Name = "tboxTolerance";
            this.tboxTolerance.Size = new System.Drawing.Size(55, 22);
            this.tboxTolerance.TabIndex = 2;
            this.tboxTolerance.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(11, 64);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 22);
            this.label2.TabIndex = 4;
            this.label2.Text = "Nail Placement:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 232);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(68, 17);
            this.label3.TabIndex = 5;
            this.label3.Text = "DWG BY:";
            // 
            // tBoxDWGBY
            // 
            this.tBoxDWGBY.Location = new System.Drawing.Point(122, 229);
            this.tBoxDWGBY.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tBoxDWGBY.Name = "tBoxDWGBY";
            this.tBoxDWGBY.Size = new System.Drawing.Size(55, 22);
            this.tBoxDWGBY.TabIndex = 6;
            this.tBoxDWGBY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.tBoxDWGBY.TextChanged += new System.EventHandler(this.tBoxDWGBY_TextChanged);
            // 
            // checkBoxExcelData
            // 
            this.checkBoxExcelData.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.checkBoxExcelData.Checked = true;
            this.checkBoxExcelData.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxExcelData.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.8F);
            this.checkBoxExcelData.Location = new System.Drawing.Point(9, 26);
            this.checkBoxExcelData.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.checkBoxExcelData.Name = "checkBoxExcelData";
            this.checkBoxExcelData.Size = new System.Drawing.Size(159, 28);
            this.checkBoxExcelData.TabIndex = 7;
            this.checkBoxExcelData.Text = "Use Excel Data";
            this.checkBoxExcelData.UseVisualStyleBackColor = true;
            this.checkBoxExcelData.CheckedChanged += new System.EventHandler(this.checkBoxExcelData_CheckedChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 159);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(92, 17);
            this.label4.TabIndex = 8;
            this.label4.Text = "Spacing (in.):";
            // 
            // tBoxScrewSpacing
            // 
            this.tBoxScrewSpacing.Location = new System.Drawing.Point(122, 159);
            this.tBoxScrewSpacing.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.tBoxScrewSpacing.Name = "tBoxScrewSpacing";
            this.tBoxScrewSpacing.Size = new System.Drawing.Size(55, 22);
            this.tBoxScrewSpacing.TabIndex = 9;
            this.tBoxScrewSpacing.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // FormNailBacksheet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(591, 329);
            this.Controls.Add(this.tBoxScrewSpacing);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.checkBoxExcelData);
            this.Controls.Add(this.tBoxDWGBY);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tboxTolerance);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCreateTable);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.Name = "FormNailBacksheet";
            this.Text = "Create Nailer Backsheet";
            this.Load += new System.EventHandler(this.FormNailBacksheet_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCreateTable;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tboxTolerance;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tBoxDWGBY;
        private System.Windows.Forms.CheckBox checkBoxExcelData;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tBoxScrewSpacing;

    }
}