namespace DESign_WordAddIn.Insert_Blank_Sheets
{
    partial class FormInsertBlankSheets
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
            this.btnInsertBirdcage = new System.Windows.Forms.Button();
            this.btnInsertBlankTPlate = new System.Windows.Forms.Button();
            this.btnInsertBlankShimPlate = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnInsertBirdcage
            // 
            this.btnInsertBirdcage.Location = new System.Drawing.Point(53, 24);
            this.btnInsertBirdcage.Name = "btnInsertBirdcage";
            this.btnInsertBirdcage.Size = new System.Drawing.Size(154, 23);
            this.btnInsertBirdcage.TabIndex = 0;
            this.btnInsertBirdcage.Text = "Insert Blank Birdcage";
            this.btnInsertBirdcage.UseVisualStyleBackColor = true;
            this.btnInsertBirdcage.Click += new System.EventHandler(this.btnInsertBirdcage_Click);
            // 
            // btnInsertBlankTPlate
            // 
            this.btnInsertBlankTPlate.Location = new System.Drawing.Point(53, 68);
            this.btnInsertBlankTPlate.Name = "btnInsertBlankTPlate";
            this.btnInsertBlankTPlate.Size = new System.Drawing.Size(154, 23);
            this.btnInsertBlankTPlate.TabIndex = 1;
            this.btnInsertBlankTPlate.Text = "Insert Blank T-Plate";
            this.btnInsertBlankTPlate.UseVisualStyleBackColor = true;
            this.btnInsertBlankTPlate.Click += new System.EventHandler(this.btnInsertBlankTPlate_Click);
            // 
            // btnInsertBlankShimPlate
            // 
            this.btnInsertBlankShimPlate.Location = new System.Drawing.Point(53, 110);
            this.btnInsertBlankShimPlate.Name = "btnInsertBlankShimPlate";
            this.btnInsertBlankShimPlate.Size = new System.Drawing.Size(154, 23);
            this.btnInsertBlankShimPlate.TabIndex = 2;
            this.btnInsertBlankShimPlate.Text = "Insert Blank Shim Detail";
            this.btnInsertBlankShimPlate.UseVisualStyleBackColor = true;
            this.btnInsertBlankShimPlate.Click += new System.EventHandler(this.btnInsertBlankShimPlate_Click);
            // 
            // FormInsertBlankSheets
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(264, 154);
            this.Controls.Add(this.btnInsertBlankShimPlate);
            this.Controls.Add(this.btnInsertBlankTPlate);
            this.Controls.Add(this.btnInsertBirdcage);
            this.Name = "FormInsertBlankSheets";
            this.Text = "FormInsertBlankSheets";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnInsertBirdcage;
        private System.Windows.Forms.Button btnInsertBlankTPlate;
        private System.Windows.Forms.Button btnInsertBlankShimPlate;
    }
}