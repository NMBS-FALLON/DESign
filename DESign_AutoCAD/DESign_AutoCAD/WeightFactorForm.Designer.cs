namespace DESign_AutoCAD
{
    partial class WeightFactorForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.btnAddWeightPercent = new System.Windows.Forms.Button();
            this.tbAddWeight = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(8, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(205, 19);
            this.label1.TabIndex = 0;
            this.label1.Text = "Additional Weight % To Add:";
            // 
            // btnAddWeightPercent
            // 
            this.btnAddWeightPercent.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAddWeightPercent.Location = new System.Drawing.Point(12, 91);
            this.btnAddWeightPercent.Name = "btnAddWeightPercent";
            this.btnAddWeightPercent.Size = new System.Drawing.Size(49, 26);
            this.btnAddWeightPercent.TabIndex = 1;
            this.btnAddWeightPercent.Text = "OK";
            this.btnAddWeightPercent.UseVisualStyleBackColor = true;
            this.btnAddWeightPercent.Click += new System.EventHandler(this.btnAddWeight_Click);
            // 
            // tbAddWeight
            // 
            this.tbAddWeight.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbAddWeight.Location = new System.Drawing.Point(12, 45);
            this.tbAddWeight.Name = "tbAddWeight";
            this.tbAddWeight.Size = new System.Drawing.Size(100, 27);
            this.tbAddWeight.TabIndex = 2;
            this.tbAddWeight.Leave += new System.EventHandler(this.tbAddWeight_Leave);
            // 
            // WeightFactorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(219, 129);
            this.Controls.Add(this.tbAddWeight);
            this.Controls.Add(this.btnAddWeightPercent);
            this.Controls.Add(this.label1);
            this.Name = "WeightFactorForm";
            this.Text = "Weight Factor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnAddWeightPercent;
        private System.Windows.Forms.TextBox tbAddWeight;
    }
}