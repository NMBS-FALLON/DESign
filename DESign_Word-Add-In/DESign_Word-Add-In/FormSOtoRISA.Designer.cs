namespace DESign_WordAddIn
{
    partial class FormSOtoRISA
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
            this.btnExportSOtoRISA = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnExportSOtoRISA
            // 
            this.btnExportSOtoRISA.Location = new System.Drawing.Point(102, 12);
            this.btnExportSOtoRISA.Name = "btnExportSOtoRISA";
            this.btnExportSOtoRISA.Size = new System.Drawing.Size(75, 23);
            this.btnExportSOtoRISA.TabIndex = 0;
            this.btnExportSOtoRISA.Text = "EXPORT";
            this.btnExportSOtoRISA.UseVisualStyleBackColor = true;
            this.btnExportSOtoRISA.Click += new System.EventHandler(this.btnExportSOtoRISA_Click);
            // 
            // FormSOtoRISA
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.btnExportSOtoRISA);
            this.Name = "FormSOtoRISA";
            this.Text = "FormSOtoRISA";
            this.Load += new System.EventHandler(this.FormSOtoRISA_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnExportSOtoRISA;
    }
}