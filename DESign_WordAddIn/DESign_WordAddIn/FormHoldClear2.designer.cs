namespace DESign_WordAddIn
{
    partial class FormHoldClear2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormHoldClear2));
            this.btnCreateHoldClears = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCreateHoldClears
            // 
            this.btnCreateHoldClears.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnCreateHoldClears.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateHoldClears.Location = new System.Drawing.Point(67, 11);
            this.btnCreateHoldClears.Margin = new System.Windows.Forms.Padding(2);
            this.btnCreateHoldClears.Name = "btnCreateHoldClears";
            this.btnCreateHoldClears.Size = new System.Drawing.Size(81, 23);
            this.btnCreateHoldClears.TabIndex = 0;
            this.btnCreateHoldClears.Text = "CREATE";
            this.btnCreateHoldClears.UseVisualStyleBackColor = true;
            this.btnCreateHoldClears.Click += new System.EventHandler(this.btnCreateHoldClears_Click);
            // 
            // FormHoldClear2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(212, 248);
            this.Controls.Add(this.btnCreateHoldClears);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "FormHoldClear2";
            this.Text = "Hold Clear";
            this.Load += new System.EventHandler(this.FormHoldClear2_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCreateHoldClears;


    }
}