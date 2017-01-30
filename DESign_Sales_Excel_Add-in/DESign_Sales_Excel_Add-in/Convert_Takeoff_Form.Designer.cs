namespace DESign_Sales_Excel_Add_in
{
    partial class Convert_Takeoff_Form
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
            this.btnConvertTakeoff = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnConvertTakeoff
            // 
            this.btnConvertTakeoff.Location = new System.Drawing.Point(24, 22);
            this.btnConvertTakeoff.Name = "btnConvertTakeoff";
            this.btnConvertTakeoff.Size = new System.Drawing.Size(89, 34);
            this.btnConvertTakeoff.TabIndex = 0;
            this.btnConvertTakeoff.Text = "Convert Takeoff";
            this.btnConvertTakeoff.UseVisualStyleBackColor = true;
            this.btnConvertTakeoff.Click += new System.EventHandler(this.btnConvertTakeoff_Click);
            // 
            // Convert_Takeoff_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.btnConvertTakeoff);
            this.Name = "Convert_Takeoff_Form";
            this.Text = "Convert_Takeoff_Form";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnConvertTakeoff;
    }
}