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
            this.cbSeperateSeismic = new System.Windows.Forms.CheckBox();
            this.labelSDS = new System.Windows.Forms.Label();
            this.tbSDS = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnConvertTakeoff
            // 
            this.btnConvertTakeoff.Location = new System.Drawing.Point(27, 216);
            this.btnConvertTakeoff.Name = "btnConvertTakeoff";
            this.btnConvertTakeoff.Size = new System.Drawing.Size(89, 34);
            this.btnConvertTakeoff.TabIndex = 0;
            this.btnConvertTakeoff.Text = "Convert Takeoff";
            this.btnConvertTakeoff.UseVisualStyleBackColor = true;
            this.btnConvertTakeoff.Click += new System.EventHandler(this.btnConvertTakeoff_Click);
            // 
            // cbSeperateSeismic
            // 
            this.cbSeperateSeismic.AutoSize = true;
            this.cbSeperateSeismic.Location = new System.Drawing.Point(27, 24);
            this.cbSeperateSeismic.Name = "cbSeperateSeismic";
            this.cbSeperateSeismic.Size = new System.Drawing.Size(108, 17);
            this.cbSeperateSeismic.TabIndex = 1;
            this.cbSeperateSeismic.Text = "Seperate Seismic";
            this.cbSeperateSeismic.UseVisualStyleBackColor = true;
            // 
            // labelSDS
            // 
            this.labelSDS.AutoSize = true;
            this.labelSDS.Location = new System.Drawing.Point(47, 44);
            this.labelSDS.Name = "labelSDS";
            this.labelSDS.Size = new System.Drawing.Size(41, 13);
            this.labelSDS.TabIndex = 2;
            this.labelSDS.Text = "SDS = ";
            // 
            // tbSDS
            // 
            this.tbSDS.Location = new System.Drawing.Point(94, 41);
            this.tbSDS.Name = "tbSDS";
            this.tbSDS.Size = new System.Drawing.Size(41, 20);
            this.tbSDS.TabIndex = 3;
            // 
            // Convert_Takeoff_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.tbSDS);
            this.Controls.Add(this.labelSDS);
            this.Controls.Add(this.cbSeperateSeismic);
            this.Controls.Add(this.btnConvertTakeoff);
            this.Name = "Convert_Takeoff_Form";
            this.Text = "Convert_Takeoff_Form";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnConvertTakeoff;
        private System.Windows.Forms.CheckBox cbSeperateSeismic;
        private System.Windows.Forms.Label labelSDS;
        private System.Windows.Forms.TextBox tbSDS;
    }
}