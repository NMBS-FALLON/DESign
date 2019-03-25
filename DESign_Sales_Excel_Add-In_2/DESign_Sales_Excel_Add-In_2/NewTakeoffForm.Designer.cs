namespace DESign_Sales_Excel_Add_In
{
    partial class NewTakeoffForm
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
            this.clbTakeoffType = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCreateNewTakeoff = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // clbTakeoffType
            // 
            this.clbTakeoffType.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clbTakeoffType.FormattingEnabled = true;
            this.clbTakeoffType.Items.AddRange(new object[] {
            "Steel-On-Steel",
            "Wood Nailer"});
            this.clbTakeoffType.Location = new System.Drawing.Point(72, 48);
            this.clbTakeoffType.Name = "clbTakeoffType";
            this.clbTakeoffType.Size = new System.Drawing.Size(140, 48);
            this.clbTakeoffType.TabIndex = 0;
            this.clbTakeoffType.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.clbTakeoffType_ItemChecked);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(68, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(144, 19);
            this.label1.TabIndex = 1;
            this.label1.Text = "Select Takeoff Type:";
            // 
            // btnCreateNewTakeoff
            // 
            this.btnCreateNewTakeoff.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateNewTakeoff.Location = new System.Drawing.Point(81, 121);
            this.btnCreateNewTakeoff.Name = "btnCreateNewTakeoff";
            this.btnCreateNewTakeoff.Size = new System.Drawing.Size(117, 36);
            this.btnCreateNewTakeoff.TabIndex = 2;
            this.btnCreateNewTakeoff.Text = "Create Takeoff";
            this.btnCreateNewTakeoff.UseVisualStyleBackColor = true;
            this.btnCreateNewTakeoff.Click += new System.EventHandler(this.btnCreateNewTakeoff_Click);
            // 
            // NewTakeoffForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(276, 169);
            this.Controls.Add(this.btnCreateNewTakeoff);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.clbTakeoffType);
            this.Name = "NewTakeoffForm";
            this.Text = "Takeoff Selection";
            this.Load += new System.EventHandler(this.NewTakeoffForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox clbTakeoffType;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCreateNewTakeoff;
    }
}