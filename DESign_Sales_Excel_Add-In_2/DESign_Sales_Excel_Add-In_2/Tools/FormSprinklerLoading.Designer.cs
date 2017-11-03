namespace DESign_Sales_Excel_Add_in.Tools
{
    partial class FormSprinklerLoading
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
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.tbLatBraceLoad = new System.Windows.Forms.TextBox();
            this.tbLongBraceLoad = new System.Windows.Forms.TextBox();
            this.tbMinBraceSpace = new System.Windows.Forms.TextBox();
            this.tbMaxJoistLength = new System.Windows.Forms.TextBox();
            this.cbFSBraceAngle = new System.Windows.Forms.ComboBox();
            this.btnAddSprinklerLoad = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.tbPipeWeight = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 27);
            this.label1.Margin = new System.Windows.Forms.Padding(3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "FS Brace Angle:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 49);
            this.label2.Margin = new System.Windows.Forms.Padding(3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Lateral Brace Load:";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(0, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 23);
            this.label3.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(22, 93);
            this.label4.Margin = new System.Windows.Forms.Padding(3);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(103, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Min. Brace Spacing:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(22, 115);
            this.label5.Margin = new System.Windows.Forms.Padding(3);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(90, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Max Joist Length:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(22, 71);
            this.label6.Margin = new System.Windows.Forms.Padding(3);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(125, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Longitudinal Brace Load:";
            // 
            // tbLatBraceLoad
            // 
            this.tbLatBraceLoad.Location = new System.Drawing.Point(150, 46);
            this.tbLatBraceLoad.Name = "tbLatBraceLoad";
            this.tbLatBraceLoad.Size = new System.Drawing.Size(100, 20);
            this.tbLatBraceLoad.TabIndex = 2;
            // 
            // tbLongBraceLoad
            // 
            this.tbLongBraceLoad.Location = new System.Drawing.Point(150, 68);
            this.tbLongBraceLoad.Name = "tbLongBraceLoad";
            this.tbLongBraceLoad.Size = new System.Drawing.Size(100, 20);
            this.tbLongBraceLoad.TabIndex = 3;
            // 
            // tbMinBraceSpace
            // 
            this.tbMinBraceSpace.Location = new System.Drawing.Point(150, 90);
            this.tbMinBraceSpace.Name = "tbMinBraceSpace";
            this.tbMinBraceSpace.Size = new System.Drawing.Size(100, 20);
            this.tbMinBraceSpace.TabIndex = 4;
            // 
            // tbMaxJoistLength
            // 
            this.tbMaxJoistLength.Location = new System.Drawing.Point(150, 112);
            this.tbMaxJoistLength.Name = "tbMaxJoistLength";
            this.tbMaxJoistLength.Size = new System.Drawing.Size(100, 20);
            this.tbMaxJoistLength.TabIndex = 5;
            // 
            // cbFSBraceAngle
            // 
            this.cbFSBraceAngle.FormattingEnabled = true;
            this.cbFSBraceAngle.Items.AddRange(new object[] {
            "<60",
            "60-90"});
            this.cbFSBraceAngle.Location = new System.Drawing.Point(150, 24);
            this.cbFSBraceAngle.Name = "cbFSBraceAngle";
            this.cbFSBraceAngle.Size = new System.Drawing.Size(100, 21);
            this.cbFSBraceAngle.TabIndex = 1;
            // 
            // btnAddSprinklerLoad
            // 
            this.btnAddSprinklerLoad.Location = new System.Drawing.Point(150, 173);
            this.btnAddSprinklerLoad.Name = "btnAddSprinklerLoad";
            this.btnAddSprinklerLoad.Size = new System.Drawing.Size(100, 23);
            this.btnAddSprinklerLoad.TabIndex = 7;
            this.btnAddSprinklerLoad.Text = "Add Load";
            this.btnAddSprinklerLoad.UseVisualStyleBackColor = true;
            this.btnAddSprinklerLoad.Click += new System.EventHandler(this.btnAddSprinklerLoad_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(22, 137);
            this.label7.Margin = new System.Windows.Forms.Padding(3);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 13);
            this.label7.TabIndex = 0;
            this.label7.Text = "Pipe Weight:";
            // 
            // tbPipeWeight
            // 
            this.tbPipeWeight.Location = new System.Drawing.Point(150, 134);
            this.tbPipeWeight.Name = "tbPipeWeight";
            this.tbPipeWeight.Size = new System.Drawing.Size(100, 20);
            this.tbPipeWeight.TabIndex = 6;
            // 
            // FormSprinklerLoading
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 208);
            this.Controls.Add(this.btnAddSprinklerLoad);
            this.Controls.Add(this.cbFSBraceAngle);
            this.Controls.Add(this.tbPipeWeight);
            this.Controls.Add(this.tbMaxJoistLength);
            this.Controls.Add(this.tbMinBraceSpace);
            this.Controls.Add(this.tbLongBraceLoad);
            this.Controls.Add(this.tbLatBraceLoad);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "FormSprinklerLoading";
            this.Text = "FormSprinklerLoading";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbLatBraceLoad;
        private System.Windows.Forms.TextBox tbLongBraceLoad;
        private System.Windows.Forms.TextBox tbMinBraceSpace;
        private System.Windows.Forms.TextBox tbMaxJoistLength;
        private System.Windows.Forms.ComboBox cbFSBraceAngle;
        private System.Windows.Forms.Button btnAddSprinklerLoad;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox tbPipeWeight;
    }
}