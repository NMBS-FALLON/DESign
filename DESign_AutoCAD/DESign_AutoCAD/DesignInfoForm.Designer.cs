namespace DESign_AutoCAD
{
    partial class DesignInfoForm
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
            this.clbInfoSelect = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnAddInfo = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // clbInfoSelect
            // 
            this.clbInfoSelect.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clbInfoSelect.FormattingEnabled = true;
            this.clbInfoSelect.Items.AddRange(new object[] {
            "JOIST TC WIDTH",
            "JOIST BOLT LENGTH",
            "GIRDER TC WIDTH"});
            this.clbInfoSelect.Location = new System.Drawing.Point(18, 44);
            this.clbInfoSelect.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.clbInfoSelect.Name = "clbInfoSelect";
            this.clbInfoSelect.Size = new System.Drawing.Size(211, 114);
            this.clbInfoSelect.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(14, 7);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(165, 19);
            this.label1.TabIndex = 1;
            this.label1.Text = "SELECT INFO TO ADD:";
            // 
            // btnAddInfo
            // 
            this.btnAddInfo.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAddInfo.Location = new System.Drawing.Point(17, 178);
            this.btnAddInfo.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnAddInfo.Name = "btnAddInfo";
            this.btnAddInfo.Size = new System.Drawing.Size(210, 36);
            this.btnAddInfo.TabIndex = 2;
            this.btnAddInfo.Text = "ADD INFO";
            this.btnAddInfo.UseVisualStyleBackColor = true;
            this.btnAddInfo.Click += new System.EventHandler(this.btnAddInfo_Click);
            // 
            // DesignInfoForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(244, 225);
            this.Controls.Add(this.btnAddInfo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.clbInfoSelect);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "DesignInfoForm";
            this.Text = "INFO SELECTION";
            this.Load += new System.EventHandler(this.DesignInfoForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox clbInfoSelect;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnAddInfo;
    }
}