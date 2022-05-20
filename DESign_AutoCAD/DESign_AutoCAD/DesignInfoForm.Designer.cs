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
            this.cbPLant = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tbJobNumber = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cbDataSource = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // clbInfoSelect
            // 
            this.clbInfoSelect.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clbInfoSelect.FormattingEnabled = true;
            this.clbInfoSelect.Items.AddRange(new object[] {
            "JOIST TC WIDTH (STANDARD)",
            "JOIST TC WIDTH (1 1/8\" GAP)",
            "GIRDER TC WIDTH",
            "JOIST BOLT LENGTH",
            "WEIGHT",
            "TC MAX BRIDGING",
            "BC MAX BRIDGING"});
            this.clbInfoSelect.Location = new System.Drawing.Point(18, 44);
            this.clbInfoSelect.Margin = new System.Windows.Forms.Padding(2);
            this.clbInfoSelect.Name = "clbInfoSelect";
            this.clbInfoSelect.Size = new System.Drawing.Size(348, 202);
            this.clbInfoSelect.TabIndex = 0;
            this.clbInfoSelect.SelectedIndexChanged += new System.EventHandler(this.clbInfoSelect_SelectedIndexChanged);
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
            this.btnAddInfo.Location = new System.Drawing.Point(18, 398);
            this.btnAddInfo.Margin = new System.Windows.Forms.Padding(2);
            this.btnAddInfo.Name = "btnAddInfo";
            this.btnAddInfo.Size = new System.Drawing.Size(210, 36);
            this.btnAddInfo.TabIndex = 2;
            this.btnAddInfo.Text = "ADD INFO";
            this.btnAddInfo.UseVisualStyleBackColor = true;
            this.btnAddInfo.Click += new System.EventHandler(this.btnAddInfo_Click);
            // 
            // cbPLant
            // 
            this.cbPLant.FormattingEnabled = true;
            this.cbPLant.Items.AddRange(new object[] {
            "Fallon",
            "Juarez"});
            this.cbPLant.Location = new System.Drawing.Point(134, 313);
            this.cbPLant.Name = "cbPLant";
            this.cbPLant.Size = new System.Drawing.Size(121, 21);
            this.cbPLant.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(15, 317);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "Plant:";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(15, 347);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 17);
            this.label3.TabIndex = 5;
            this.label3.Text = "Job #:";
            // 
            // tbJobNumber
            // 
            this.tbJobNumber.Location = new System.Drawing.Point(134, 344);
            this.tbJobNumber.Name = "tbJobNumber";
            this.tbJobNumber.Size = new System.Drawing.Size(121, 20);
            this.tbJobNumber.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(15, 282);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(103, 17);
            this.label4.TabIndex = 8;
            this.label4.Text = "Data Source:";
            // 
            // cbDataSource
            // 
            this.cbDataSource.FormattingEnabled = true;
            this.cbDataSource.Items.AddRange(new object[] {
            "SQL",
            "Joist Details"});
            this.cbDataSource.Location = new System.Drawing.Point(134, 282);
            this.cbDataSource.Name = "cbDataSource";
            this.cbDataSource.Size = new System.Drawing.Size(121, 21);
            this.cbDataSource.TabIndex = 9;
            // 
            // DesignInfoForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(389, 460);
            this.Controls.Add(this.cbDataSource);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tbJobNumber);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cbPLant);
            this.Controls.Add(this.btnAddInfo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.clbInfoSelect);
            this.Margin = new System.Windows.Forms.Padding(2);
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
        private System.Windows.Forms.ComboBox cbPLant;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbJobNumber;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbDataSource;
    }
}