namespace DESign_BOT
{
    partial class ModifyBomForm
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
            this.clbBomModifications = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnModifyBom = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // clbBomModifications
            // 
            this.clbBomModifications.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clbBomModifications.FormattingEnabled = true;
            this.clbBomModifications.Items.AddRange(new object[] {
            "Seperate Seismic",
            "Check Girder IP"});
            this.clbBomModifications.Location = new System.Drawing.Point(21, 63);
            this.clbBomModifications.Name = "clbBomModifications";
            this.clbBomModifications.Size = new System.Drawing.Size(276, 92);
            this.clbBomModifications.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(17, 19);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(188, 19);
            this.label1.TabIndex = 2;
            this.label1.Text = "SELECT MODIFICATIONS:";
            // 
            // btnModifyBom
            // 
            this.btnModifyBom.Font = new System.Drawing.Font("Corbel", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnModifyBom.Location = new System.Drawing.Point(21, 181);
            this.btnModifyBom.Name = "btnModifyBom";
            this.btnModifyBom.Size = new System.Drawing.Size(276, 39);
            this.btnModifyBom.TabIndex = 3;
            this.btnModifyBom.Text = "MODIFY BOM";
            this.btnModifyBom.UseVisualStyleBackColor = true;
            this.btnModifyBom.Click += new System.EventHandler(this.BtnModifyBom_Click);
            // 
            // ModifyBomForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(321, 232);
            this.Controls.Add(this.btnModifyBom);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.clbBomModifications);
            this.Name = "ModifyBomForm";
            this.Text = "SELECT MODIFICATIONS";
            this.Load += new System.EventHandler(this.ModifyBomForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox clbBomModifications;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnModifyBom;
    }
}