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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Convert_Takeoff_Form));
            this.btnConvertTakeoff = new System.Windows.Forms.Button();
            this.dataGridSeperateSeismic = new System.Windows.Forms.DataGridView();
            this.sequence = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.seperateSeismic = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.sds = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridSeperateSeismic)).BeginInit();
            this.SuspendLayout();
            // 
            // btnConvertTakeoff
            // 
            this.btnConvertTakeoff.Location = new System.Drawing.Point(12, 12);
            this.btnConvertTakeoff.Name = "btnConvertTakeoff";
            this.btnConvertTakeoff.Size = new System.Drawing.Size(95, 26);
            this.btnConvertTakeoff.TabIndex = 0;
            this.btnConvertTakeoff.Text = "Convert Takeoff";
            this.btnConvertTakeoff.UseVisualStyleBackColor = true;
            this.btnConvertTakeoff.Click += new System.EventHandler(this.btnConvertTakeoff_Click);
            // 
            // dataGridSeperateSeismic
            // 
            this.dataGridSeperateSeismic.AllowUserToAddRows = false;
            this.dataGridSeperateSeismic.AllowUserToDeleteRows = false;
            this.dataGridSeperateSeismic.AllowUserToResizeRows = false;
            this.dataGridSeperateSeismic.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridSeperateSeismic.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridSeperateSeismic.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.sequence,
            this.seperateSeismic,
            this.sds});
            this.dataGridSeperateSeismic.Location = new System.Drawing.Point(12, 44);
            this.dataGridSeperateSeismic.Name = "dataGridSeperateSeismic";
            this.dataGridSeperateSeismic.RowHeadersVisible = false;
            this.dataGridSeperateSeismic.Size = new System.Drawing.Size(279, 115);
            this.dataGridSeperateSeismic.TabIndex = 1;
            this.dataGridSeperateSeismic.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridSeperateSeismic_CellContentClick);
            // 
            // sequence
            // 
            this.sequence.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.sequence.HeaderText = "Sequence";
            this.sequence.MinimumWidth = 100;
            this.sequence.Name = "sequence";
            this.sequence.ReadOnly = true;
            this.sequence.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // seperateSeismic
            // 
            this.seperateSeismic.HeaderText = "Seperate Seismic";
            this.seperateSeismic.Name = "seperateSeismic";
            this.seperateSeismic.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.seperateSeismic.Width = 60;
            // 
            // sds
            // 
            this.sds.HeaderText = "SDS";
            this.sds.Name = "sds";
            this.sds.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.sds.Width = 54;
            // 
            // richTextBox1
            // 
            this.richTextBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.richTextBox1.Location = new System.Drawing.Point(12, 165);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(278, 110);
            this.richTextBox1.TabIndex = 2;
            this.richTextBox1.Text = resources.GetString("richTextBox1.Text");
            // 
            // Convert_Takeoff_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(303, 287);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.dataGridSeperateSeismic);
            this.Controls.Add(this.btnConvertTakeoff);
            this.Name = "Convert_Takeoff_Form";
            this.Text = "Convert_Takeoff_Form";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridSeperateSeismic)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnConvertTakeoff;
        private System.Windows.Forms.DataGridView dataGridSeperateSeismic;
        private System.Windows.Forms.DataGridViewTextBoxColumn sequence;
        private System.Windows.Forms.DataGridViewCheckBoxColumn seperateSeismic;
        private System.Windows.Forms.DataGridViewTextBoxColumn sds;
        private System.Windows.Forms.RichTextBox richTextBox1;
    }
}