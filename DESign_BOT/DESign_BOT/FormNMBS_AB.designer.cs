namespace DESign_BOT
{
    partial class FormNMBS_AB
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormNMBS_AB));
            this.btnBOMtoExcel = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.holdBackDataViewNoteColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.holdBackDataViewAColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.holdBackDataViewBColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSpacing = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnBOMtoExcel
            // 
            this.btnBOMtoExcel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnBOMtoExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBOMtoExcel.Location = new System.Drawing.Point(128, 233);
            this.btnBOMtoExcel.Margin = new System.Windows.Forms.Padding(2);
            this.btnBOMtoExcel.Name = "btnBOMtoExcel";
            this.btnBOMtoExcel.Size = new System.Drawing.Size(80, 43);
            this.btnBOMtoExcel.TabIndex = 2;
            this.btnBOMtoExcel.TabStop = false;
            this.btnBOMtoExcel.Text = "CREATE";
            this.btnBOMtoExcel.UseVisualStyleBackColor = true;
            this.btnBOMtoExcel.Click += new System.EventHandler(this.btnBOMtoExcel_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.holdBackDataViewNoteColumn,
            this.holdBackDataViewAColumn,
            this.holdBackDataViewBColumn,
            this.colSpacing});
            this.dataGridView1.Location = new System.Drawing.Point(12, 10);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(305, 209);
            this.dataGridView1.TabIndex = 9;
            this.dataGridView1.TabStop = false;
            // 
            // holdBackDataViewNoteColumn
            // 
            this.holdBackDataViewNoteColumn.HeaderText = "NOTES:";
            this.holdBackDataViewNoteColumn.Name = "holdBackDataViewNoteColumn";
            // 
            // holdBackDataViewAColumn
            // 
            this.holdBackDataViewAColumn.HeaderText = "A:";
            this.holdBackDataViewAColumn.Name = "holdBackDataViewAColumn";
            // 
            // holdBackDataViewBColumn
            // 
            this.holdBackDataViewBColumn.HeaderText = "B:";
            this.holdBackDataViewBColumn.Name = "holdBackDataViewBColumn";
            // 
            // colSpacing
            // 
            this.colSpacing.HeaderText = "SPACING:";
            this.colSpacing.Name = "colSpacing";
            // 
            // FormNMBS_AB
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(327, 286);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnBOMtoExcel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "FormNMBS_AB";
            this.Text = "CREATE HOLDBACK FILE";
            this.Load += new System.EventHandler(this.formBOMtoExcel_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnBOMtoExcel;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn holdBackDataViewNoteColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn holdBackDataViewAColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn holdBackDataViewBColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSpacing;
    }
}