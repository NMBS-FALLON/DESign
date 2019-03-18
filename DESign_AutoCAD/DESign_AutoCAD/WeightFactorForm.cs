using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DESign_AutoCAD
{
    public partial class WeightFactorForm : Form
    {
        public WeightFactorForm()
        {
            InitializeComponent();
            tbAddWeight.Text = "2.0 %";
        }

        public double WeightPercentToAdd { get; set; }

        private void tbAddWeight_Leave(object sender, EventArgs e)
        {
            var text = tbAddWeight.Text;
            var parsedValue = 0.0;
            var parsed = double.TryParse(text.Replace(" %", ""), out parsedValue);
            if (parsed)
            {
                tbAddWeight.Text = parsedValue + " %";
                this.btnAddWeightPercent.Click += new System.EventHandler(this.btnAddWeight_Click);
            }
            else
            {
                MessageBox.Show("Value must be a decimal value.");
                tbAddWeight.Focus();
                this.btnAddWeightPercent.Click += new System.EventHandler(delegate (Object o, EventArgs a) { });
            }
        }

        private void btnAddWeight_Click(object sender, EventArgs e)
        {
            this.WeightPercentToAdd = double.Parse(tbAddWeight.Text.Replace(" %", "")) / 100.0;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
