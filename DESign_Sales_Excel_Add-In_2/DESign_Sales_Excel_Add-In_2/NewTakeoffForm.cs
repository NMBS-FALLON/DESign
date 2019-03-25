using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DESign_Sales_Excel_Add_In
{
    public enum TakeoffType { SteelOnSteel, WoodNailer };
    public partial class NewTakeoffForm : Form
    {
        public NewTakeoffForm()
        {
            InitializeComponent();
        }

        private void NewTakeoffForm_Load(object sender, EventArgs e)
        {

        }

        

        public TakeoffType TakeoffType {get; set;}

        private void btnCreateNewTakeoff_Click(object sender, EventArgs e)
        {
            if (clbTakeoffType.CheckedIndices.Contains(0))
            {
                TakeoffType = TakeoffType.SteelOnSteel;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else if (clbTakeoffType.CheckedIndices.Contains(1))
            {
                TakeoffType = TakeoffType.WoodNailer;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("No Takeoff Type Selected");
            }
        }

        private void clbTakeoffType_ItemChecked(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue == CheckState.Checked)
            {
                for (int ix = 0; ix < clbTakeoffType.Items.Count; ++ix)
                {
                    if (e.Index != ix)
                    {
                        clbTakeoffType.SetItemChecked(ix, false);
                    }
                }
            }

        }
    }
}
