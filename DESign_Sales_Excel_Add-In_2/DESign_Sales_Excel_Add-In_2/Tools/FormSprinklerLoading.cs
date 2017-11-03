using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DESign_Sales_Excel_Add_in.Tools
{
    public partial class FormSprinklerLoading : Form
    {
        public FormSprinklerLoading()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void btnAddSprinklerLoad_Click(object sender, EventArgs e)
        {
            string braceAngle = cbFSBraceAngle.Text;
            double latBraceLoad = Convert.ToDouble(tbLatBraceLoad.Text);
            double longBraceLoad = Convert.ToDouble(tbLongBraceLoad.Text);
            double minSpace = Convert.ToDouble(tbMinBraceSpace.Text);
            double maxJoistLength = Convert.ToDouble(tbMaxJoistLength.Text);
            int pipeWeight = Convert.ToInt16(tbPipeWeight.Text);

            double adjustedLatBraceLoad = latBraceLoad;
            if (braceAngle == "60-90") { adjustedLatBraceLoad = latBraceLoad * 0.58; }

            double cpLoad = Math.Max(longBraceLoad, adjustedLatBraceLoad);
            double uLoad = Math.Ceiling((((maxJoistLength / minSpace) - 1) * adjustedLatBraceLoad) / maxJoistLength);

            MessageBox.Show(String.Format(@"
CP = {0}
U = {1}
AX = {2}
CL = {3}", cpLoad, uLoad, longBraceLoad, pipeWeight));

        }
    }
}
