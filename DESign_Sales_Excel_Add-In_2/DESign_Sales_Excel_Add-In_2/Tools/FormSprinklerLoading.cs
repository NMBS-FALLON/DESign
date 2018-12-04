using System;
using System.Windows.Forms;

namespace DESign_Sales_Excel_Add_In_2.Tools
{
  public class FormSprinklerLoading : Form
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
      var latBraceLoad = Convert.ToDouble(tbLatBraceLoad.Text);
      var longBraceLoad = Convert.ToDouble(tbLongBraceLoad.Text);
      var minSpace = Convert.ToDouble(tbMinBraceSpace.Text);
      var maxJoistLength = Convert.ToDouble(tbMaxJoistLength.Text);
      int pipeWeight = Convert.ToInt16(tbPipeWeight.Text);

      var adjustedLatBraceLoad = latBraceLoad;
      if (braceAngle == "60-90") adjustedLatBraceLoad = latBraceLoad * 0.58;

      var cpLoad = Math.Max(longBraceLoad, adjustedLatBraceLoad);
      var uLoad = Math.Ceiling((maxJoistLength / minSpace - 1) * adjustedLatBraceLoad / maxJoistLength);

      MessageBox.Show(string.Format(@"
CP = {0}
U = {1}
AX = {2}
CL = {3}", cpLoad, uLoad, longBraceLoad, pipeWeight));
    }
  }
}