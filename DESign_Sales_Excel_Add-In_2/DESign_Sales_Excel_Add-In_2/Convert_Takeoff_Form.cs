using System;
using System.Windows.Forms;
using DESign_Sales_Excel_Add_In_2.Worksheet_Values;

namespace DESign_Sales_Excel_Add_In_2
{
  public partial class Convert_Takeoff_Form : Form
  {
    private readonly Takeoff takeoff = new Takeoff();

    public Convert_Takeoff_Form()
    {
      InitializeComponent();
      takeoff = takeoff.ImportTakeoff();

      for (var i = 0; i < takeoff.Sequences.Count; i++)
      {
        dataGridSeperateSeismic.Rows.Add();

        dataGridSeperateSeismic.Rows[i].Cells[0].Value = takeoff.Sequences[i].Name.Text;
        dataGridSeperateSeismic.Rows[i].Cells[2].Value = takeoff.SDS;
      }
    }

    private void btnConvertTakeoff_Click(object sender, EventArgs e)
    {
      for (var i = 0; i < takeoff.Sequences.Count; i++)
      {
        takeoff.Sequences[i].SeperateSeismic = Convert.ToBoolean(dataGridSeperateSeismic.Rows[i].Cells[1].Value);
        takeoff.Sequences[i].SDS = Convert.ToDouble(dataGridSeperateSeismic.Rows[i].Cells[2].Value);
      }

      Close();
      takeoff.SeperateSeismic();
      takeoff.CreateOriginalTakeoff();
    }

    private void dataGridSeperateSeismic_CellContentClick(object sender, DataGridViewCellEventArgs e)
    {
    }
  }
}