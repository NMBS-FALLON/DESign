using System;
using System.Collections.Generic;

using System.Windows.Forms;
using DESign_Sales_Excel_Add_in.Worksheet_Values;

namespace DESign_Sales_Excel_Add_in
{
    public partial class Convert_Takeoff_Form : Form
    {
        Takeoff takeoff = new Takeoff();
        public Convert_Takeoff_Form()
        {
            InitializeComponent();
            takeoff = takeoff.ImportTakeoff();

            for (int i = 0; i < takeoff.Sequences.Count; i++)
            {

                dataGridSeperateSeismic.Rows.Add();

                dataGridSeperateSeismic.Rows[i].Cells[0].Value = takeoff.Sequences[i].Name.Text;
            }

        }

        private void btnConvertTakeoff_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < takeoff.Sequences.Count; i++)
            {
                takeoff.Sequences[i].SeperateSeismic = Convert.ToBoolean(dataGridSeperateSeismic.Rows[i].Cells[1].Value);
                takeoff.Sequences[i].SDS = Convert.ToDouble(dataGridSeperateSeismic.Rows[i].Cells[2].Value);
            }
            this.Close();
            takeoff.SeperateSeismic();
            takeoff.CreateOriginalTakeoff();
            
        }

        private void dataGridSeperateSeismic_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnRules_Click(object sender, EventArgs e)
        {
            MessageBox.Show("RULES FOR SEPERATING SEISMIC:\r\n" +
                            "   - SEPERATION IS ONLY ALLOWED ON ROOFS.\r\n" +
                            "   - IF THE JOIST DESIGNATION LL IS FROM\r\n" +
                            "     SNOW, THE FLAT ROOF SNOW LOAD(Pf)\r\n" +
                            "     MUST BE LESS THAN 30 PSF.\r\n" +
                            "   - FOR SEPERATION TO OCCUR ON GIRDERS,\r\n" +
                            "     THE DESIGNATION MUST BE IN TL/LL FORM\r\n" +
                            "     (i.e. 54G7N12.5/5.8K).");
        }
    }
}
