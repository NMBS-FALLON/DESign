using System;

using System.Windows.Forms;
using DESign_Sales_Excel_Add_in.Worksheet_Values;

namespace DESign_Sales_Excel_Add_in
{
    public partial class Convert_Takeoff_Form : Form
    {
        public Convert_Takeoff_Form()
        {
            InitializeComponent();
        }

        private void btnConvertTakeoff_Click(object sender, EventArgs e)
        {
            
            Takeoff takeoff = new Takeoff();
            takeoff = takeoff.ImportTakeoff();
            if(cbSeperateSeismic.Checked == true)
            {
                takeoff.SeperateSeismic(Math.Ceiling(Convert.ToDouble(tbSDS.Text)));
            }
            
            takeoff.CreateOriginalTakeoff();
            this.Close();
        }
    }
}
