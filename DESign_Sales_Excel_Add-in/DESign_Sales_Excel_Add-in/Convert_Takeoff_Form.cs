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
                takeoff.SeperateSeismic(Convert.ToDouble(tbSDS.Text));
            }
            
            takeoff.CreateOriginalTakeoff();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dfObjectArray = Deedle.Frame.ReadCsv("C:\\Users\\darien.shannon\\Desktop\\DECK\\Deck Tables.csv");
            string s = "s";
        }
    }
}
