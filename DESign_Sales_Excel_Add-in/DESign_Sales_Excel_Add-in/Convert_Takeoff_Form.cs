using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            this.Close();
            Takeoff takeoff = new Takeoff();
            takeoff = takeoff.ImportTakeoff();
            takeoff.CreateOriginalTakeoff(takeoff);          
        }
    }
}
