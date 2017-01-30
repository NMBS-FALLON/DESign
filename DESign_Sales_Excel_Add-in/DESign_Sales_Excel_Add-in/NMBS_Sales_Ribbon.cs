using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using DESign_Sales_Excel_Add_in.Worksheet_Values;

namespace DESign_Sales_Excel_Add_in
{
    public partial class NMBS_Sales_Ribbon
    {
        private void NMBS_Sales_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
        
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Takeoff takeoff = new Takeoff();
            takeoff = takeoff.ImportTakeoff();
            takeoff.CreateOriginalTakeoff(takeoff);
        }




    }
}
