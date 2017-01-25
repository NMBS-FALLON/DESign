using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace DESign_Sales_Excel_Add_in
{
    public partial class NMBS_Sales_Ribbon
    {
        private void NMBS_Sales_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application thisApp = Globals.ThisAddIn.Application;
            Excel.Workbook thisWB = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet thisWS = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            
        }
    }
}
