using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using DESign_Sales_Excel_Add_in.Worksheet_Values;
using System.Runtime.InteropServices;

namespace DESign_Sales_Excel_Add_in
{
    public partial class NMBS_Sales_Ribbon
    {
        private void NMBS_Sales_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
        
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Convert_Takeoff_Form convert_Takeoff_Form = new Convert_Takeoff_Form();
            convert_Takeoff_Form.Show();
        }

        private void btnNewTakeoff_Click(object sender, RibbonControlEventArgs e)
        {

            string excelPath = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllBytes(excelPath, Properties.Resources.TAKEOFF_CONCEPT);

            Excel.Application oXL = Globals.ThisAddIn.Application;
            Excel.Workbooks workbooks = oXL.Workbooks;
            Excel.Workbook workbook;
            workbook = workbooks.Open(excelPath);

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsm)|*.xlsm";
            saveFileDialog.ShowDialog();
            if (saveFileDialog.FileName != "")
            {
                workbook.CheckCompatibility = false;
                workbook.SaveAs(saveFileDialog.FileName);
            }

            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(oXL);
            GC.Collect();

        }
    }
}
