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

            Excel.Application oXLTemp = Globals.ThisAddIn.Application;
            Excel.Workbooks workbooks = oXLTemp.Workbooks;
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
            Marshal.ReleaseComObject(oXLTemp);
            GC.Collect();

        }

        private void btnInfo_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("SPECIAL BASETYPES:\r\n" +
                "   - [ALL]   : BRINGS INFO ONTO ALL MARKS\r\n" +
                "   - [ALL J] : BRINGS INFO ONTO ALL JOISTS\r\n" +
                "   - [ALL G] : BRINGS INFO ONTO ALL GIRDERS\r\n" +
                "RULES FOR SEPERATING SEISMIC:\r\n" +
                "   - ADD LOADS AND BEND CHECKS SHOULD BE ADDED\r\n" +
                "     USING THE SPECIAL BASETYPES RATHER THAN VIA\r\n" +
                "     COVER SHEET NOTES\r\n" +
                "   - SEPERATION IS ONLY ALLOWED ON ROOFS.\r\n" +
                "   - IF THE JOIST DESIGNATION LL IS FROM\r\n" +
                "     SNOW, THE FLAT ROOF SNOW LOAD(Pf)\r\n" +
                "     MUST BE LESS THAN 30 PSF.\r\n" +
                "   - FOR SEPERATION TO OCCUR ON GIRDERS,\r\n" +
                "     THE DESIGNATION MUST BE IN TL/LL FORM\r\n" +
                "     (i.e. 54G7N12.5/5.8K).\r\n" +
                "OTHER NOTES:\r\n" +
                "   - SEQUENCES NEED TO BE BETWEEN THE { & } CHARACTERS");
        }
    }
}
