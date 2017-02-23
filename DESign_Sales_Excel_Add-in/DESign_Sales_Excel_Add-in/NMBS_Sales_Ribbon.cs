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

        private void btnCreateBasetypeDropDown_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook oWB = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel._Worksheet baseTypesWS = (Excel._Worksheet)oWB.Worksheets["Base Types"];
            Excel._Worksheet marksWS = (Excel._Worksheet)oWB.Worksheets["Marks"];


            // Create a range for the 'BaseLine' tab 
            int rows = baseTypesWS.UsedRange.Rows.Count;
            object[,] baseTypesArray = baseTypesWS.Range["A4:A" + rows].Value2;

            List<string> baseTypes = new List<string>();
            foreach(object baseType in baseTypesArray)
            {
                if(baseType != null && (string)baseType != "" && baseType.ToString().Contains("[ALL") == false)
                {
                    baseTypes.Add((string)(baseType));
                }
            }
            baseTypes.Add("'");

            rows = marksWS.UsedRange.Rows.Count;
            Excel.Range marksWSBaseTypeRange = marksWS.Range["B5:B" + rows];
            string baseTypesCommaSeperated = string.Join(",", baseTypes);
            MessageBox.Show(baseTypesCommaSeperated);
            marksWSBaseTypeRange.Validation.Delete();
            marksWSBaseTypeRange.Validation.Add(Excel.XlDVType.xlValidateList,
                Excel.XlDVAlertStyle.xlValidAlertInformation,
                Excel.XlFormatConditionOperator.xlBetween,
                baseTypesCommaSeperated);
                

        }
    }
}
