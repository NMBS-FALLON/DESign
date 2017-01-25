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
            Excel.Application oXL = Globals.ThisAddIn.Application;
            Excel._Workbook oWB = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel._Worksheet marksWS = (Excel._Worksheet)oWB.Worksheets["Marks"];
            Excel._Worksheet baseTypesWS = (Excel._Worksheet)oWB.Worksheets["Base Types"];

            Excel.Range marksRange = marksWS.UsedRange;

            object[,] marksCells = (object[,])marksRange.Value2;
            List<string> listOfMarks = new List<string>();
            List<int> listOfSpaces = new List<int>();

            // Determine the row of the first mark since estimators dont always place the first mark at the top
            bool firstMarkReached = false;
            int firstMarkRow = 0;

            int i = 4;
            while(firstMarkReached ==false)
            {
                if (marksCells[i, 1] != null)
                {
                    firstMarkReached = true;
                    firstMarkRow = i;
                }
                i++;
            }



            // Create a list containing all information for each mark in an object array

            List<int> rowsPerMarkList = new List<int>();


            int rowsPerMark = 1;
            for (i = firstMarkRow; i < marksCells.GetLength(0); i++)
            {
                if (marksCells[i + 1, 1] == null)
                {
                    rowsPerMark++;
                }
                else
                {
                    rowsPerMarkList.Add(rowsPerMark);
                    rowsPerMark = 1;
                }
                if(i == marksCells.GetLength(0)-1)
                {
                    rowsPerMarkList.Add(rowsPerMark);
                }

            }
            foreach (int row in rowsPerMarkList)
            {
                MessageBox.Show(row.ToString());
            }

        }




    }
}
