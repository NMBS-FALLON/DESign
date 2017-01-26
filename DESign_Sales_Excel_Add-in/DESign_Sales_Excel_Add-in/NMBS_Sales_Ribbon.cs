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
            // Initialize the necessary Excel objects:
            Excel.Application oXL = Globals.ThisAddIn.Application;
            Excel._Workbook oWB = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel._Worksheet marksWS = (Excel._Worksheet)oWB.Worksheets["Marks"];
            Excel._Worksheet baseTypesWS = (Excel._Worksheet)oWB.Worksheets["Base Types"];


            // Create a range for the 'BaseLine' tab 
            Excel.Range baseTypesRange = baseTypesWS.UsedRange;

            //Create an object array containing all information from the 'Base Types' tab, in the form of a multidimensional array [row, column]
            object[,] baseTypesCells = (object[,])baseTypesRange.Value2;

            ///////////////////
            // CODE FOR Creating BaseTypes 
            ///////////////////
            
            // Create a range for the 'Marks' tab
            Excel.Range marksRange = marksWS.UsedRange;

            // Create an object array containing all information from the 'Marks' tab, in the form of a multidimensional array [rows, column]
            object[,] marksCells = (object[,])marksRange.Value2;

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

            // Create a list containing the number of rows between each mark
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

            // Now that we can break out the chunks of information for each mark, we can create the list of joistLines
            List<JoistLine> joistLines = new List<JoistLine>();
            
            int rowCount = firstMarkRow;
            foreach (int rowsForThisMark in rowsPerMarkList)
            {
                JoistLine joistLine = new JoistLine();
                joistLine.Mark = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 1] };
                joistLine.Quantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 3] };
                joistLine.Description = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 4] };
                joistLine.BaseLengthFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 5] };
                joistLine.BaseLengthIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 6] };
                joistLine.TcxlQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 7] };
                joistLine.TcxlLengthFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 8] };
                joistLine.TcxlLengthIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 9] };
                joistLine.TcxrQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 10] };
                joistLine.TcxrLengthFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 11] };
                joistLine.TcxrLengthIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 12] };
                joistLine.SeatDepthLE = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 13] };
                joistLine.SeatDepthRE = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 14] };
                joistLine.BcxQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 15] };
                joistLine.Uplift = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 16] };
                joistLine.Erfos = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 28] };
                joistLine.DeflectionTL = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 29] };
                joistLine.DeflectionLL = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 30] };
                joistLine.WnSpacing = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 31] };

                List<StringWithUpdateCheck> baseTypes = new List<StringWithUpdateCheck>();
                List<Load> loads = new List<Load>();
                List<StringWithUpdateCheck> notes = new List<StringWithUpdateCheck>();

                for (i = 0; i < rowsForThisMark; i++)
                {
                    StringWithUpdateCheck baseType = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 2] };
                    if (baseType.Text != null && baseType.IsUpdated == false)
                    {
                        baseTypes.Add(baseType);
                    }



                    Load load = new Load();
                    load.LoadInfoType = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 17]};
                    load.LoadInfoCategory = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 18] };
                    load.LoadInfoPosition = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 19] };
                    load.Load1Value = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 20] };
                    load.Load1DistanceFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 21] };
                    load.Load1DistanceIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 22] };
                    load.Load2Value = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 23] };
                    load.Load2DistanceFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 24] };
                    load.Load2DistanceIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 25] };
                    load.CaseNumber = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 26] };
                    load.LoadNote = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 27] };
                    if(load.IsNull == false)
                    {
                        loads.Add(load);
                    }

                    StringWithUpdateCheck note = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 32] };
                    if (note.Text != null && note.IsUpdated == false)
                    {
                        notes.Add(note);
                    }

                }
                joistLine.BaseTypes = baseTypes;
                joistLine.Loads = loads;
                joistLine.Notes = notes;

                joistLines.Add(joistLine);
                rowCount = rowCount + rowsForThisMark;
            }

            

            Takeoff takeoff = new Takeoff();

            takeoff.JoistLines = joistLines;

            foreach(JoistLine joistLine in takeoff.JoistLines)
            {
                foreach(Load load in joistLine.Loads)
                {
                    MessageBox.Show(string.Format("{0} : {1} : {2} : {3} : {4} : {5} : {6}", joistLine.Mark.Text, load.LoadInfoType.Text, load.LoadInfoCategory.Text, load.LoadInfoPosition.Text, load.Load1Value.Value, load.Load1DistanceFt.Value, load.Load1DistanceIn.Value));
                }
            }

        }




    }
}
