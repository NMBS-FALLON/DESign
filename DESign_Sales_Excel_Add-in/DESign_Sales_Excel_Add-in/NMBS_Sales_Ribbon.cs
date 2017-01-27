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
            // Determine the row of the first baseType since estimators dont always place the first baseType at the top
            bool firstBaseTypeReached = false;
            int firstBaseTypeRow = 0;

            int i = 4;
            while (firstBaseTypeReached == false)
            {
                if (baseTypesCells[i, 1] != null)
                {
                    firstBaseTypeReached = true;
                    firstBaseTypeRow = i;
                }
                i++;
            }

            // Create a list containing the number of rows between each baseType
            List<int> rowsPerBaseTypeList = new List<int>();
            int rowsPerBaseType = 1;
            for (i = firstBaseTypeRow; i < baseTypesCells.GetLength(0); i++)
            {
                if (baseTypesCells[i + 1, 1] == null)
                {
                    rowsPerBaseType++;
                }
                else
                {
                    rowsPerBaseTypeList.Add(rowsPerBaseType);
                    rowsPerBaseType = 1;
                }
                if (i == baseTypesCells.GetLength(0) - 1)
                {
                    rowsPerBaseTypeList.Add(rowsPerBaseType);
                }
            }

            // Now that we can break out the chunks of information for each baseType, we can create the list of baseTypeLines
            List<BaseType> baseTypes = new List<BaseType>();

            int rowCount = firstBaseTypeRow;
            foreach (int rowsForThisBaseType in rowsPerBaseTypeList)
            {
                BaseType baseType = new BaseType();

                baseType.Name = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount, 1] };
                baseType.Description = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount, 2] };
                baseType.BaseLengthFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 3] };
                baseType.BaseLengthIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 4] };
                baseType.TcxlQuantity = new IntWithUpdateCheck { Value = (int?)(double?)baseTypesCells[rowCount, 5] };
                baseType.TcxlLengthFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 6] };
                baseType.TcxlLengthIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 7] };
                baseType.TcxrQuantity = new IntWithUpdateCheck { Value = (int?)(double?)baseTypesCells[rowCount, 8] };
                baseType.TcxrLengthFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 9] };
                baseType.TcxrLengthIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 10] };
                baseType.SeatDepthLE = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 11] };
                baseType.SeatDepthRE = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 12] };
                baseType.BcxQuantity = new IntWithUpdateCheck { Value = (int?)(double?)baseTypesCells[rowCount, 13] };
                baseType.Uplift = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 14] };
                baseType.Erfos = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount, 26] };
                baseType.DeflectionTL = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 27] };
                baseType.DeflectionLL = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 28] };
                baseType.WnSpacing = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount, 29] };

                
                List<Load> loads = new List<Load>();
                List<StringWithUpdateCheck> notes = new List<StringWithUpdateCheck>();

                for (i = 0; i < rowsForThisBaseType; i++)
                {

                    Load load = new Load();
                    load.LoadInfoType = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 15] };
                    load.LoadInfoCategory = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 16] };
                    load.LoadInfoPosition = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 17] };
                    load.Load1Value = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 18] };
                    load.Load1DistanceFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 19] };
                    load.Load1DistanceIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 20] };
                    load.Load2Value = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 21] };
                    load.Load2DistanceFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 22] };
                    load.Load2DistanceIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 23] };
                    load.CaseNumber = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 24] };
                    load.LoadNote = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 25] };
                    if (load.IsNull == false)
                    {
                        loads.Add(load);
                    }

                    StringWithUpdateCheck note = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 30] };
                    if (note.Text != null && note.IsUpdated == false)
                    {
                        notes.Add(note);
                    }

                }
                baseType.Loads = loads;
                baseType.Notes = notes;

                baseTypes.Add(baseType);
                rowCount = rowCount + rowsForThisBaseType;
            }

            ///////////////////

            // Create a range for the 'Marks' tab
            Excel.Range marksRange = marksWS.UsedRange;

            // Create an object array containing all information from the 'Marks' tab, in the form of a multidimensional array [rows, column]
            object[,] marksCells = (object[,])marksRange.Value2;

            // Determine the row of the first mark since estimators dont always place the first mark at the top
            bool firstMarkReached = false;
            int firstMarkRow = 0;

            i = 4;
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
            List<Joist> joistLines = new List<Joist>();
            
            rowCount = firstMarkRow;
            foreach (int rowsForThisMark in rowsPerMarkList)
            {
                Joist joist = new Joist();
                joist.Mark = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 1] };
                joist.Quantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 3] };
                joist.Description = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 4] };
                joist.BaseLengthFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 5] };
                joist.BaseLengthIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 6] };
                joist.TcxlQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 7] };
                joist.TcxlLengthFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 8] };
                joist.TcxlLengthIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 9] };
                joist.TcxrQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 10] };
                joist.TcxrLengthFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 11] };
                joist.TcxrLengthIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 12] };
                joist.SeatDepthLE = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 13] };
                joist.SeatDepthRE = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 14] };
                joist.BcxQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 15] };
                joist.Uplift = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 16] };
                joist.Erfos = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 28] };
                joist.DeflectionTL = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 29] };
                joist.DeflectionLL = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 30] };
                joist.WnSpacing = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 31] };

                List<StringWithUpdateCheck> baseTypesOnMark = new List<StringWithUpdateCheck>();
                List<Load> loads = new List<Load>();
                List<StringWithUpdateCheck> notes = new List<StringWithUpdateCheck>();

                for (i = 0; i < rowsForThisMark; i++)
                {
                    StringWithUpdateCheck baseTypeOnMark = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 2] };
                    if (baseTypeOnMark.Text != null && baseTypeOnMark.IsUpdated == false)
                    {
                        baseTypesOnMark.Add(baseTypeOnMark);
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
                joist.BaseTypesOnMark = baseTypesOnMark;
                joist.Loads = loads;
                joist.Notes = notes;

                joistLines.Add(joist);
                rowCount = rowCount + rowsForThisMark;
            }

            

            Takeoff takeoff = new Takeoff();

            takeoff.BaseTypes = baseTypes;
            takeoff.Joists = joistLines;

            foreach(Joist joistLine in takeoff.Joists)
            {
                foreach(Load load in joistLine.Loads)
                {
                    MessageBox.Show(string.Format("{0} : {1} : {2} : {3} : {4} : {5} : {6}", joistLine.Mark.Text, load.LoadInfoType.Text, load.LoadInfoCategory.Text, load.LoadInfoPosition.Text, load.Load1Value.Value, load.Load1DistanceFt.Value, load.Load1DistanceIn.Value));
                }
            }

        }




    }
}
