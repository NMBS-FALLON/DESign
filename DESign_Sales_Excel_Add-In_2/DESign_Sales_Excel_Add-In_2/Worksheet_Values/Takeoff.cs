using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace DESign_Sales_Excel_Add_In_2.Worksheet_Values
{
    public class Takeoff
    {
        public double? SDS { get; set; }
        public class Sequence
        {

            public StringWithUpdateCheck Name { get; set; }
            public List<Joist> Joists { get; set; }
            public List<Bridging> Bridging { get; set; }
            private bool seperateSeismic = false;
            
            public bool SeperateSeismic
            {
                get
                {
                    return seperateSeismic;
                }
                set
                {
                    seperateSeismic = value;
                }
            }
            public double? SDS { get; set; }
        }

        public List<Bridging> Bridging { get; set; }

        public List<BaseType> BaseTypes { get; set; }

        public List<Sequence> Sequences { get; set; }


        // Initialize the necessary Excel objects:
        Excel.Application oXL = Globals.ThisAddIn.Application;
        Excel.Workbook workbook;
        Excel._Workbook oWB = Globals.ThisAddIn.Application.ActiveWorkbook;

        public Takeoff ImportTakeoff()
        {
            //
            Excel._Worksheet marksWS = (Excel._Worksheet)oWB.Worksheets["Marks"];
            Excel._Worksheet baseTypesWS = (Excel._Worksheet)oWB.Worksheets["Base Types"];
            Excel._Worksheet cover = (Excel._Worksheet)oWB.Worksheets["Cover"];

            double? sds = null;
            if (cover.Range["K12"].Value != null && cover.Range["K12"].Value.Contains("SDS"))
            {
                sds = cover.Range["M12"].Value;
            }

            bool bridgingSheetExists = false;
            foreach(Excel.Worksheet sheet in oWB.Sheets)
            {
                if (sheet.Name.Equals("Bridging"))
                {
                    bridgingSheetExists = true;
                }
            }

            if (bridgingSheetExists == false)
            {
                oWB.Worksheets.Add(After: oWB.Worksheets[baseTypesWS.Index]);
                oWB.Worksheets[baseTypesWS.Index + 1].Name = "Bridging";
                Excel._Worksheet bridgeWS = (Excel._Worksheet)oWB.Worksheets["Bridging"];
                bridgeWS.Cells[1, 1] = "Temp";
                bridgeWS.Cells[10, 8] = "Temp";
                
            }
            
            Excel._Worksheet bridgingWS = (Excel._Worksheet)oWB.Worksheets["Bridging"];

            ///// GET BRIDGING ////

            List<Bridging> bridging = new List<Bridging>();
            Excel.Range bridgingRange = bridgingWS.UsedRange;

            object[,] bridgingCells = (object[,])bridgingRange.Value2;

            string bridgingSequence = "";
            string size = "";
            string type = "";
            double rows = 0.0;
            double length = 0.0;

            int startRow = 5;
            int lastRow = bridgingCells.GetLength(0);
            for (int row = startRow; row <= lastRow; row++)
            {
                if (bridgingCells[row, 2] != null && bridgingCells[row, 2].ToString() != "")
                {
                    bridgingSequence = bridgingCells[row, 2].ToString();
                }

                if (bridgingCells[row, 3] != null && bridgingCells[row, 3].ToString() != "")
                {
                    size = bridgingCells[row, 3].ToString();
                }

                if (bridgingCells[row, 4] != null && bridgingCells[row, 4].ToString() != "")
                {
                    type = bridgingCells[row, 4].ToString();
                }

                rows = Convert.ToDouble(bridgingCells[row, 5]);
                length = Convert.ToDouble(bridgingCells[row, 6]);

                Bridging br = new Bridging();
                br.Sequence = bridgingSequence;
                br.Size = size;
                br.HorX = type;
                br.PlanFeet = rows * length * 1.02;

                bridging.Add(br);
            }

            bridging =
                (from br in bridging
                 group br by new
                 {
                     br.Sequence,
                     br.Size,
                     br.HorX,
                 } into brcs
                 select new Bridging()
                 {
                     Sequence = brcs.Key.Sequence,
                     Size = brcs.Key.Size,
                     HorX = brcs.Key.HorX,
                     PlanFeet = brcs.Sum(br => br.PlanFeet)
                 }).ToList();

            // Create a range for the 'BaseLine' tab 
            Excel.Range baseTypesRange = baseTypesWS.UsedRange;

            //Create an object array containing all information from the 'Base Types' tab, in the form of a multidimensional array [row, column]
            object[,] baseTypesCells = (object[,])baseTypesRange.Value2;

            //CHANGE ALL CELLS WITH "" TO NULL
            for (int row = 1; row <= baseTypesCells.GetLength(0); row++)
            {
                for (int col = 1; col <= baseTypesCells.GetLength(1); col++)
                {
                    if (baseTypesCells[row, col] is string)
                    {
                        if (Regex.Replace((string)baseTypesCells[row, col], @"\s+", "") == "")
                        {
                            baseTypesCells[row, col] = null;
                        }
                    }
                }
            }

            //Create a multidemnsional bool array that is true if the cell is highlighted (i.e. estimator marked it as updated) and false if it is not highlighted (i.e. cell has not been updated).
            int numRows = baseTypesRange.Rows.Count;
            int numColumns = baseTypesRange.Columns.Count;

            bool[,] isUpdated = new bool[numRows, numColumns];
            for (int row = 1; row <= numRows; row++)
            {
                for (int col = 1; col <= numColumns; col++)
                {
                    if (baseTypesRange[row, col].Interior.ColorIndex != -4142)
                    {
                        isUpdated[row - 1, col - 1] = true;
                    }
                    else
                    {
                        isUpdated[row - 1, col - 1] = false;
                    }
                }
            }

            ///////////////////

            // Determine the row of the first baseType since estimators dont always place the first baseType at the top
            bool firstBaseTypeReached = false;
            int firstBaseTypeRow = 4;

            int i = 4;
            while (firstBaseTypeReached == false && i < baseTypesCells.GetLength(0))
            {
                if (baseTypesCells != null)
                {
                    if (baseTypesCells[i, 1] != null)
                    {
                        firstBaseTypeReached = true;
                        firstBaseTypeRow = i;
                    }
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
            if (rowCount != 0)
            {
                bool errorMessageShown2 = false;
                foreach (int rowsForThisBaseType in rowsPerBaseTypeList)
                {
                    BaseType baseType = new BaseType();
                    try
                    {
                        baseType.Name = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount, 1], IsUpdated = isUpdated[rowCount - 1, 0] };
                        baseType.Description = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount, 2], IsUpdated = isUpdated[rowCount - 1, 1] };
                        baseType.BaseLengthFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 3], IsUpdated = isUpdated[rowCount - 1, 2] };
                        baseType.BaseLengthIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 4], IsUpdated = isUpdated[rowCount - 1, 3] };
                        baseType.TcxlQuantity = new IntWithUpdateCheck { Value = (int?)(double?)baseTypesCells[rowCount, 5], IsUpdated = isUpdated[rowCount - 1, 4] };
                        baseType.TcxlLengthFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 6], IsUpdated = isUpdated[rowCount - 1, 5] };
                        baseType.TcxlLengthIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 7], IsUpdated = isUpdated[rowCount - 1, 6] };
                        baseType.TcxrQuantity = new IntWithUpdateCheck { Value = (int?)(double?)baseTypesCells[rowCount, 8], IsUpdated = isUpdated[rowCount - 1, 7] };
                        baseType.TcxrLengthFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 9], IsUpdated = isUpdated[rowCount - 1, 8] };
                        baseType.TcxrLengthIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 10], IsUpdated = isUpdated[rowCount - 1, 9] };
                        baseType.SeatDepthLE = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 11], IsUpdated = isUpdated[rowCount - 1, 10] };
                        baseType.SeatDepthRE = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 12], IsUpdated = isUpdated[rowCount - 1, 11] };
                        if (Convert.ToString(baseTypesCells[rowCount, 13]) == "<ALL>")
                        {
                            baseType.BcxQuantity = new IntWithUpdateCheck { Value = -1, IsUpdated = isUpdated[rowCount - 1, 12] };
                        }
                        else
                        {
                            baseType.BcxQuantity = new IntWithUpdateCheck { Value = (int?)(double?)baseTypesCells[rowCount, 13], IsUpdated = isUpdated[rowCount - 1, 12] };
                        }
                        baseType.Uplift = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 14], IsUpdated = isUpdated[rowCount - 1, 13] };
                        baseType.Erfos = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount, 26], IsUpdated = isUpdated[rowCount - 1, 25] };
                        baseType.DeflectionTL = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 27], IsUpdated = isUpdated[rowCount - 1, 26] };
                        baseType.DeflectionLL = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount, 28], IsUpdated = isUpdated[rowCount - 1, 27] };
                        baseType.WnSpacing = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount, 29], IsUpdated = isUpdated[rowCount - 1, 28] };


                        List<Load> loads = new List<Load>();
                        List<StringWithUpdateCheck> notes = new List<StringWithUpdateCheck>();

                        for (i = 0; i < rowsForThisBaseType; i++)
                        {

                            Load load = new Load();

                            load.LoadInfoType = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 15], IsUpdated = isUpdated[rowCount + i - 1, 14] };
                            load.LoadInfoCategory = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 16], IsUpdated = isUpdated[rowCount + i - 1, 15] };
                            load.LoadInfoPosition = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 17], IsUpdated = isUpdated[rowCount + i - 1, 16] };
                            load.Load1Value = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 18], IsUpdated = isUpdated[rowCount + i - 1, 17] };
                            if (baseTypesCells[rowCount + i, 19] is double)
                            {
                                load.Load1DistanceFt = new StringWithUpdateCheck { Text = Convert.ToString((double?)baseTypesCells[rowCount + i, 19]), IsUpdated = isUpdated[rowCount + i - 1, 18] };
                            }
                            else
                            {
                                load.Load1DistanceFt = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 19], IsUpdated = isUpdated[rowCount + i - 1, 18] };
                            }

                            load.Load1DistanceIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 20], IsUpdated = isUpdated[rowCount + i - 1, 19] };
                            load.Load2Value = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 21], IsUpdated = isUpdated[rowCount + i - 1, 20] };
                            if (baseTypesCells[rowCount + i, 22] is double)
                            {
                                load.Load2DistanceFt = new StringWithUpdateCheck { Text = Convert.ToString((double?)baseTypesCells[rowCount + i, 22]), IsUpdated = isUpdated[rowCount + i - 1, 21] };
                            }
                            else
                            {
                                load.Load2DistanceFt = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 22], IsUpdated = isUpdated[rowCount + i - 1, 21] };
                            }
                            load.Load2DistanceIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 23], IsUpdated = isUpdated[rowCount + i - 1, 22] };
                            load.CaseNumber = new DoubleWithUpdateCheck { Value = ToNullableDouble((string)baseTypesCells[rowCount + i, 24]), IsUpdated = isUpdated[rowCount + i - 1, 23] };
                            load.LoadNote = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 25], IsUpdated = isUpdated[rowCount + i - 1, 24] };
                            if (load.IsNull == false)
                            {
                                loads.Add(load);
                            }

                            StringWithUpdateCheck note = new StringWithUpdateCheck { Text = (string)baseTypesCells[rowCount + i, 30], IsUpdated = isUpdated[rowCount + i - 1, 29] };
                            if (note.Text != null)
                            {
                                notes.Add(note);
                            }
                            if (note.Text == null && note.IsUpdated == true)
                            {
                                notes.Add(note);
                            }

                        }
                        baseType.Loads = loads;
                        baseType.Notes = notes;

                        baseTypes.Add(baseType);
                        rowCount = rowCount + rowsForThisBaseType;
                    }
                    catch
                    {
                        if (errorMessageShown2 == false)
                        {
                            MessageBox.Show(String.Format(@"BASETYPE {0}:
    ISSUE PULLING INFO FROM BASE TYPES TAB.
    PLEASE CHECK THAT COLUMNS ARE FILLED IN CORRECTLY.
    THIS MUST BE FIXED BEFORE CONVERTING THE TAKEOFF.", baseType.Name.Text));
                            errorMessageShown2 = true;
                        }
                    }
                }
            }

            ///////////////////

            // Create a range for the 'Marks' tab
            Excel.Range marksRange = marksWS.UsedRange;

            // Create an object array containing all information from the 'Marks' tab, in the form of a multidimensional array [rows, column]
            object[,] marksCells = (object[,])marksRange.Value2;

            //CHANGE ALL CELLS WITH "" TO NULL
            for (int row = 1; row <= marksCells.GetLength(0); row++)
            {
                for (int col = 1; col <= marksCells.GetLength(1); col++)
                {
                    if (marksCells[row, col] is string)
                    {
                        if (Regex.Replace((string)marksCells[row, col], @"\s+", "") == "")
                        {
                            marksCells[row, col] = null;
                        }
                    }
                }
            }

            //Create a multidemnsional bool array that is true if the cell is highlighted (i.e. estimator marked it as updated) and false if it is not highlighted (i.e. cell has not been updated).
            numRows = marksRange.Rows.Count;
            numColumns = marksRange.Columns.Count;

            isUpdated = new bool[numRows, numColumns];
            for (int row = 1; row <= numRows; row++)
            {
                for (int col = 1; col <= numColumns; col++)
                {
                    if (marksRange[row, col].Interior.ColorIndex != -4142)
                    {
                        isUpdated[row - 1, col - 1] = true;
                    }
                    else
                    {
                        isUpdated[row - 1, col - 1] = false;
                    }
                }
            }



            // Determine the row of the first mark or sequence since estimators dont always place it at the top
            bool firstLineReached = false;
            int firstMarkRow = 4;

            i = 4;
            while (firstLineReached == false && i < marksCells.GetLength(0))
            {
                if (marksCells != null)
                {
                    if (marksCells[i, 1] != null)
                    {
                        firstLineReached = true;
                        firstMarkRow = i;
                    }
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
                if (i == marksCells.GetLength(0) - 1)
                {
                    rowsPerMarkList.Add(rowsPerMark);
                }
            }

            // Now that we can break out the chunks of information for each mark, we can create the list of joistLines
            List<Joist> joistLines = new List<Joist>();

            rowCount = firstMarkRow;
            bool errorMessageShown = false;
            foreach (int rowsForThisMark in rowsPerMarkList)
            {

                Joist joist = new Joist();
                try {
                    joist.Mark = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 1], IsUpdated = isUpdated[rowCount - 1, 0] };
                    joist.Quantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 3], IsUpdated = isUpdated[rowCount - 1, 2] };
                    joist.Description = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 4], IsUpdated = isUpdated[rowCount - 1, 3] };
                    joist.BaseLengthFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 5], IsUpdated = isUpdated[rowCount - 1, 4] };
                    joist.BaseLengthIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 6], IsUpdated = isUpdated[rowCount - 1, 5] };
                    joist.TcxlQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 7], IsUpdated = isUpdated[rowCount - 1, 6] };
                    joist.TcxlLengthFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 8], IsUpdated = isUpdated[rowCount - 1, 7] };
                    joist.TcxlLengthIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 9], IsUpdated = isUpdated[rowCount - 1, 8] };
                    joist.TcxrQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 10], IsUpdated = isUpdated[rowCount - 1, 9] };
                    joist.TcxrLengthFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 11], IsUpdated = isUpdated[rowCount - 1, 10] };
                    joist.TcxrLengthIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 12], IsUpdated = isUpdated[rowCount - 1, 11] };
                    joist.SeatDepthLE = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 13], IsUpdated = isUpdated[rowCount - 1, 12] };
                    joist.SeatDepthRE = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 14], IsUpdated = isUpdated[rowCount - 1, 13] };
                    if (Convert.ToString(marksCells[rowCount, 15]) == "<ALL>")
                    {
                        joist.BcxQuantity = new IntWithUpdateCheck { Value = joist.Quantity.Value * 2, IsUpdated = isUpdated[rowCount - 1, 14] };
                    }
                    else
                    {
                        joist.BcxQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 15], IsUpdated = isUpdated[rowCount - 1, 14] };
                    }
                    joist.Uplift = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 16], IsUpdated = isUpdated[rowCount - 1, 15] };
                    joist.Erfos = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 28], IsUpdated = isUpdated[rowCount - 1, 27] };
                    joist.DeflectionTL = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 29], IsUpdated = isUpdated[rowCount - 1, 28] };
                    joist.DeflectionLL = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount, 30], IsUpdated = isUpdated[rowCount - 1, 29] };
                    joist.WnSpacing = new StringWithUpdateCheck { Text = (string)marksCells[rowCount, 31], IsUpdated = isUpdated[rowCount - 1, 30] };

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
                        load.LoadInfoType = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 17], IsUpdated = isUpdated[rowCount + i - 1, 16] };
                        load.LoadInfoCategory = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 18], IsUpdated = isUpdated[rowCount + i - 1, 17] };
                        load.LoadInfoPosition = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 19], IsUpdated = isUpdated[rowCount + i - 1, 18] };
                        load.Load1Value = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 20], IsUpdated = isUpdated[rowCount + i - 1, 19] };
                        if (marksCells[rowCount + i, 21] is double)
                        {
                            load.Load1DistanceFt = new StringWithUpdateCheck { Text = Convert.ToString((double?)marksCells[rowCount + i, 21]), IsUpdated = isUpdated[rowCount + i - 1, 20] };
                        }
                        else
                        {
                            load.Load1DistanceFt = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 21], IsUpdated = isUpdated[rowCount + i - 1, 20] };
                        }
                        load.Load1DistanceIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 22], IsUpdated = isUpdated[rowCount + i - 1, 21] };
                        load.Load2Value = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 23], IsUpdated = isUpdated[rowCount + i - 1, 22] };
                        if (marksCells[rowCount + i, 21] is double)
                        {
                            load.Load2DistanceFt = new StringWithUpdateCheck { Text = Convert.ToString((double?)marksCells[rowCount + i, 24]), IsUpdated = isUpdated[rowCount + i - 1, 23] };
                        }
                        else
                        {
                            load.Load2DistanceFt = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 24], IsUpdated = isUpdated[rowCount + i - 1, 23] };
                        }
                        load.Load2DistanceIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 25], IsUpdated = isUpdated[rowCount + i - 1, 24] };
                        if (marksCells[rowCount + i, 26] is string)
                        {
                            load.CaseNumber = new DoubleWithUpdateCheck { Value = Convert.ToDouble(marksCells[rowCount + i, 26]), IsUpdated = isUpdated[rowCount + i - 1, 25] };
                        }
                        else {
                            load.CaseNumber = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 26], IsUpdated = isUpdated[rowCount + i - 1, 25] };
                        }
                        load.LoadNote = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 27], IsUpdated = isUpdated[rowCount + i - 1, 26] };
                        if (load.IsNull == false)
                        {
                            loads.Add(load);
                        }

                        StringWithUpdateCheck note = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 32], IsUpdated = isUpdated[rowCount + i - 1, 31] };
                        if (note.Text != null)
                        {
                            notes.Add(note);
                        }
                        if (note.Text == null && note.IsUpdated == true)
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
                catch
                {
                    if (errorMessageShown == false)
                    {
                        MessageBox.Show(String.Format(@"MARK {0}:
    ISSUE PULLING INFO FROM MARKS TAB.
    PLEASE CHECK THAT COLUMNS ARE FILLED IN CORRECTLY.
    THIS MUST BE FIXED BEFORE CONVERTING THE TAKEOFF.", joist.Mark.Text));
                        errorMessageShown = true;
                    }
                }
            }

            //Seperate Sequences
            List<Sequence> sequences = new List<Sequence>();
            var sequenceQuery = from jst in joistLines
                                where jst.Mark.Text.Contains("{") && jst.Mark.Text.Contains("}")
                                select jst;
            if (!sequenceQuery.Any()) //No named sequences on takeoff
            {
                Sequence sequence = new Sequence();
                sequence.Name = new StringWithUpdateCheck { Text = "" };
                sequence.Joists = joistLines;
                sequences.Add(sequence);
            }
            else
            {


                if (joistLines[0].Quantity.Value != null || joistLines[0].Description.Text != null)
                {
                    MessageBox.Show("Please name your first sequence");
                }
                else
                {
                    Sequence sequence = new Sequence();
                    sequence.Name = new StringWithUpdateCheck { Text = "" };


                    int jstIndex = 0;

                    for (int joistIndex = jstIndex; joistIndex < joistLines.Count; joistIndex++)
                    {
                        if (joistLines[joistIndex].Quantity.Value == null && joistLines[joistIndex].Description.Text == null)
                        {
                            sequence.Joists = new List<Joist>();
                            sequence.Name.Text = joistLines[joistIndex].Mark.Text;
                            sequence.Name.IsUpdated = joistLines[joistIndex].Mark.IsUpdated;
                        }
                        else
                        {
                            Joist joist = new Joist();
                            joist = joistLines[joistIndex];
                            sequence.Joists.Add(joist);
                        }
                        if (joistIndex + 1 < joistLines.Count)
                        {
                            if (joistLines[joistIndex + 1].Quantity.Value == null && joistLines[joistIndex + 1].Description.Text == null && joistLines[joistIndex + 1].BaseTypesOnMark.Count == 0)
                            {
                                Sequence coppiedSequence = new Sequence();
                                List<Joist> newJoists = new List<Joist>();
                                foreach (Joist jst in sequence.Joists)
                                {
                                    Joist newJoist = new Joist();
                                    newJoist = DeepClone(jst);
                                    newJoists.Add(newJoist);
                                }
                                StringWithUpdateCheck coppiedName = new StringWithUpdateCheck();
                                coppiedName = DeepClone(sequence.Name);
                                coppiedSequence.Name = coppiedName;
                                coppiedSequence.Joists = newJoists;
                                sequences.Add(coppiedSequence);
                            }
                        }
                        else
                        {
                            Sequence coppiedSequence = new Sequence();
                            List<Joist> newJoists = new List<Joist>();
                            foreach (Joist jst in sequence.Joists)
                            {
                                Joist newJoist = new Joist();
                                newJoist = DeepClone(jst);
                                newJoists.Add(newJoist);
                            }
                            StringWithUpdateCheck coppiedName = new StringWithUpdateCheck();
                            coppiedName = DeepClone(sequence.Name);
                            coppiedSequence.Name = coppiedName;
                            coppiedSequence.Joists = newJoists;
                            sequences.Add(coppiedSequence);
                        }
                    }
                }
            }

            Takeoff takeoff = new Takeoff();
            takeoff.SDS = sds;
            takeoff.BaseTypes = baseTypes;
            takeoff.Sequences = sequences;
            foreach (Bridging br in bridging)
            {
                br.PlanFeet = Math.Ceiling(br.PlanFeet / 20.0) * 20.0;
            }
            takeoff.Bridging = bridging;

            


            foreach (Sequence seq in takeoff.Sequences)
            {



                // ADD BASE TYPES TO JOISTS

                foreach (var joist in seq.Joists)
                {
                    foreach (var baseType in joist.BaseTypesOnMark)
                    {
                        // Select the matching base type. THIS WILL NEED TO BE UPDATED TO CHECK FOR TYPOS AND TO MAKE SURE BASETYPES EXIST
                        var matchedBaseType = from bT in takeoff.BaseTypes
                                              where bT.Name.Text == baseType.Text
                                              select bT;

                        foreach (var bT in matchedBaseType)
                        {
                            //ADD VALUES    ???DO I NEED TO CHECK ANYTHING THAT MAY BE UPDATED??? IF SO HOW TO IMPLEMENT?
                            AddBaseType(joist, bT);
                        }

                    }
                    //ADD BASETYPES DESIGNATED [ALL], [ALL J] (ALL JOISTS), & [ALL G] (ALL GIRDERS). 

                    var all = from bT1 in baseTypes
                              where bT1.Name.Text != null
                              where bT1.Name.Text.ToUpper().Replace(" ", string.Empty).Contains("[ALL]")
                              select bT1;

                    var allJoist = from bT1 in baseTypes
                                   where bT1.Name.Text != null
                                   where bT1.Name.Text.ToUpper().Replace(" ", string.Empty).Contains("[ALLJ]")
                                   select bT1;

                    var allGirder = from bT1 in baseTypes
                                    where bT1.Name.Text != null
                                    where bT1.Name.Text.ToUpper().Replace(" ", string.Empty).Contains("[ALLG]")
                                    select bT1;

                    var allSequence = from bT1 in baseTypes
                                      where bT1.Name.Text != null
                                      where bT1.Name.Text.ToUpper().Replace(" ", string.Empty).Contains("[ALL:" + 
                                             seq.Name.Text.ToUpper().Replace(" ", string.Empty) +"]")
                                      select bT1;

                    var allJoistSequence = from bT1 in baseTypes
                                           where bT1.Name.Text != null
                                           where bT1.Name.Text.ToUpper().Replace(" ", string.Empty).Contains("[ALLJ:" +
                                                  seq.Name.Text.ToUpper().Replace(" ", string.Empty) + "]")
                                           select bT1;

                    var allGirderSequence = from bT1 in baseTypes
                                            where bT1.Name.Text != null
                                            where bT1.Name.Text.ToUpper().Replace(" ", string.Empty).Contains("[ALLG:" +
                                                  seq.Name.Text.ToUpper().Replace(" ", string.Empty) + "]")
                                            select bT1;

                    if (all.Any())
                    {
                        foreach (BaseType bT1 in all)
                        {
                            AddBaseType(joist, DeepClone(bT1));
                        }
                    }

                    if (allJoist.Any() && joist.IsGirder == false)
                    {
                        foreach (BaseType bT1 in allJoist)
                        {
                            AddBaseType(joist, DeepClone(bT1));
                        }
                    }

                    if (allGirder.Any() && joist.IsGirder == true)
                    {
                        foreach (BaseType bT1 in allGirder)
                        {
                            AddBaseType(joist, DeepClone(bT1));
                        }
                    }

                    if (allSequence.Any())
                    {
                        foreach (BaseType bT1 in allSequence)
                        {
                            AddBaseType(joist, DeepClone(bT1));
                        }
                    }

                    if (allJoistSequence.Any() && joist.IsGirder == false)
                    {
                        foreach (BaseType bT1 in allJoistSequence)
                        {
                            AddBaseType(joist, DeepClone(bT1));
                        }
                    }

                    if (allGirderSequence.Any() && joist.IsGirder == true)
                    {
                        foreach (BaseType bT1 in allGirderSequence)
                        {
                            AddBaseType(joist, DeepClone(bT1));
                        }
                    }


                }
            }
            // Checks:
            string errors = "";
            int joistWithErrorCount = 0;

            var joistMarks = from seq in sequences
                             from jst in seq.Joists
                             select jst.Mark.Text;
            // check that there are no duplicate marks
            var markGroups = joistMarks.GroupBy(x => x)
                             .Where(g => g.Count() > 1);
            if (markGroups.Any())
            {
                foreach (var group in markGroups)
                {
                    errors += string.Format("  There are ({0}) marks labeled \"{1}\"\r\n\r\n", group.Count().ToString(), group.Key);
                }
            }

           foreach(BaseType bt in takeoff.BaseTypes)
            {
                if (bt.Errors.Count != 0)
                {
                    errors += string.Format("  BASETYPE {0}:\r\n", bt.Name.Text);
                    foreach (string error in bt.Errors)
                    {
                        errors += "      " + error + "\r\n";
                    }
                    errors += "\r\n";
                }
            }

            var baseTypeNames = from bt in baseTypes
                                select bt.Name.Text;


            // ADJUST SPECIAL LOADS
            // ..... FUTURE .....
            // NEED TO ADD CHECKS TO MAKE SURE ALL OF THE SPECIAL LOADS ARE PROVIDIG ACCURATE INFORMATION.
            foreach (Sequence sequence in takeoff.Sequences)
            {
                foreach (Joist joist in sequence.Joists)
                {
                    List<Load> newLoads = new List<Load>();
                    foreach (Load load in joist.Loads)
                    {
                        if (load.LoadInfoCategory.Text == "SMU")
                        {
                            load.Load1Value.Value = 1 * (int)Math.Ceiling((decimal)(load.Load1Value.Value * 0.7 / 1.0));
                            if (load.Load2Value.Value != null)
                            {
                                load.Load2Value.Value = 1 * (int)Math.Ceiling((decimal)(load.Load2Value.Value * 0.7 / 1.0));
                            }
                            load.LoadInfoCategory.Text = "SM";
                        }

                        if (load.LoadInfoCategory.Text == "WLU")
                        {
                            load.Load1Value.Value = 1 * (int)Math.Ceiling((decimal)(load.Load1Value.Value * 0.6 / 1.0));
                            if (load.Load2Value.Value != null)
                            {
                                load.Load2Value.Value = 1 * (int)Math.Ceiling((decimal)(load.Load2Value.Value * 0.6 / 1.0));
                            }
                            load.LoadInfoCategory.Text = "WL";
                        }

                        if (load.LoadInfoType.Text == "CMP")
                        {
                            if (joist.IsGirder == false)
                            {
                                load.Errors.Add("'CMP' LOAD CANNOT BE ADDED TO A JOIST");
                            }
                            else
                            {
                                if (load.Load1DistanceFt.Text.Replace(" ", "").ToUpper() == "ALL")
                                {
                                    int numPanelPoints = Convert.ToInt16(joist.Description.Text.Split(new string[] { "G", "N" }, StringSplitOptions.None)[1]) - 1;
                                    for (int j = 1; j <= numPanelPoints; j++)
                                    {
                                        Load ppLoad = DeepClone(load);
                                        ppLoad.LoadInfoType.Text = "C";
                                        ppLoad.Load1DistanceFt.Text = "P" + DeepClone(j).ToString();
                                        newLoads.Add(ppLoad);
                                    }
                                }
                                else
                                {
                                    string[] loadLocations = load.Load1DistanceFt.Text.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                    loadLocations = loadLocations.Select(loadLocation => "P" + loadLocation.Replace(" ", "")).ToArray();
                                    foreach (string loadLocation in loadLocations)
                                    {
                                        Load newLoad = DeepClone(load);
                                        newLoad.LoadInfoType.Text = "C";
                                        newLoad.Load1DistanceFt.Text = loadLocation;
                                        newLoads.Add(newLoad);
                                    }

                                }
                            }
                        }

                        if (load.LoadInfoType.Text == "CUP" || load.LoadInfoType.Text == "CUA")
                        {
                            double joistFt = joist.BaseLengthFt.Value == null ? 0.0 : (double)(joist.BaseLengthFt.Value);
                            double joistIn = joist.BaseLengthIn.Value == null ? 0.0 : (double)(joist.BaseLengthIn.Value);
                            double joistLengthInFt = joistFt + joistIn / 12.0;

                            double loadFt = load.Load1DistanceFt.Text == null ? 0.0 : Double.Parse(load.Load1DistanceFt.Text);
                            double loadIn = load.Load1DistanceIn.Value == null ? 0.0 : (double)(load.Load1DistanceIn.Value);
                            double spaceInFt = loadFt + loadIn / 12.0;

                            double ptLoad = load.Load1Value.Value == null ? 0.0 : (double)(load.Load1Value.Value);
                            double uniformLoadValue = Math.Ceiling(ptLoad / spaceInFt - ptLoad / joistLengthInFt);

                            Load uniformLoad = DeepClone(load);
                            uniformLoad.Load1Value.Value = uniformLoadValue;
                            uniformLoad.LoadInfoType.Text = "U";
                            uniformLoad.Load1DistanceFt.Text = null;
                            uniformLoad.Load1DistanceIn.Value = null;

                            Load cpLoad = DeepClone(load);
                            cpLoad.LoadInfoType.Text = load.LoadInfoType.Text == "CUP" ? "CP" : "CA";
                            cpLoad.Load1DistanceFt.Text = null;
                            cpLoad.Load1DistanceIn.Value = null;

                            newLoads.Add(uniformLoad);
                            newLoads.Add(cpLoad);
                        }
                    }

                    joist.Loads.AddRange(newLoads);

                    string[] loadInfoTypesToRemove = new string[] { "CMP", "CUP", "CUA" };
                    joist.Loads.RemoveAll(load => (loadInfoTypesToRemove.Contains(load.LoadInfoType.Text)));

                }
            }


            foreach (Sequence seq in sequences)
            {
                
                foreach (Joist joist in seq.Joists)
                {
                    foreach (var bt in joist.BaseTypesOnMark)
                    {
                        if (baseTypeNames.Contains(bt.Text) == false)
                        {
                            //errors += string.Format("\r\nMark {0}:\r\n    'Base Types' tab does not contain a definition for \"{1}\"\r\n\r\n", joist.Mark.Text, bt.Text);
                            joist.AddError(string.Format("'Base Types' tab does not contain a definition for \"{0}\"", bt.Text));
                        }
                    }
                    if (joist.Errors.Count != 0)
                    {
                        errors += string.Format("  MARK {0}:\r\n", joist.Mark.Text);
                        foreach (string error in joist.Errors)
                        {
                            errors += "      " + error + "\r\n";
                        }
                        errors += "\r\n";
                    }

                }
                
            }
            if (errors != "")
            {
                string filePath = Path.GetTempPath() + "Errors.txt";
                File.WriteAllText(filePath, "Takeoff Errors:\r\n\r\n" + errors);
                System.Diagnostics.Process.Start(filePath);
            }

            return takeoff;
        }

        public static double? ToNullableDouble(string s)
        {
            if (s == null) return null;
            double i;
            if (double.TryParse(s, out i)) return i;
            return null;
        }
        public void CreateOriginalTakeoff()
        {

            string excelPath = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllBytes(excelPath, DESign_Sales_Excel_Add_In.Properties.Resources.BLANK_SALES_BOM);

            Excel.Application oXL2 = new Excel.Application();
            Excel.Workbooks workbooks = oXL.Workbooks;
            workbook = workbooks.Open(excelPath);
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = new Excel.Worksheet();

            oXL2.Visible = false;




            int sheetIndex = 6;

            int sheetCount = 0;

            foreach (Sequence sequence in Sequences)
            {
                sheetCount++;
                sheetIndex++;
                Excel.Worksheet firstSheetOfSequence = workbook.Worksheets["J(BLANK)"];
                firstSheetOfSequence.Copy(Type.Missing, After: sheets[sheetIndex-1]);
                firstSheetOfSequence = workbook.Worksheets[sheetIndex];
                firstSheetOfSequence.Name = "J (" + Convert.ToString(sheetCount) + ")";
                sheet = workbook.Worksheets[sheetIndex];
                CellInsert(sheet, 5, 3, sequence.Name.Text, sequence.Name.IsUpdated);

                int row = 7;
                int pageRowCounter = 0;
                
                for (int markCounter = 0; markCounter < sequence.Joists.Count;)
                {
                    Joist joist = sequence.Joists[markCounter];

                    int maxRows = Math.Max(joist.Loads.Count, joist.Notes.Count);
                    if (maxRows > 32)
                    {
                        MessageBox.Show(String.Format("Mark {0} has too many loads on it.\r\n NOTE THAT THIS JOIST WILL NOT BE ADDED TO THE TAKEOFFF!\r\n Either add this joist manually or send to Darien to convert.",
                                                      joist.Mark.Text));
                        markCounter++;
                        goto SkipLoop;
                    }


                    pageRowCounter = pageRowCounter + Math.Max(Math.Max(joist.Loads.Count, joist.Notes.Count), 1) + 3;
                    if (pageRowCounter > 35)
                    {
                        sheetCount = sheetCount + 1;
                        Excel.Worksheet worksheet_copy = workbook.Worksheets["J(BLANK)"];
                        worksheet_copy.Copy(Type.Missing, After: sheets[sheetIndex]);
                        worksheet_copy = workbook.Worksheets[sheetIndex + 1];
                        worksheet_copy.Name = "J (" + Convert.ToString(sheetCount) + ")";
                        sheetIndex++;
                        sheet = workbook.Worksheets[sheetIndex];
                        row = 7;
                        pageRowCounter = 0;
                        goto SkipLoop;
                    }

                    CellInsert(sheet, row, 1, joist.Mark.Text, joist.Mark.IsUpdated);
                    CellInsert(sheet, row, 2, joist.Quantity.Value, joist.Quantity.IsUpdated);
                    CellInsert(sheet, row, 3, joist.DescriptionAdjusted.Text, joist.DescriptionAdjusted.IsUpdated);
                    CellInsert(sheet, row, 4, joist.BaseLengthFt.Value, joist.BaseLengthFt.IsUpdated);
                    CellInsert(sheet, row, 5, joist.BaseLengthIn.Value, joist.BaseLengthIn.IsUpdated);
                    CellInsert(sheet, row, 6, joist.TcxlQuantity.Value, joist.TcxlQuantity.IsUpdated);
                    CellInsert(sheet, row, 7, joist.TcxlLengthFt.Value, joist.TcxlLengthFt.IsUpdated);
                    CellInsert(sheet, row, 8, joist.TcxlLengthIn.Value, joist.TcxlLengthIn.IsUpdated);
                    CellInsert(sheet, row, 9, joist.TcxrQuantity.Value, joist.TcxrQuantity.IsUpdated);
                    CellInsert(sheet, row, 10, joist.TcxrLengthFt.Value, joist.TcxrLengthFt.IsUpdated);
                    CellInsert(sheet, row, 11, joist.TcxrLengthIn.Value, joist.TcxrLengthIn.IsUpdated);
                    CellInsert(sheet, row, 12, joist.SeatDepthLE.Value, joist.SeatDepthLE.IsUpdated);
                    CellInsert(sheet, row, 13, joist.SeatDepthRE.Value, joist.SeatDepthRE.IsUpdated);
                    CellInsert(sheet, row, 14, joist.BcxQuantity.Value, joist.BcxQuantity.IsUpdated);
                    CellInsert(sheet, row, 15, joist.Uplift.Value, joist.Uplift.IsUpdated);

                    int loadRow = row;
                    foreach (Load load in joist.Loads)
                    {

                        CellInsert(sheet, loadRow, 16, load.LoadInfoType.Text, load.LoadInfoType.IsUpdated);
                        CellInsert(sheet, loadRow, 17, load.LoadInfoCategory.Text, load.LoadInfoCategory.IsUpdated);
                        CellInsert(sheet, loadRow, 18, load.LoadInfoPosition.Text, load.LoadInfoPosition.IsUpdated);
                        CellInsert(sheet, loadRow, 19, load.Load1Value.Value, load.Load1Value.IsUpdated);
                        CellInsert(sheet, loadRow, 20, load.Load1DistanceFt.Text, load.Load1DistanceFt.IsUpdated);
                        CellInsert(sheet, loadRow, 21, load.Load1DistanceIn.Value, load.Load1DistanceIn.IsUpdated);
                        CellInsert(sheet, loadRow, 22, load.Load2Value.Value, load.Load2Value.IsUpdated);
                        CellInsert(sheet, loadRow, 23, load.Load2DistanceFt.Text, load.Load2DistanceFt.IsUpdated);
                        CellInsert(sheet, loadRow, 24, load.Load2DistanceIn.Value, load.Load2DistanceIn.IsUpdated);
                        CellInsert(sheet, loadRow, 25, load.CaseNumber.Value, load.CaseNumber.IsUpdated);
                        loadRow++;

                    }

                    int noteRow = row;
                    foreach (StringWithUpdateCheck note in joist.Notes)
                    {
                        CellInsert(sheet, noteRow, 26, note.Text, note.IsUpdated);
                        noteRow++;
                    }

                    markCounter++;
                    row = row + Math.Max(Math.Max(joist.Loads.Count, joist.Notes.Count), 1) + 3;

                SkipLoop:
                    ;
                }
            }


            //COPY COVER SHEET INTO NEW TAKEOFF
            Excel.Worksheet cover = oWB.Sheets["Cover"];
            CellInsert(cover, 2, 10, "=INDEX(INDIRECT(\"ProjectTypes[Category]\"),MATCH(INDIRECT(\"ProjectCat\"),INDIRECT(\"ProjectTypes[Type]\"),0))", false);
            cover.Copy(Type.Missing, After: workbook.Sheets["Cover"]);
            Excel.Worksheet oldCover = workbook.Sheets["Cover"];
            oXL.DisplayAlerts = false;
            oldCover.Delete();
            oXL.DisplayAlerts = true;
            Excel.Worksheet newCover = workbook.Sheets["Cover (2)"];
            newCover.Name = "Cover";

            int bridgingRow = 39;
            int columnIndex = 0;
            Bridging = Bridging.Where(br => !(br.Size == "" && br.HorX =="" && br.PlanFeet == 0.0)).ToList();
            var bridgingBySequence = Bridging.GroupBy(br => br.Sequence);
                
            foreach(var seq in bridgingBySequence)
            {
                if(bridgingRow + seq.Count() > 53)
                {
                    if (columnIndex == 0)
                    {
                        columnIndex = 4;
                        bridgingRow = 39;
                    }
                    else
                    {
                        MessageBox.Show("NOT ENOUGH ROOM FOR BRIDGING, PLEASE ADJUST BRIDGING MANUALLY");
                        if (bridgingRow < 55) { bridgingRow = 55; }
                    }
                }
                
                CellInsert(workbook.Sheets["Cover"], bridgingRow, 2+columnIndex, seq.Key, false);
                bridgingRow++;
                foreach(Bridging br in seq)
                {
                    CellInsert(workbook.Sheets["Cover"], bridgingRow, 2+columnIndex, br.Size, false);
                    CellInsert(workbook.Sheets["Cover"], bridgingRow, 3+columnIndex, br.HorX, false);
                    CellInsert(workbook.Sheets["Cover"], bridgingRow, 4+columnIndex, br.PlanFeet, false);
                    bridgingRow++;
                }
                bridgingRow++;
            }

            //COPY NOTE AND BRIDGING SHEETS INTO NEW TAKEOFF
            foreach(Excel.Worksheet s in oWB.Sheets)
            {
                if (s.Name.Contains("N (") && s.Name != "N (0)")
                {
                    s.Copy(Type.Missing, After: workbook.Sheets["Cover"]);
                }
                if (s.Name.Contains("Bridging"))
                {
                    s.Copy(Before: workbook.Sheets["Check List"]);
                }
            }

            newCover.Activate();

            Excel.Worksheet blankWS = workbook.Sheets["J(BLANK)"];
            oXL.DisplayAlerts = false;
            blankWS.Delete();
            oXL.DisplayAlerts = true;




            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsm)|*.xlsm";
            saveFileDialog.ShowDialog();
            if (saveFileDialog.FileName != "")
            {  
                workbook.CheckCompatibility = false;
                workbook.SaveAs(saveFileDialog.FileName);      
            }

            oXL2.Visible = true;
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(oWB);
            Marshal.ReleaseComObject(oXL2);
            GC.Collect();

        }
        private void CellInsert(Excel.Worksheet sheet, int row, int column, object o, bool isUpdated)
        {
            if (o == null) { }
            else
            {
                sheet.Cells[row, column] = o;
            }

            if (isUpdated == true)
            {
                workbook.Worksheets["HighlightedCell"].Range["A1"].Copy();
                sheet.Cells[row, column].PasteSpecial(Excel.XlPasteType.xlPasteFormats,
                                                      Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                
                if (column == 26)
                {
                    sheet.Range[sheet.Cells[row, 26], sheet.Cells[row, 29]].Merge();
                    sheet.Cells[row, 26].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                }

            }
        }

        public void SeperateSeismic()
        {
            foreach (Sequence sequence in Sequences)
            {
                if (sequence.SeperateSeismic == true)
                {
                    /*bool lc3Taken = false;
                    foreach (Joist joist in sequence.Joists)
                    {
                        var listOfLCs = from load in joist.Loads
                                        select Convert.ToInt32(load.CaseNumber.Value);


                        if (listOfLCs.Contains(3) == true)
                        {
                            lc3Taken = true;
                        }

                        
                    }
                    
                    if (lc3Taken == true)
                    {
                        MessageBox.Show(string.Format("Sequence {0}: LC 3 MUST BE AVAILABLE FOR SEISMIC SEPERATION",
                                        sequence.Name.Text));
                        //throw new SystemException();
                    }
                    */
                    
                    foreach (Joist joist in sequence.Joists)
                    {


                        //Determine if joist has Seismic Loads

                        var listOfLoadTypes = from load in joist.Loads
                                              select load.LoadInfoCategory.Text;
                        bool hasSeismic = false;
                        foreach (string type in listOfLoadTypes)
                        {
                            if (type == "SM")
                            {
                                hasSeismic = true;
                            }
                        }

                        if (hasSeismic)
                        {
                            //DETERMINE IF JOIST IS LOAD OVER LOAD. If not then message saying that joist is not load over load and seismic cannot be seperated
                            if (joist.IsLoadOverLoad == true)
                            {



                                // Move seismic loads to seismic load case

                                foreach (Load load in joist.Loads)
                                {
                                    if (load.LoadInfoCategory.Text == "SM" && (load.CaseNumber.Value == null || load.CaseNumber.Value == 1))
                                    {
                                        load.CaseNumber.Value = 3;
                                    }
                                }

                                // Copy all other positive loads from LC1 to LC3. 
                                //ISSUES: no important loads can be in any other load case than LC1. 
                                List<Load> newLoads = new List<Load>();
                                Load copiedLoad = new Load();
                                foreach (Load load in joist.Loads)
                                {
                                    if ((load.CaseNumber.Value == 1 || load.CaseNumber.Value == null)
                                        && load.Load1Value.Value >= 0
                                        && load.LoadInfoCategory.Text != "WL"
                                        && load.LoadInfoCategory.Text != "IP")
                                    {

                                        copiedLoad = DeepClone(load);
                                        copiedLoad.CaseNumber.Value = 3;
                                        newLoads.Add(copiedLoad);
                                    }
                                }
                                joist.Loads.AddRange(newLoads);

                                //ADD JOIST U DL
                                Load uDL = new Load();
                                uDL.LoadInfoType = new StringWithUpdateCheck { Text = "U" };
                                uDL.LoadInfoCategory = new StringWithUpdateCheck { Text = "CL" };
                                uDL.LoadInfoPosition = new StringWithUpdateCheck { Text = "TC" };
                                uDL.Load1Value = new DoubleWithUpdateCheck { Value = joist.UDL };
                                uDL.Load1DistanceFt = new StringWithUpdateCheck { Text = null };
                                uDL.Load1DistanceIn = new DoubleWithUpdateCheck { Value = null };
                                uDL.Load2Value = new DoubleWithUpdateCheck { Value = null };
                                uDL.Load2DistanceFt = new StringWithUpdateCheck { Text = null };
                                uDL.Load2DistanceIn = new DoubleWithUpdateCheck { Value = null };
                                uDL.LoadNote = new StringWithUpdateCheck { Text = null };
                                uDL.CaseNumber = new DoubleWithUpdateCheck { Value = 3 };
                                joist.Loads.Add(uDL);

                                //ADD JOIST U SM 
                                Load uSM = new Load();
                                uSM.LoadInfoType = new StringWithUpdateCheck { Text = "U" };
                                uSM.LoadInfoCategory = new StringWithUpdateCheck { Text = "SM" };
                                uSM.LoadInfoPosition = new StringWithUpdateCheck { Text = "TC" };
                                if (joist.UDL == null || sequence.SDS == null)
                                {
                                    uSM.Load1Value = new DoubleWithUpdateCheck { Value = null };
                                }
                                else
                                {
                                    if (joist.IsGirder == false)
                                    {
                                        uSM.Load1Value = new DoubleWithUpdateCheck { Value = Math.Ceiling((float)(0.14 * sequence.SDS * joist.UDL)) };
                                    }
                                    else
                                    {
                                        uSM.Load1Value = new DoubleWithUpdateCheck { Value = 5 * (int)Math.Ceiling((float)((0.14 * sequence.SDS * joist.UDL)/5.0)) };
                                    }
                                }
                                uSM.Load1DistanceFt = new StringWithUpdateCheck { Text = null };
                                uSM.Load1DistanceIn = new DoubleWithUpdateCheck { Value = null };
                                uSM.Load2Value = new DoubleWithUpdateCheck { Value = null };
                                uSM.Load2DistanceFt = new StringWithUpdateCheck { Text = null };
                                uSM.Load2DistanceIn = new DoubleWithUpdateCheck { Value = null };
                                uSM.LoadNote = new StringWithUpdateCheck { Text = null };
                                uSM.CaseNumber = new DoubleWithUpdateCheck { Value = 3 };
                                joist.Loads.Add(uSM);

                            }
                            else
                            {
                                string message = string.Format("MARK {0} IS NOT GIVEN IN TL/LL FORMAT; SEISMIC LC WILL NOT BE SEPERTATED", joist.Mark.Text);
                                MessageBox.Show(message);
                            }
                        }

                    }
                }
            }
        }

        public static T DeepClone<T>(T obj)
        {
            using (var ms = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(ms, obj);
                ms.Position = 0;

                return (T)formatter.Deserialize(ms);
            }
        }
        private void AddBaseType(Joist joist, BaseType bT1)
        {

            if (joist.Description.Text != null && bT1.Description.Text != null) { joist.AddError("Base Type description interferes with original; using original "); }
            if (joist.Description.Text == null && bT1.Description.Text != null) { joist.Description = bT1.Description; }
            if (joist.BaseLengthFt.Value != null && bT1.BaseLengthFt.Value != null) { joist.AddError("Base Type base length ft. interferes with original; using original "); }
            if (joist.BaseLengthFt.Value == null && bT1.BaseLengthFt.Value != null) { joist.BaseLengthFt = bT1.BaseLengthFt; }
            if (joist.BaseLengthIn.Value != null && bT1.BaseLengthIn.Value != null) { joist.AddError("Base Type base length in. interferes with original; using original "); }
            if (joist.BaseLengthIn.Value == null && bT1.BaseLengthIn.Value != null) { joist.BaseLengthIn = bT1.BaseLengthIn; }
            if (joist.TcxlQuantity.Value != null && bT1.TcxlQuantity.Value != null) { joist.AddError("Base Type TCXL quantity interferes with original; using original "); }
            if (joist.TcxlQuantity.Value == null && bT1.TcxlQuantity.Value != null) { joist.TcxlQuantity = bT1.TcxlQuantity; }
            if (joist.TcxlLengthFt.Value != null && bT1.TcxlLengthFt.Value != null) { joist.AddError("Base Type TCXL length ft. interferes with original; using original "); }
            if (joist.TcxlLengthFt.Value == null && bT1.TcxlLengthFt.Value != null) { joist.TcxlLengthFt = bT1.TcxlLengthFt; }
            if (joist.TcxlLengthIn.Value != null && bT1.TcxlLengthIn.Value != null) { joist.AddError("Base Type TCXL length in. interferes with original; using original "); }
            if (joist.TcxlLengthIn.Value == null && bT1.TcxlLengthIn.Value != null) { joist.TcxlLengthIn = bT1.TcxlLengthIn; }
            if (joist.TcxrQuantity.Value != null && bT1.TcxrQuantity.Value != null) { joist.AddError("Base Type TCXR quantity interferes with original; using original "); }
            if (joist.TcxrQuantity.Value == null && bT1.TcxrQuantity.Value != null) { joist.TcxrQuantity = bT1.TcxrQuantity; }
            if (joist.TcxrLengthFt.Value != null && bT1.TcxrLengthFt.Value != null) { joist.AddError("Base Type TCXR length ft. interferes with original; using original "); }
            if (joist.TcxrLengthFt.Value == null && bT1.TcxrLengthFt.Value != null) { joist.TcxrLengthFt = bT1.TcxrLengthFt; }
            if (joist.TcxrLengthIn.Value != null && bT1.TcxrLengthIn.Value != null) { joist.AddError("Base Type TCXR length in. interferes with original; using original "); }
            if (joist.TcxrLengthIn.Value == null && bT1.TcxrLengthIn.Value != null) { joist.TcxrLengthIn = bT1.TcxrLengthIn; }
            if (joist.SeatDepthLE.Value != null && bT1.SeatDepthLE.Value != null) { joist.AddError("Base Type LE seat depth interferes with original; using original "); }
            if (joist.SeatDepthLE.Value == null && bT1.SeatDepthLE.Value != null) { joist.SeatDepthLE = bT1.SeatDepthLE; }
            if (joist.SeatDepthRE.Value != null && bT1.SeatDepthRE.Value != null) { joist.AddError("Base Type RE seat depth interferes with original; using original "); }
            if (joist.SeatDepthRE.Value == null && bT1.SeatDepthRE.Value != null) { joist.SeatDepthRE = bT1.SeatDepthRE; }
            if (joist.BcxQuantity.Value != null && bT1.BcxQuantity.Value != null) { joist.AddError("Base Type BCX quantity interferes with original; using original "); }
            if (joist.BcxQuantity.Value == null && bT1.BcxQuantity.Value != null) { joist.BcxQuantity = bT1.BcxQuantity; }
            if (joist.Uplift.Value != null && bT1.Uplift.Value != null) { joist.AddError("Base Type uplift interferes with original; using original "); }
            if (joist.Uplift.Value == null && bT1.Uplift.Value != null) { joist.Uplift = bT1.Uplift; }
            if (joist.Erfos.Text != null && bT1.Erfos.Text != null) { joist.AddError("Base Type erfos interferes with original; using original "); }
            if (joist.Erfos.Text == null && bT1.Erfos.Text != null) { joist.Erfos = bT1.Erfos; }
            if (joist.DeflectionTL.Value != null && bT1.DeflectionTL.Value != null) { joist.AddError("Base Type TL deflection interferes with original; using original "); }
            if (joist.DeflectionTL.Value == null && bT1.DeflectionTL.Value != null) { joist.DeflectionTL = bT1.DeflectionTL; }
            if (joist.DeflectionLL.Value != null && bT1.DeflectionLL.Value != null) { joist.AddError("Base Type LL deflection interferes with original; using original "); }
            if (joist.DeflectionLL.Value == null && bT1.DeflectionLL.Value != null) { joist.DeflectionLL = bT1.DeflectionLL; }
            if (joist.WnSpacing.Text != null && bT1.WnSpacing.Text != null) { joist.AddError("Base Type WN spacing interferes with original; using original "); }
            if (joist.WnSpacing.Text == null && bT1.WnSpacing.Text != null) { joist.WnSpacing = bT1.WnSpacing; }


            //ADD THE LOADS
            foreach (Load load in bT1.Loads)
            {
                Load coppiedLoad = DeepClone(load);
                joist.Loads.Add(coppiedLoad);
            }


            //ADD THE NOTES
            foreach (StringWithUpdateCheck note in bT1.Notes)
            {
                joist.Notes.Add(note);
            }
        }

        public void AddQuantitiesFromBB(Takeoff bbTakeoff)
        {
            var blueBeamJoists =
                from s in bbTakeoff.Sequences
                from j in s.Joists
                select j;


            Regex rg = new Regex(@"\d+");

            var bbJoistTups =
                blueBeamJoists
                .GroupBy(x => x.Mark.Text)
                .Select(g => new Tuple<string, int?>(rg.Match(g.Key).Value, g.Sum(x => x.Quantity.Value)));


            Excel.Worksheet marksSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["Marks"];
            int lastUsedRow = marksSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            object[,] array = marksSheet.get_Range("A6", "C" + lastUsedRow).Value2;

            for (int i = 1; i <= array.GetLength(0); i++)
            {
                object marksColValue = array[i, 1];
                if (marksColValue != null)
                {
                    string mark = (string)marksColValue;
                    var bbMatchedJoists = bbJoistTups.Where(joist => joist.Item1 == mark);
                    if (bbMatchedJoists.Any())
                    {
                        var bbJoist = bbMatchedJoists.First();
                        var bbQty = bbJoist.Item2;
                        array[i, 3] = bbQty;
                    }
                    else
                    {
                        MessageBox.Show(String.Format("Takeoff Mark {0} is not in the BlueBeam markups.\r\n\r\n", mark));
                    }
                }
            }

            marksSheet.get_Range("A6", "C" + lastUsedRow).Value2 = array;


        }
    }

}
