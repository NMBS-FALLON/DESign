using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DESign_Sales_Excel_Add_In.Properties;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DESign_Sales_Excel_Add_In_2.Worksheet_Values
{
    public class Takeoff
    {
        private readonly _Workbook oWB = Globals.ThisAddIn.Application.ActiveWorkbook;


        // Initialize the necessary Excel objects:
        private readonly Application oXL = Globals.ThisAddIn.Application;
        private Workbook workbook;
        public double? SDS { get; set; }

        public List<Bridging> Bridging { get; set; }

        public List<BaseType> BaseTypes { get; set; }

        public List<Sequence> Sequences { get; set; }

        public Dictionary<string, (double Mf, double I)> AdditionalTakeoffInfo
        {
            get
            {
                var mfLoads =
                    this.Sequences
                    .SelectMany(s => s.Joists)
                    .Where(j => j.Notes.Where(n => n.Text.Contains("Mf =")).Count() != 0)
                    .Select(j => (Mark: j.Mark.Text,
                                  Mf: double.Parse(
                                         j.Notes
                                         .Where(n => n.Text.Contains("Mf ="))
                                         .First()
                                         .Text
                                         .Replace("Mf = ", "")
                                         .Replace("<lb-ft>", "")
                                         )));

                var inertiaNotes =
                    this.Sequences
                    .SelectMany(s => s.Joists)
                    .Where(j => j.Notes.Where(n => Regex.IsMatch(n.Text, @"I *= *(\d+\.?\d*)")).Count() != 0)
                    .Select(j => (Mark: j.Mark.Text,
                                  I: double.Parse(
                                      Regex.Match(
                                         j.Notes
                                         .Where(n => Regex.IsMatch(n.Text, @"I *= *(\d+\.?\d*)"))
                                         .First()
                                         .Text,
                                         @"I *= *(\d+\.?\d*)")
                                         .Groups[1]
                                         .Value
                                         )));

                var dict = new Dictionary<string, (double Mf, double I)>();

                if (mfLoads.Count() != 0 || inertiaNotes.Count() != 0)

                {


                    foreach (var mfLoad in mfLoads)
                    {
                        dict.Add(mfLoad.Mark, (mfLoad.Mf, 0.0));
                    }
                    foreach (var iNote in inertiaNotes)
                    {
                        var mark = iNote.Mark;
                        if (dict.ContainsKey(mark))
                        {
                            dict[mark] = (dict[mark].Mf, iNote.I);
                        }
                        else
                        {
                            dict.Add(mark, (0.0, iNote.I));
                        }
                    }
                }

                return dict;
            }
        }

        public Takeoff ImportTakeoff()
        {
            //
            var marksWS = (_Worksheet)oWB.Worksheets["Marks"];
            var baseTypesWS = (_Worksheet)oWB.Worksheets["Base Types"];
            var cover = (_Worksheet)oWB.Worksheets["Cover"];

            double? sds = null;
            if (cover.Range["K12"].Value != null && cover.Range["K12"].Value.Contains("SDS")) sds = cover.Range["M12"].Value;

            var bridgingSheetExists = false;
            foreach (Worksheet sheet in oWB.Sheets)
                if (sheet.Name.Equals("Bridging"))
                    bridgingSheetExists = true;

            if (bridgingSheetExists == false)
            {
                oWB.Worksheets.Add(After: oWB.Worksheets[baseTypesWS.Index]);
                oWB.Worksheets[baseTypesWS.Index + 1].Name = "Bridging";
                var bridgeWS = (_Worksheet)oWB.Worksheets["Bridging"];
                bridgeWS.Cells[1, 1] = "Temp";
                bridgeWS.Cells[10, 8] = "Temp";
            }

            var bridgingWS = (_Worksheet)oWB.Worksheets["Bridging"];

            var zeroOrMoreSpaces = new Regex(@"^ *$");
            Func<object, bool> cellIsBlank = s => s == null || zeroOrMoreSpaces.IsMatch(s.ToString());

            ///// GET BRIDGING ////

            var bridging = new List<Bridging>();
            var bridgingRange = bridgingWS.UsedRange;

            var bridgingCells = (object[,])bridgingRange.Value2;


            var bridgingSequence = "";
            var size = "";
            var type = "";
            var rows = 0.0;
            var length = 0.0;

            var startRow = 5;
            var lastRow = bridgingCells.GetLength(0);
            for (var row = startRow; row <= lastRow; row++)
            {
                if (!cellIsBlank(bridgingCells[row, 2])) bridgingSequence = bridgingCells[row, 2].ToString();

                if (!cellIsBlank(bridgingCells[row, 3])) size = bridgingCells[row, 3].ToString();

                if (!cellIsBlank(bridgingCells[row, 4])) type = bridgingCells[row, 4].ToString();

                rows = cellIsBlank(bridgingCells[row, 5]) ? 0.0 : Convert.ToDouble(bridgingCells[row, 5]);

                length = cellIsBlank(bridgingCells[row, 6]) ? 0.0 : Convert.ToDouble(bridgingCells[row, 6]);


                var br = new Bridging();
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
                   br.HorX
               }
                into brcs
               select new Bridging
               {
                   Sequence = brcs.Key.Sequence,
                   Size = brcs.Key.Size,
                   HorX = brcs.Key.HorX,
                   PlanFeet = brcs.Sum(br => br.PlanFeet)
               }).ToList();

            // Create a range for the 'BaseLine' tab 
            var baseTypesRange = baseTypesWS.UsedRange;

            // Add 'Base Types' column if it is not there
            if (baseTypesRange.Range["B2"].Value.Contains("DESCRIPTION") == true)
            {
                baseTypesRange.Range["B1"].EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight,
                  XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                baseTypesRange.Range["B2"].Value = "BASE TYPES";
            }

            //Create an object array containing all information from the 'Base Types' tab, in the form of a multidimensional array [row, column]
            var baseTypesCells = (object[,])baseTypesRange.Value2;

            //CHANGE ALL CELLS WITH "" TO NULL
            for (var row = 1; row <= baseTypesCells.GetLength(0); row++)
                for (var col = 1; col <= baseTypesCells.GetLength(1); col++)
                    if (baseTypesCells[row, col] is string)
                        if (cellIsBlank(baseTypesCells[row, col]))
                            baseTypesCells[row, col] = null;

            //Create a multidemnsional bool array that is true if the cell is highlighted (i.e. estimator marked it as updated) and false if it is not highlighted (i.e. cell has not been updated).
            var numRows = baseTypesRange.Rows.Count;
            var numColumns = baseTypesRange.Columns.Count;

            var isUpdated = new bool[numRows, numColumns];
            for (var row = 1; row <= numRows; row++)
                for (var col = 1; col <= numColumns; col++)
                    if (baseTypesRange[row, col].Interior.ColorIndex != -4142)
                        isUpdated[row - 1, col - 1] = true;
                    else
                        isUpdated[row - 1, col - 1] = false;

            ///////////////////

            // Determine the row of the first baseType since estimators dont always place the first baseType at the top
            var firstBaseTypeReached = false;
            var firstBaseTypeRow = 4;

            var i = 4;
            while (firstBaseTypeReached == false && i < baseTypesCells.GetLength(0))
            {
                if (baseTypesCells != null)
                    if (baseTypesCells[i, 1] != null)
                    {
                        firstBaseTypeReached = true;
                        firstBaseTypeRow = i;
                    }

                i++;
            }

            // Create a list containing the number of rows between each baseType
            var rowsPerBaseTypeList = new List<int>();
            var rowsPerBaseType = 1;
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

                if (i == baseTypesCells.GetLength(0) - 1) rowsPerBaseTypeList.Add(rowsPerBaseType);
            }


            // Now that we can break out the chunks of information for each baseType, we can create the list of baseTypeLines


            var baseTypes = new List<BaseType>();

            var rowCount = firstBaseTypeRow;
            if (rowCount != 0)
            {
                var errorMessageShown2 = false;
                foreach (var rowsForThisBaseType in rowsPerBaseTypeList)
                {
                    var baseType = new BaseType();
                    try
                    {
                        baseType.Name = new StringWithUpdateCheck
                        { Text = (string)baseTypesCells[rowCount, 1], IsUpdated = isUpdated[rowCount - 1, 0] };
                        baseType.Description = new StringWithUpdateCheck
                        { Text = (string)baseTypesCells[rowCount, 3], IsUpdated = isUpdated[rowCount - 1, 2] };
                        baseType.BaseLengthFt = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 4], IsUpdated = isUpdated[rowCount - 1, 3] };
                        baseType.BaseLengthIn = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 5], IsUpdated = isUpdated[rowCount - 1, 4] };
                        baseType.TcxlQuantity = new IntWithUpdateCheck
                        { Value = (int?)(double?)baseTypesCells[rowCount, 6], IsUpdated = isUpdated[rowCount - 1, 5] };
                        baseType.TcxlLengthFt = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 7], IsUpdated = isUpdated[rowCount - 1, 6] };
                        baseType.TcxlLengthIn = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 8], IsUpdated = isUpdated[rowCount - 1, 7] };
                        baseType.TcxrQuantity = new IntWithUpdateCheck
                        { Value = (int?)(double?)baseTypesCells[rowCount, 9], IsUpdated = isUpdated[rowCount - 1, 8] };
                        baseType.TcxrLengthFt = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 10], IsUpdated = isUpdated[rowCount - 1, 9] };
                        baseType.TcxrLengthIn = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 11], IsUpdated = isUpdated[rowCount - 1, 10] };
                        baseType.SeatDepthLE = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 12], IsUpdated = isUpdated[rowCount - 1, 11] };
                        baseType.SeatDepthRE = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 13], IsUpdated = isUpdated[rowCount - 1, 12] };
                        if (Convert.ToString(baseTypesCells[rowCount, 14]).ToUpper().Contains("BE"))
                            baseType.BcxQuantity = new IntWithUpdateCheck { Value = -1, IsUpdated = isUpdated[rowCount - 1, 13] };
                        else if (Convert.ToString(baseTypesCells[rowCount, 14]).ToUpper().Contains("1E"))
                            baseType.BcxQuantity = new IntWithUpdateCheck { Value = -2, IsUpdated = isUpdated[rowCount - 1, 13] };
                        else
                            baseType.BcxQuantity = new IntWithUpdateCheck
                            { Value = (int?)(double?)baseTypesCells[rowCount, 14], IsUpdated = isUpdated[rowCount - 1, 13] };
                        baseType.Uplift = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 15], IsUpdated = isUpdated[rowCount - 1, 14] };
                        baseType.Erfos = new StringWithUpdateCheck
                        { Text = (string)baseTypesCells[rowCount, 27], IsUpdated = isUpdated[rowCount - 1, 26] };
                        baseType.DeflectionTL = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 28], IsUpdated = isUpdated[rowCount - 1, 27] };
                        baseType.DeflectionLL = new DoubleWithUpdateCheck
                        { Value = (double?)baseTypesCells[rowCount, 29], IsUpdated = isUpdated[rowCount - 1, 28] };
                        baseType.WnSpacing = new StringWithUpdateCheck
                        { Text = (string)baseTypesCells[rowCount, 30], IsUpdated = isUpdated[rowCount - 1, 29] };


                        var loads = new List<Load>();
                        var notes = new List<StringWithUpdateCheck>();
                        var baseTypeStrings = new List<string>();

                        for (i = 0; i < rowsForThisBaseType; i++)
                        {
                            if (baseTypesCells[rowCount + i, 2] != null && (string)baseTypesCells[rowCount + i, 2] != "")
                                baseTypeStrings.Add(baseTypesCells[rowCount + i, 2].ToString());

                            var load = new Load();
                            load.LoadInfoType = new StringWithUpdateCheck
                            { Text = (string)baseTypesCells[rowCount + i, 16], IsUpdated = isUpdated[rowCount + i - 1, 15] };
                            load.LoadInfoCategory = new StringWithUpdateCheck
                            { Text = (string)baseTypesCells[rowCount + i, 17], IsUpdated = isUpdated[rowCount + i - 1, 16] };
                            load.LoadInfoPosition = new StringWithUpdateCheck
                            { Text = (string)baseTypesCells[rowCount + i, 18], IsUpdated = isUpdated[rowCount + i - 1, 17] };
                            load.Load1Value = new DoubleWithUpdateCheck
                            { Value = (double?)baseTypesCells[rowCount + i, 19], IsUpdated = isUpdated[rowCount + i - 1, 18] };
                            if (baseTypesCells[rowCount + i, 20] is double)
                                load.Load1DistanceFt = new StringWithUpdateCheck
                                {
                                    Text = Convert.ToString((double?)baseTypesCells[rowCount + i, 20]),
                                    IsUpdated = isUpdated[rowCount + i - 1, 19]
                                };
                            else
                                load.Load1DistanceFt = new StringWithUpdateCheck
                                { Text = (string)baseTypesCells[rowCount + i, 20], IsUpdated = isUpdated[rowCount + i - 1, 19] };

                            load.Load1DistanceIn = new DoubleWithUpdateCheck
                            { Value = (double?)baseTypesCells[rowCount + i, 21], IsUpdated = isUpdated[rowCount + i - 1, 20] };
                            load.Load2Value = new DoubleWithUpdateCheck
                            { Value = (double?)baseTypesCells[rowCount + i, 22], IsUpdated = isUpdated[rowCount + i - 1, 21] };
                            if (baseTypesCells[rowCount + i, 23] is double)
                                load.Load2DistanceFt = new StringWithUpdateCheck
                                {
                                    Text = Convert.ToString((double?)baseTypesCells[rowCount + i, 23]),
                                    IsUpdated = isUpdated[rowCount + i - 1, 22]
                                };
                            else
                                load.Load2DistanceFt = new StringWithUpdateCheck
                                { Text = (string)baseTypesCells[rowCount + i, 23], IsUpdated = isUpdated[rowCount + i - 1, 22] };
                            load.Load2DistanceIn = new DoubleWithUpdateCheck
                            { Value = (double?)baseTypesCells[rowCount + i, 24], IsUpdated = isUpdated[rowCount + i - 1, 23] };
                            load.CaseNumber = new DoubleWithUpdateCheck
                            {
                                Value = ToNullableDouble((string)baseTypesCells[rowCount + i, 25]),
                                IsUpdated = isUpdated[rowCount + i - 1, 24]
                            };
                            load.Reference = new StringWithUpdateCheck
                            { Text = (string)baseTypesCells[rowCount + i, 26], IsUpdated = isUpdated[rowCount + i - 1, 25] };
                            if (load.IsNull == false) loads.Add(load);

                            var note = new StringWithUpdateCheck
                            { Text = (string)baseTypesCells[rowCount + i, 31], IsUpdated = isUpdated[rowCount + i - 1, 30] };
                            if (note.Text != null) notes.Add(note);
                            if (note.Text == null && note.IsUpdated) notes.Add(note);
                        }

                        baseType.Loads = loads;
                        baseType.Notes = notes;
                        baseType.BaseTypeStrings = baseTypeStrings;

                        baseTypes.Add(baseType);
                        rowCount = rowCount + rowsForThisBaseType;
                    }
                    catch
                    {
                        if (errorMessageShown2 == false)
                        {
                            MessageBox.Show(string.Format(@"BASETYPE {0}:
    ISSUE PULLING INFO FROM BASE TYPES TAB.
    PLEASE CHECK THAT COLUMNS ARE FILLED IN CORRECTLY.
    THIS MUST BE FIXED BEFORE CONVERTING THE TAKEOFF.", baseType.Name.Text));
                            errorMessageShown2 = true;
                        }
                    }
                }
            }

            /// Add Recursive BaseTypes
            /// 

            foreach (var baseType in baseTypes)
            {
                var allBaseTypeStrings = baseType.BaseTypeStrings;
                var addedAllBaseTypes = false;
                // Select the matching base type. THIS WILL NEED TO BE UPDATED TO CHECK FOR TYPOS AND TO MAKE SURE BASETYPES EXIST
                while (addedAllBaseTypes == false)
                {
                    var originalNumBaseTypeStrings = DeepClone(baseType.BaseTypeStrings.Count);

                    var matchedBaseType = from bT in baseTypes
                                          where baseType.BaseTypeStrings.Contains(bT.Name.Text)
                                          select bT;
                    foreach (var bT in matchedBaseType)
                        foreach (var baseTypeString in bT.BaseTypeStrings)
                            if (baseTypeString != baseType.Name.Text && baseType.BaseTypeStrings.Contains(baseTypeString) == false)
                                baseType.BaseTypeStrings.Add(baseTypeString);

                    addedAllBaseTypes = originalNumBaseTypeStrings == baseType.BaseTypeStrings.Count ? true : false;
                }

                var _matchedBaseType = from bT in baseTypes
                                       where baseType.BaseTypeStrings.Contains(bT.Name.Text)
                                       select bT;

                foreach (var bT in _matchedBaseType)
                {
                    //ADD VALUES    ???DO I NEED TO CHECK ANYTHING THAT MAY BE UPDATED??? IF SO HOW TO IMPLEMENT?
                    if (baseType.Description.Text == null || baseType.Description.Text == "")
                        baseType.Description = bT.Description;
                    if (baseType.BaseLengthFt.Value == null) baseType.BaseLengthFt = bT.BaseLengthFt;
                    if (baseType.BaseLengthIn.Value == null) baseType.BaseLengthIn = bT.BaseLengthIn;
                    if (baseType.TcxlQuantity.Value == null) baseType.TcxlQuantity = bT.TcxlQuantity;
                    if (baseType.TcxlLengthFt.Value == null) baseType.TcxlLengthFt = bT.TcxlLengthFt;
                    if (baseType.TcxlLengthIn.Value == null) baseType.TcxlLengthIn = bT.TcxlLengthIn;
                    if (baseType.SeatDepthLE.Value == null) baseType.SeatDepthLE = bT.SeatDepthLE;
                    if (baseType.SeatDepthRE.Value == null) baseType.SeatDepthRE = bT.SeatDepthRE;
                    if (baseType.BcxQuantity.Value == null) baseType.BcxQuantity = bT.BcxQuantity;
                    if (baseType.Uplift.Value == null) baseType.Uplift = bT.Uplift;
                    baseType.Loads.AddRange(bT.Loads);
                    baseType.Notes.AddRange(bT.Notes);
                }
            }

            ///////////////////

            // Create a range for the 'Marks' tab
            var marksRange = marksWS.UsedRange;

            // Create an object array containing all information from the 'Marks' tab, in the form of a multidimensional array [rows, column]
            var marksCells = (object[,])marksRange.Value2;

            //CHANGE ALL CELLS WITH "" TO NULL
            for (var row = 1; row <= marksCells.GetLength(0); row++)
                for (var col = 1; col <= marksCells.GetLength(1); col++)
                    if (marksCells[row, col] is string)
                        if (cellIsBlank(marksCells[row, col]))
                            marksCells[row, col] = null;

            //Create a multidemnsional bool array that is true if the cell is highlighted (i.e. estimator marked it as updated) and false if it is not highlighted (i.e. cell has not been updated).
            numRows = marksRange.Rows.Count;
            numColumns = Math.Min(marksRange.Columns.Count, 100);

            isUpdated = new bool[numRows, numColumns];
            for (var row = 1; row <= numRows; row++)
                for (var col = 1; col <= numColumns; col++)
                    if (marksRange[row, col].Interior.ColorIndex != -4142)
                        isUpdated[row - 1, col - 1] = true;
                    else
                        isUpdated[row - 1, col - 1] = false;


            // Determine the row of the first mark or sequence since estimators dont always place it at the top
            var firstLineReached = false;
            var firstMarkRow = 4;

            i = 4;
            while (firstLineReached == false && i < marksCells.GetLength(0))
            {
                if (marksCells != null)
                    if (marksCells[i, 1] != null)
                    {
                        firstLineReached = true;
                        firstMarkRow = i;
                    }

                i++;
            }


            // Create a list containing the number of rows between each mark
            var rowsPerMarkList = new List<int>();
            var rowsPerMark = 1;
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

                if (i == marksCells.GetLength(0) - 1) rowsPerMarkList.Add(rowsPerMark);
            }

            // Now that we can break out the chunks of information for each mark, we can create the list of joistLines
            var joistLines = new List<Joist>();

            rowCount = firstMarkRow;
            var errorMessageShown = false;
            foreach (var rowsForThisMark in rowsPerMarkList)
            {
                var joist = new Joist();
                try
                {
                    joist.Mark = new StringWithUpdateCheck
                    { Text = (string)marksCells[rowCount, 1], IsUpdated = isUpdated[rowCount - 1, 0] };
                    joist.Quantity = new IntWithUpdateCheck
                    { Value = (int?)(double?)marksCells[rowCount, 3], IsUpdated = isUpdated[rowCount - 1, 2] };
                    joist.Description = new StringWithUpdateCheck
                    { Text = (string)marksCells[rowCount, 4], IsUpdated = isUpdated[rowCount - 1, 3] };
                    joist.BaseLengthFt = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 5], IsUpdated = isUpdated[rowCount - 1, 4] };
                    joist.BaseLengthIn = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 6], IsUpdated = isUpdated[rowCount - 1, 5] };
                    joist.TcxlQuantity = new IntWithUpdateCheck
                    { Value = (int?)(double?)marksCells[rowCount, 7], IsUpdated = isUpdated[rowCount - 1, 6] };
                    joist.TcxlLengthFt = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 8], IsUpdated = isUpdated[rowCount - 1, 7] };
                    joist.TcxlLengthIn = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 9], IsUpdated = isUpdated[rowCount - 1, 8] };
                    joist.TcxrQuantity = new IntWithUpdateCheck
                    { Value = (int?)(double?)marksCells[rowCount, 10], IsUpdated = isUpdated[rowCount - 1, 9] };
                    joist.TcxrLengthFt = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 11], IsUpdated = isUpdated[rowCount - 1, 10] };
                    joist.TcxrLengthIn = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 12], IsUpdated = isUpdated[rowCount - 1, 11] };
                    joist.SeatDepthLE = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 13], IsUpdated = isUpdated[rowCount - 1, 12] };
                    joist.SeatDepthRE = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 14], IsUpdated = isUpdated[rowCount - 1, 13] };
                    if (Convert.ToString(marksCells[rowCount, 15]).ToUpper().Contains("BE"))
                        joist.BcxQuantity = new IntWithUpdateCheck
                        { Value = joist.Quantity.Value * 2, IsUpdated = isUpdated[rowCount - 1, 14] };
                    else if (Convert.ToString(marksCells[rowCount, 15]).ToUpper().Contains("1E"))
                        joist.BcxQuantity = new IntWithUpdateCheck
                        { Value = joist.Quantity.Value, IsUpdated = isUpdated[rowCount - 1, 14] };
                    else
                        joist.BcxQuantity = new IntWithUpdateCheck
                        { Value = (int?)(double?)marksCells[rowCount, 15], IsUpdated = isUpdated[rowCount - 1, 14] };
                    joist.Uplift = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 16], IsUpdated = isUpdated[rowCount - 1, 15] };
                    joist.Erfos = new StringWithUpdateCheck
                    { Text = (string)marksCells[rowCount, 28], IsUpdated = isUpdated[rowCount - 1, 27] };
                    joist.DeflectionTL = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 29], IsUpdated = isUpdated[rowCount - 1, 28] };
                    joist.DeflectionLL = new DoubleWithUpdateCheck
                    { Value = (double?)marksCells[rowCount, 30], IsUpdated = isUpdated[rowCount - 1, 29] };
                    joist.WnSpacing = new StringWithUpdateCheck
                    { Text = (string)marksCells[rowCount, 31], IsUpdated = isUpdated[rowCount - 1, 30] };

                    var baseTypesOnMark = new List<StringWithUpdateCheck>();
                    var loads = new List<Load>();
                    var notes = new List<StringWithUpdateCheck>();

                    for (i = 0; i < rowsForThisMark; i++)
                    {
                        var baseTypeOnMark = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 2] };
                        if (baseTypeOnMark.Text != null && baseTypeOnMark.IsUpdated == false) baseTypesOnMark.Add(baseTypeOnMark);


                        var load = new Load();
                        load.LoadInfoType = new StringWithUpdateCheck
                        { Text = (string)marksCells[rowCount + i, 17], IsUpdated = isUpdated[rowCount + i - 1, 16] };
                        load.LoadInfoCategory = new StringWithUpdateCheck
                        { Text = (string)marksCells[rowCount + i, 18], IsUpdated = isUpdated[rowCount + i - 1, 17] };
                        load.LoadInfoPosition = new StringWithUpdateCheck
                        { Text = (string)marksCells[rowCount + i, 19], IsUpdated = isUpdated[rowCount + i - 1, 18] };
                        load.Load1Value = new DoubleWithUpdateCheck
                        { Value = (double?)marksCells[rowCount + i, 20], IsUpdated = isUpdated[rowCount + i - 1, 19] };
                        if (marksCells[rowCount + i, 21] is double)
                            load.Load1DistanceFt = new StringWithUpdateCheck
                            {
                                Text = Convert.ToString((double?)marksCells[rowCount + i, 21]),
                                IsUpdated = isUpdated[rowCount + i - 1, 20]
                            };
                        else
                            load.Load1DistanceFt = new StringWithUpdateCheck
                            { Text = (string)marksCells[rowCount + i, 21], IsUpdated = isUpdated[rowCount + i - 1, 20] };
                        load.Load1DistanceIn = new DoubleWithUpdateCheck
                        { Value = (double?)marksCells[rowCount + i, 22], IsUpdated = isUpdated[rowCount + i - 1, 21] };
                        load.Load2Value = new DoubleWithUpdateCheck
                        { Value = (double?)marksCells[rowCount + i, 23], IsUpdated = isUpdated[rowCount + i - 1, 22] };
                        if (marksCells[rowCount + i, 21] is double)
                            load.Load2DistanceFt = new StringWithUpdateCheck
                            {
                                Text = Convert.ToString((double?)marksCells[rowCount + i, 24]),
                                IsUpdated = isUpdated[rowCount + i - 1, 23]
                            };
                        else
                            load.Load2DistanceFt = new StringWithUpdateCheck
                            { Text = (string)marksCells[rowCount + i, 24], IsUpdated = isUpdated[rowCount + i - 1, 23] };
                        load.Load2DistanceIn = new DoubleWithUpdateCheck
                        { Value = (double?)marksCells[rowCount + i, 25], IsUpdated = isUpdated[rowCount + i - 1, 24] };
                        if (marksCells[rowCount + i, 26] is string)
                            load.CaseNumber = new DoubleWithUpdateCheck
                            { Value = Convert.ToDouble(marksCells[rowCount + i, 26]), IsUpdated = isUpdated[rowCount + i - 1, 25] };
                        else
                            load.CaseNumber = new DoubleWithUpdateCheck
                            { Value = (double?)marksCells[rowCount + i, 26], IsUpdated = isUpdated[rowCount + i - 1, 25] };
                        load.Reference = new StringWithUpdateCheck
                        { Text = (string)marksCells[rowCount + i, 27], IsUpdated = isUpdated[rowCount + i - 1, 26] };
                        if (load.IsNull == false) loads.Add(load);

                        var note = new StringWithUpdateCheck
                        { Text = (string)marksCells[rowCount + i, 32], IsUpdated = isUpdated[rowCount + i - 1, 31] };
                        if (note.Text != null) notes.Add(note);
                        if (note.Text == null && note.IsUpdated) notes.Add(note);
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
                        MessageBox.Show(string.Format(@"MARK {0}:
    ISSUE PULLING INFO FROM MARKS TAB.
    PLEASE CHECK THAT COLUMNS ARE FILLED IN CORRECTLY.
    THIS MUST BE FIXED BEFORE CONVERTING THE TAKEOFF.", joist.Mark.Text));
                        errorMessageShown = true;
                    }
                }
            }

            //Seperate Sequences
            var sequences = new List<Sequence>();
            var sequenceQuery = from jst in joistLines
                                where jst.Mark.Text.Contains("{") && jst.Mark.Text.Contains("}")
                                select jst;
            if (!sequenceQuery.Any()) //No named sequences on takeoff
            {
                var sequence = new Sequence();
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
                    var sequence = new Sequence();
                    sequence.Name = new StringWithUpdateCheck { Text = "" };


                    var jstIndex = 0;

                    for (var joistIndex = jstIndex; joistIndex < joistLines.Count; joistIndex++)
                    {
                        if (joistLines[joistIndex].Quantity.Value == null && joistLines[joistIndex].Description.Text == null)
                        {
                            sequence.Joists = new List<Joist>();
                            sequence.Name.Text = joistLines[joistIndex].Mark.Text;
                            sequence.Name.IsUpdated = joistLines[joistIndex].Mark.IsUpdated;
                        }
                        else
                        {
                            var joist = new Joist();
                            joist = joistLines[joistIndex];
                            sequence.Joists.Add(joist);
                        }

                        if (joistIndex + 1 < joistLines.Count)
                        {
                            if (joistLines[joistIndex + 1].Quantity.Value == null &&
                                joistLines[joistIndex + 1].Description.Text == null &&
                                joistLines[joistIndex + 1].BaseTypesOnMark.Count == 0)
                            {
                                var coppiedSequence = new Sequence();
                                var newJoists = new List<Joist>();
                                foreach (var jst in sequence.Joists)
                                {
                                    var newJoist = new Joist();
                                    newJoist = DeepClone(jst);
                                    newJoists.Add(newJoist);
                                }

                                var coppiedName = new StringWithUpdateCheck();
                                coppiedName = DeepClone(sequence.Name);
                                coppiedSequence.Name = coppiedName;
                                coppiedSequence.Joists = newJoists;
                                sequences.Add(coppiedSequence);
                            }
                        }
                        else
                        {
                            var coppiedSequence = new Sequence();
                            var newJoists = new List<Joist>();
                            foreach (var jst in sequence.Joists)
                            {
                                var newJoist = new Joist();
                                newJoist = DeepClone(jst);
                                newJoists.Add(newJoist);
                            }

                            var coppiedName = new StringWithUpdateCheck();
                            coppiedName = DeepClone(sequence.Name);
                            coppiedSequence.Name = coppiedName;
                            coppiedSequence.Joists = newJoists;
                            sequences.Add(coppiedSequence);
                        }
                    }
                }
            }

            var takeoff = new Takeoff();
            takeoff.SDS = sds;
            takeoff.BaseTypes = baseTypes;
            takeoff.Sequences = sequences;
            foreach (var br in bridging) br.PlanFeet = Math.Ceiling(br.PlanFeet / 20.0) * 20.0;
            takeoff.Bridging = bridging;


            foreach (var seq in takeoff.Sequences)
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
                            //ADD VALUES    ???DO I NEED TO CHECK ANYTHING THAT MAY BE UPDATED??? IF SO HOW TO IMPLEMENT?
                            AddBaseType(joist, bT);
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
                                                                                                        seq.Name.Text.ToUpper()
                                                                                                          .Replace(" ", string.Empty) + "]")
                                      select bT1;

                    var allJoistSequence = from bT1 in baseTypes
                                           where bT1.Name.Text != null
                                           where bT1.Name.Text.ToUpper().Replace(" ", string.Empty).Contains("[ALLJ:" +
                                                                                                             seq.Name.Text.ToUpper()
                                                                                                               .Replace(" ", string.Empty) + "]")
                                           select bT1;

                    var allGirderSequence = from bT1 in baseTypes
                                            where bT1.Name.Text != null
                                            where bT1.Name.Text.ToUpper().Replace(" ", string.Empty).Contains("[ALLG:" +
                                                                                                              seq.Name.Text.ToUpper()
                                                                                                                .Replace(" ", string.Empty) + "]")
                                            select bT1;

                    if (all.Any())
                        foreach (var bT1 in all)
                            AddBaseType(joist, DeepClone(bT1));

                    if (allJoist.Any() && joist.IsGirder == false)
                        foreach (var bT1 in allJoist)
                            AddBaseType(joist, DeepClone(bT1));

                    if (allGirder.Any() && joist.IsGirder)
                        foreach (var bT1 in allGirder)
                            AddBaseType(joist, DeepClone(bT1));

                    if (allSequence.Any())
                        foreach (var bT1 in allSequence)
                            AddBaseType(joist, DeepClone(bT1));

                    if (allJoistSequence.Any() && joist.IsGirder == false)
                        foreach (var bT1 in allJoistSequence)
                            AddBaseType(joist, DeepClone(bT1));

                    if (allGirderSequence.Any() && joist.IsGirder)
                        foreach (var bT1 in allGirderSequence)
                            AddBaseType(joist, DeepClone(bT1));
                }

            // Checks:
            var errors = "";
            var joistWithErrorCount = 0;

            var joistMarks = from seq in sequences
                             from jst in seq.Joists
                             select jst.Mark.Text;
            // check that there are no duplicate marks
            var markGroups = joistMarks.GroupBy(x => x)
              .Where(g => g.Count() > 1);
            if (markGroups.Any())
                foreach (var group in markGroups)
                    errors += string.Format("  There are ({0}) marks labeled \"{1}\"\r\n\r\n", group.Count().ToString(),
                      group.Key);

            foreach (var bt in takeoff.BaseTypes)
                if (bt.Errors.Count != 0)
                {
                    errors += string.Format("  BASETYPE {0}:\r\n", bt.Name.Text);
                    foreach (var error in bt.Errors) errors += "      " + error + "\r\n";
                    errors += "\r\n";
                }

            var baseTypeNames = from bt in baseTypes
                                select bt.Name.Text;


            // ADJUST SPECIAL LOADS
            // ..... FUTURE .....
            // NEED TO ADD CHECKS TO MAKE SURE ALL OF THE SPECIAL LOADS ARE PROVIDIG ACCURATE INFORMATION.
            foreach (var sequence in takeoff.Sequences)
                foreach (var joist in sequence.Joists)
                {
                    var newLoads = new List<Load>();
                    foreach (var load in joist.Loads)
                    {
                        if (load.LoadInfoCategory.Text == "SMU")
                        {
                            if (load.Load1Value.Value == null)
                            {
                                MessageBox.Show("Mark " + joist.Mark.Text + ": There is a load with a missing 'Load 1 Value'");
                            }

                            load.Load1Value.Value = 1 * (int)Math.Ceiling((decimal)(load.Load1Value.Value * 0.7 / 1.0));
                            if (load.Load2Value.Value != null)
                                load.Load2Value.Value = 1 * (int)Math.Ceiling((decimal)(load.Load2Value.Value * 0.7 / 1.0));
                            load.LoadInfoCategory.Text = "SM";
                        }

                        if (load.LoadInfoCategory.Text == "WLU")
                        {
                            load.Load1Value.Value = 1 * (int)Math.Ceiling((decimal)(load.Load1Value.Value * 0.6 / 1.0));
                            if (load.Load2Value.Value != null)
                                load.Load2Value.Value = 1 * (int)Math.Ceiling((decimal)(load.Load2Value.Value * 0.6 / 1.0));
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
                                    var numPanelPoints =
                                      Convert.ToInt16(joist.Description.Text.Split(new[] { "G", "N" }, StringSplitOptions.None)[1]) - 1;
                                    for (var j = 1; j <= numPanelPoints; j++)
                                    {
                                        var ppLoad = DeepClone(load);
                                        ppLoad.LoadInfoType.Text = "C";
                                        ppLoad.Load1DistanceFt.Text = "P" + DeepClone(j);
                                        newLoads.Add(ppLoad);
                                    }
                                }
                                else
                                {
                                    var loadLocations = load.Load1DistanceFt.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                    loadLocations = loadLocations.Select(loadLocation => "P" + loadLocation.Replace(" ", "")).ToArray();
                                    foreach (var loadLocation in loadLocations)
                                    {
                                        var newLoad = DeepClone(load);
                                        newLoad.LoadInfoType.Text = "C";
                                        newLoad.Load1DistanceFt.Text = loadLocation;
                                        newLoads.Add(newLoad);
                                    }
                                }
                            }
                        }

                        if (load.LoadInfoType.Text == "CUP" || load.LoadInfoType.Text == "CUA")
                        {
                            var joistFt = joist.BaseLengthFt.Value == null ? 0.0 : (double)joist.BaseLengthFt.Value;
                            var joistIn = joist.BaseLengthIn.Value == null ? 0.0 : (double)joist.BaseLengthIn.Value;
                            var joistLengthInFt = joistFt + joistIn / 12.0;

                            var loadFt = load.Load1DistanceFt.Text == null ? 0.0 : double.Parse(load.Load1DistanceFt.Text);
                            var loadIn = load.Load1DistanceIn.Value == null ? 0.0 : (double)load.Load1DistanceIn.Value;
                            var spaceInFt = loadFt + loadIn / 12.0;

                            var ptLoad = load.Load1Value.Value == null ? 0.0 : (double)load.Load1Value.Value;
                            var uniformLoadValue = Math.Ceiling(ptLoad / spaceInFt - ptLoad / joistLengthInFt);

                            var uniformLoad = DeepClone(load);
                            uniformLoad.Load1Value.Value = uniformLoadValue;
                            uniformLoad.LoadInfoType.Text = "U";
                            uniformLoad.Load1DistanceFt.Text = null;
                            uniformLoad.Load1DistanceIn.Value = null;

                            var cpLoad = DeepClone(load);
                            cpLoad.LoadInfoType.Text = load.LoadInfoType.Text == "CUP" ? "CP" : "CA";
                            cpLoad.Load1DistanceFt.Text = null;
                            cpLoad.Load1DistanceIn.Value = null;

                            newLoads.Add(uniformLoad);
                            newLoads.Add(cpLoad);
                        }
                    }

                    joist.Loads.AddRange(newLoads);

                    string[] loadInfoTypesToRemove = { "CMP", "CUP", "CUA" };
                    joist.Loads.RemoveAll(load => loadInfoTypesToRemove.Contains(load.LoadInfoType.Text));
                }


            foreach (var seq in sequences)
                foreach (var joist in seq.Joists)
                {
                    foreach (var bt in joist.BaseTypesOnMark)
                        if (baseTypeNames.Contains(bt.Text) == false)
                            joist.AddError(string.Format("'Base Types' tab does not contain a definition for \"{0}\"", bt.Text));
                    if (joist.Errors.Count != 0)
                    {
                        errors += string.Format("  MARK {0}:\r\n", joist.Mark.Text);
                        foreach (var error in joist.Errors) errors += "      " + error + "\r\n";
                        errors += "\r\n";
                    }
                }

            // adjust load distances based on reference
            foreach (var seq in sequences)
            {
                foreach (var joist in seq.Joists)
                {
                    var baseLength =
                        joist.BaseLengthFt.Value == null ? 0.0 : (double)joist.BaseLengthFt.Value
                        + (joist.BaseLengthIn.Value == null ? 0.0 : (double)joist.BaseLengthIn.Value) / 12.0;
                    var tcxlLength =
                        joist.TcxlLengthFt.Value == null ? 0.0 : (double)joist.TcxlLengthFt.Value
                        + (joist.TcxlLengthIn.Value == null ? 0.0 : (double)joist.TcxlLengthIn.Value) / 12.0;
                    var tcxrLength =
                        joist.TcxrLengthFt.Value == null ? 0.0 : (double)joist.TcxrLengthFt.Value
                        + (joist.TcxrLengthIn.Value == null ? 0.0 : (double)joist.TcxrLengthIn.Value) / 12.0;

                   foreach (var load in joist.Loads)
                    {
                    
                        if (load.Reference.Text == "RBL")
                        {
                            if (load.Load1DistanceFt.Text != null && load.Load1DistanceFt.Text.Contains("P"))
                            {
                                var numPanelPoints = int.Parse(Regex.Match(joist.DescriptionAdjusted.Text, @"\d*.?\d*[a-zA-Z]+(\d+)N").Groups[1].Value);
                                var load1PanelPoint = int.Parse(Regex.Match(load.Load1DistanceFt.Text, @"P(\d+)").Groups[1].Value);
                                load.Load1DistanceFt.Text = "P" + (numPanelPoints - load1PanelPoint).ToString();
                            }
                            else
                            {
                                double? currentLoad1Distance = null;
                                if (load.Load1DistanceFt.Text != null && load.Load1DistanceFt.Text != "")
                                {
                                    currentLoad1Distance = 
                                        double.Parse(load.Load1DistanceFt.Text) +
                                        (load.Load1DistanceIn.Value == null ? 0.0 : (double)load.Load1DistanceIn.Value) / 12.0;
                                }

                                var newLoad1Distance =
                                    currentLoad1Distance == null ? null : baseLength - currentLoad1Distance;
                                load.Load1DistanceFt.Text = newLoad1Distance == null ? null : newLoad1Distance.ToString();
                                load.Load1DistanceIn.Value = 0.0;
                            }
                            if (load.Load2DistanceFt.Text != null && load.Load2DistanceFt.Text.Contains("P"))
                            {
                                var numPanelPoints = int.Parse(Regex.Match(joist.DescriptionAdjusted.Text, @"\d*.?\d*[a-zA-Z]+(\d+)N").Groups[1].Value);
                                var Load2PanelPoint = int.Parse(Regex.Match(load.Load2DistanceFt.Text, @"P(\d+)").Groups[1].Value);
                                load.Load2DistanceFt.Text = "P" + (numPanelPoints - Load2PanelPoint).ToString();
                            }
                            else
                            {
                                double? currentLoad2Distance = null;
                                if (load.Load2DistanceFt.Text != null && load.Load2DistanceFt.Text != "")
                                {
                                    currentLoad2Distance =
                                        double.Parse(load.Load2DistanceFt.Text) +
                                        (load.Load2DistanceIn.Value == null ? 0.0 : (double)load.Load2DistanceIn.Value) / 12.0;
                                }

                                var newLoad2Distance =
                                    currentLoad2Distance == null ? null : baseLength - currentLoad2Distance;
                                load.Load2DistanceFt.Text = newLoad2Distance == null ? null : newLoad2Distance.ToString();
                                load.Load2DistanceIn.Value = 0.0;
                            }
                        }
                    }
                }
            }

            foreach (var seq in sequences)
                foreach (var joist in seq.Joists)
                    if (joist.isComposite)
                    {
                        var moment = 0.0;
                        var loadsInCase1 = joist.Loads.Where(l => l.CaseNumber.Value == null || l.CaseNumber.Value == 1.0);
                        foreach (var load in loadsInCase1)
                        {
                            var factor = 1.0;
                            switch (load.LoadInfoCategory.Text)
                            {
                                case "LL":
                                    factor = 1.6;
                                    break;
                                case "TL":
                                    factor = 1.44;
                                    break;
                                case "CL":
                                case "DL":
                                    factor = 1.2;
                                    break;
                                default:
                                    factor = 1.0;
                                    break;
                            }

                            var firstLoadValue = Convert.ToDouble(load.Load1Value.Value);
                            firstLoadValue = firstLoadValue > 0 ? firstLoadValue : 0;
                            var firstLoadDistanceFtinFt =
                              Convert.ToDouble(load.Load1DistanceFt.Text == "" ? "0" : load.Load1DistanceFt.Text);
                            var firstLoadDistanceIninFt = Convert.ToDouble(load.Load1DistanceIn.Value) / 12.0;
                            var firstLoadDistanceinFt = firstLoadDistanceFtinFt + firstLoadDistanceIninFt;
                            var secondLoadValue = Convert.ToDouble(load.Load2Value.Value);
                            secondLoadValue = secondLoadValue > 0 ? secondLoadValue : 0;
                            var secondLoadDistanceFtinFt =
                              Convert.ToDouble(load.Load2DistanceFt.Text == "" ? "0" : load.Load2DistanceFt.Text);
                            var secondLoadDistanceIninFt = Convert.ToDouble(load.Load2DistanceIn.Value) / 12.0;
                            var secondLoadDistanceinFt = secondLoadDistanceFtinFt + secondLoadDistanceIninFt;
                            var joistLength = Convert.ToDouble(joist.BaseLengthFt.Value);

                            var firstLoadAddMomentAtLoad = 0.0;
                            var firstLoadAddMomentAtMidSpan = 0.0;
                            var secondLoadAddMomentAtLoad = 0.0;
                            var secondLoadAddMomentAtMidSpan = 0.0;

                            switch (load.LoadInfoType.Text)
                            {
                                case "U":
                                    moment = moment + firstLoadValue * Math.Pow(joistLength, 2.0) / 8.0;
                                    break;
                                case "C":
                                    firstLoadAddMomentAtLoad = factor * firstLoadValue * firstLoadDistanceinFt *
                                                               (joistLength - firstLoadDistanceinFt) / joistLength;
                                    secondLoadAddMomentAtLoad = factor * secondLoadValue * secondLoadDistanceinFt *
                                                                (joistLength - secondLoadDistanceinFt) / joistLength;
                                    var firstLoadDistanceFromMid = Math.Abs(joistLength / 2.0 - firstLoadDistanceinFt);
                                    var firstLoadMomentSlope = firstLoadAddMomentAtLoad / (joistLength / 2.0 + firstLoadDistanceFromMid);
                                    firstLoadAddMomentAtMidSpan =
                                      firstLoadAddMomentAtLoad - firstLoadMomentSlope * firstLoadDistanceFromMid;

                                    var secondLoadDistanceFromMid = Math.Abs(joistLength / 2.0 - secondLoadDistanceinFt);
                                    var secondLoadMomentSlope = secondLoadAddMomentAtLoad / (joistLength / 2.0 + secondLoadDistanceFromMid);
                                    secondLoadAddMomentAtMidSpan =
                                      secondLoadAddMomentAtLoad - secondLoadMomentSlope * secondLoadDistanceFromMid;

                                    moment = moment + firstLoadAddMomentAtMidSpan + secondLoadAddMomentAtMidSpan;
                                    break;
                                case "CB":
                                    moment = moment;
                                    break;
                                case "CP":
                                    firstLoadAddMomentAtLoad = factor * firstLoadValue * joistLength / 4.0;
                                    moment = moment + firstLoadAddMomentAtLoad;
                                    break;
                                case "CA":
                                    firstLoadAddMomentAtLoad = factor * firstLoadValue * joistLength / 4.0;
                                    moment = moment + firstLoadAddMomentAtLoad;
                                    break;
                                case "C3":
                                    firstLoadAddMomentAtLoad = factor * firstLoadValue * firstLoadDistanceinFt *
                                                               (joistLength - firstLoadDistanceinFt) / joistLength;
                                    moment = moment + firstLoadAddMomentAtLoad;
                                    break;
                                case "CZ":
                                    firstLoadAddMomentAtLoad = factor * firstLoadValue * joistLength / 4.0;
                                    moment = moment + firstLoadAddMomentAtLoad;
                                    break;
                                case "AX":
                                    var depthInFt = double.Parse(joist.Description.Text.Split(new[] { "LH" }, StringSplitOptions.None)[0]) /
                                                    12.0;
                                    var axialLoad = Convert.ToDouble(load.Load1Value.Value);
                                    moment = moment + axialLoad * depthInFt;
                                    break;
                            }
                        }

                        joist.Notes.Add(new StringWithUpdateCheck
                        { Text = string.Format("Mf = {0:F0}<lb-ft>", moment), IsUpdated = false });
                    }

            if (errors != "")
            {
                var filePath = Path.GetTempPath() + "Errors.txt";
                File.WriteAllText(filePath, "Takeoff Errors:\r\n\r\n" + errors);
                Process.Start(filePath);
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
            var excelPath = Path.GetTempFileName();
            File.WriteAllBytes(excelPath, Resources.BLANK_SALES_BOM);

            var oXL2 = Globals.ThisAddIn.Application;
            var workbooks = oXL.Workbooks;
            workbook = workbooks.Open(excelPath);
            var sheets = workbook.Worksheets;
            var sheet = new Worksheet();

            //oXL2.Visible = false;


            var sheetIndex = 6;

            var sheetCount = 0;

            foreach (var sequence in Sequences)
            {
                sheetCount++;
                sheetIndex++;
                Worksheet firstSheetOfSequence = workbook.Worksheets["J(BLANK)"];
                firstSheetOfSequence.Copy(Type.Missing, After: sheets[sheetIndex - 1]);
                firstSheetOfSequence = workbook.Worksheets[sheetIndex];
                firstSheetOfSequence.Name = "J (" + Convert.ToString(sheetCount) + ")";
                sheet = workbook.Worksheets[sheetIndex];
                CellInsert(sheet, 5, 3, sequence.Name.Text, sequence.Name.IsUpdated);

                var row = 7;
                var pageRowCounter = 0;

                for (var markCounter = 0; markCounter < sequence.Joists.Count;)
                {
                    var joist = sequence.Joists[markCounter];

                    var maxRows = Math.Max(joist.Loads.Count, joist.Notes.Count);
                    if (maxRows > 32)
                    {
                        MessageBox.Show(string.Format(
                          "Mark {0} has too many loads on it.\r\n NOTE THAT THIS JOIST WILL NOT BE ADDED TO THE TAKEOFFF!\r\n Either add this joist manually or send to Darien to convert.",
                          joist.Mark.Text));
                        markCounter++;
                        goto SkipLoop;
                    }


                    pageRowCounter = pageRowCounter + Math.Max(Math.Max(joist.Loads.Count, joist.Notes.Count), 1) + 3;
                    if (pageRowCounter > 35)
                    {
                        sheetCount = sheetCount + 1;
                        Worksheet worksheet_copy = workbook.Worksheets["J(BLANK)"];
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

                    var loadRow = row;
                    foreach (var load in joist.Loads)
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

                    var noteRow = row;
                    foreach (var note in joist.Notes)
                    {
                        CellInsert(sheet, noteRow, 26, note.Text, note.IsUpdated);
                        noteRow++;
                    }

                    markCounter++;
                    row = row + Math.Max(Math.Max(joist.Loads.Count, joist.Notes.Count), 1) + 3;

                SkipLoop:;
                }
            }


            //COPY COVER SHEET INTO NEW TAKEOFF
            Worksheet cover = oWB.Sheets["Cover"];
            CellInsert(cover, 2, 10,
              "=INDEX(INDIRECT(\"ProjectTypes[Category]\"),MATCH(INDIRECT(\"ProjectCat\"),INDIRECT(\"ProjectTypes[Type]\"),0))",
              false);
            cover.Copy(Type.Missing, After: workbook.Sheets["Cover"]);
            Worksheet oldCover = workbook.Sheets["Cover"];
            oXL.DisplayAlerts = false;
            oldCover.Delete();
            oXL.DisplayAlerts = true;
            Worksheet newCover = workbook.Sheets["Cover (2)"];
            newCover.Name = "Cover";

            var bridgingRow = 39;
            var columnIndex = 0;
            Bridging = Bridging.Where(br => !(br.Size == "" && br.HorX == "" && br.PlanFeet == 0.0)).ToList();
            var bridgingBySequence = Bridging.GroupBy(br => br.Sequence);

            foreach (var seq in bridgingBySequence)
            {
                if (bridgingRow + seq.Count() > 53)
                {
                    if (columnIndex == 0)
                    {
                        columnIndex = 4;
                        bridgingRow = 39;
                    }
                    else
                    {
                        MessageBox.Show("NOT ENOUGH ROOM FOR BRIDGING, PLEASE ADJUST BRIDGING MANUALLY");
                        if (bridgingRow < 55) bridgingRow = 55;
                    }
                }

                CellInsert(workbook.Sheets["Cover"], bridgingRow, 2 + columnIndex, seq.Key, false);
                bridgingRow++;
                foreach (var br in seq)
                {
                    CellInsert(workbook.Sheets["Cover"], bridgingRow, 2 + columnIndex, br.Size, false);
                    CellInsert(workbook.Sheets["Cover"], bridgingRow, 3 + columnIndex, br.HorX, false);
                    CellInsert(workbook.Sheets["Cover"], bridgingRow, 4 + columnIndex, br.PlanFeet, false);
                    bridgingRow++;
                }

                bridgingRow++;
            }

            //COPY NOTE AND BRIDGING SHEETS INTO NEW TAKEOFF
            foreach (Worksheet s in oWB.Sheets)
            {
                if (s.Name.Contains("N (") && s.Name != "N (0)") s.Copy(Type.Missing, After: workbook.Sheets["Cover"]);
                if (s.Name.Contains("Bridging")) s.Copy(Before: workbook.Sheets["Check List"]);
            }

            //Create Additional Joist Info Tab

            if (this.AdditionalTakeoffInfo.Count != 0)
            {

                Worksheet additionalTakeoffInfoSheet = workbook.Sheets["Additional Takeoff Info"];
                additionalTakeoffInfoSheet.Visible = XlSheetVisibility.xlSheetVisible;
                var numRows = this.AdditionalTakeoffInfo.Count + 1;
                object[,] additionalTakeoffInfoArray = new object[numRows, 3];
                additionalTakeoffInfoArray[0, 0] = "Mark";
                additionalTakeoffInfoArray[0, 1] = "Mf";
                additionalTakeoffInfoArray[0, 2] = "Min I";

                var i = 0;
                foreach (var addInfo in AdditionalTakeoffInfo)
                {
                    i = i + 1;
                    additionalTakeoffInfoArray[i, 0] = addInfo.Key;
                    additionalTakeoffInfoArray[i, 1] = addInfo.Value.Mf;
                    additionalTakeoffInfoArray[i, 2] = addInfo.Value.I;
                }


                additionalTakeoffInfoSheet.Range["A1", "C" + numRows].Value2 = additionalTakeoffInfoArray;
                additionalTakeoffInfoSheet.Copy(Before: workbook.Sheets["J (1)"]);

                oXL.DisplayAlerts = false;
                additionalTakeoffInfoSheet.Delete();
                additionalTakeoffInfoSheet = workbook.Sheets["Additional Takeoff Info (2)"];
                additionalTakeoffInfoSheet.Name = "Additional Takeoff Info";
                oXL.DisplayAlerts = true;

            }

            newCover.Activate();

            Worksheet blankWS = workbook.Sheets["J(BLANK)"];
            oXL.DisplayAlerts = false;
            blankWS.Delete();
            oXL.DisplayAlerts = true;


            var saveFileDialog = new SaveFileDialog();
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
            GC.Collect();
        }

        private void CellInsert(Worksheet sheet, int row, int column, object o, bool isUpdated)
        {
            if (o == null)
            {
            }
            else
            {
                sheet.Cells[row, column] = o;
            }

            if (isUpdated)
            {
                workbook.Worksheets["HighlightedCell"].Range["A1"].Copy();
                sheet.Cells[row, column].PasteSpecial(XlPasteType.xlPasteFormats,
                  XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                if (column == 26)
                {
                    sheet.Range[sheet.Cells[row, 26], sheet.Cells[row, 29]].Merge();
                    sheet.Cells[row, 26].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }
            }
        }

        public void SeperateSeismic()
        {
            foreach (var sequence in Sequences)
                if (sequence.SeperateSeismic)
                    foreach (var joist in sequence.Joists)
                    {
                        //Determine if joist has Seismic Loads

                        var listOfLoadTypes = from load in joist.Loads
                                              select load.LoadInfoCategory.Text;
                        var hasSeismic = false;
                        foreach (var type in listOfLoadTypes)
                            if (type == "SM")
                                hasSeismic = true;

                        if (hasSeismic)
                        {
                            //DETERMINE IF JOIST IS LOAD OVER LOAD. If not then message saying that joist is not load over load and seismic cannot be seperated
                            if (joist.IsLoadOverLoad)
                            {
                                // Move seismic loads to seismic load case

                                foreach (var load in joist.Loads)
                                    if (load.LoadInfoCategory.Text == "SM" &&
                                        (load.CaseNumber.Value == null || load.CaseNumber.Value == 1))
                                        load.CaseNumber.Value = 3;

                                // Copy all other positive loads from LC1 to LC3. 
                                //ISSUES: no important loads can be in any other load case than LC1. 
                                var newLoads = new List<Load>();
                                var copiedLoad = new Load();
                                foreach (var load in joist.Loads)
                                    if ((load.CaseNumber.Value == 1 || load.CaseNumber.Value == null)
                                        && load.Load1Value.Value >= 0
                                        && load.LoadInfoCategory.Text != "WL"
                                        && load.LoadInfoCategory.Text != "IP"
                                        && load.LoadInfoCategory.Text != "R")
                                    {
                                        copiedLoad = DeepClone(load);
                                        copiedLoad.CaseNumber.Value = 3;
                                        newLoads.Add(copiedLoad);
                                    }

                                joist.Loads.AddRange(newLoads);

                                //Switch "R" loads in LC1 to "SL"
                                foreach (var load in joist.Loads)
                                    if (load.LoadInfoCategory.Text == "R")
                                        load.LoadInfoCategory.Text = "SL";

                                //ADD JOIST U DL
                                var uDL = new Load();
                                uDL.LoadInfoType = new StringWithUpdateCheck { Text = "U" };
                                uDL.LoadInfoCategory = new StringWithUpdateCheck { Text = "CL" };
                                uDL.LoadInfoPosition = new StringWithUpdateCheck { Text = "TC" };
                                uDL.Load1Value = new DoubleWithUpdateCheck { Value = joist.UDL };
                                uDL.Load1DistanceFt = new StringWithUpdateCheck { Text = null };
                                uDL.Load1DistanceIn = new DoubleWithUpdateCheck { Value = null };
                                uDL.Load2Value = new DoubleWithUpdateCheck { Value = null };
                                uDL.Load2DistanceFt = new StringWithUpdateCheck { Text = null };
                                uDL.Load2DistanceIn = new DoubleWithUpdateCheck { Value = null };
                                uDL.Reference = new StringWithUpdateCheck { Text = null };
                                uDL.CaseNumber = new DoubleWithUpdateCheck { Value = 3 };
                                joist.Loads.Add(uDL);

                                //ADD JOIST U SM 
                                var uSM = new Load();
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
                                        uSM.Load1Value = new DoubleWithUpdateCheck
                                        { Value = Math.Ceiling((float)(0.14 * sequence.SDS * joist.UDL)) };
                                    else
                                        uSM.Load1Value = new DoubleWithUpdateCheck
                                        { Value = 5 * (int)Math.Ceiling((float)(0.14 * sequence.SDS * joist.UDL / 5.0)) };
                                }

                                uSM.Load1DistanceFt = new StringWithUpdateCheck { Text = null };
                                uSM.Load1DistanceIn = new DoubleWithUpdateCheck { Value = null };
                                uSM.Load2Value = new DoubleWithUpdateCheck { Value = null };
                                uSM.Load2DistanceFt = new StringWithUpdateCheck { Text = null };
                                uSM.Load2DistanceIn = new DoubleWithUpdateCheck { Value = null };
                                uSM.Reference = new StringWithUpdateCheck { Text = null };
                                uSM.CaseNumber = new DoubleWithUpdateCheck { Value = 3 };
                                joist.Loads.Add(uSM);
                            }
                            else
                            {
                                var message = string.Format("MARK {0} IS NOT GIVEN IN TL/LL FORMAT; SEISMIC LC WILL NOT BE SEPERTATED",
                                  joist.Mark.Text);
                                MessageBox.Show(message);
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
            if (joist.Description.Text != null && bT1.Description.Text != null)
                joist.AddError("Base Type description interferes with original; using original ");
            if (joist.Description.Text == null && bT1.Description.Text != null) joist.Description = bT1.Description;
            if (joist.BaseLengthFt.Value != null && bT1.BaseLengthFt.Value != null)
                joist.AddError("Base Type base length ft. interferes with original; using original ");
            if (joist.BaseLengthFt.Value == null && bT1.BaseLengthFt.Value != null) joist.BaseLengthFt = bT1.BaseLengthFt;
            if (joist.BaseLengthIn.Value != null && bT1.BaseLengthIn.Value != null)
                joist.AddError("Base Type base length in. interferes with original; using original ");
            if (joist.BaseLengthIn.Value == null && bT1.BaseLengthIn.Value != null) joist.BaseLengthIn = bT1.BaseLengthIn;
            if (joist.TcxlQuantity.Value != null && bT1.TcxlQuantity.Value != null)
                joist.AddError("Base Type TCXL quantity interferes with original; using original ");
            if (joist.TcxlQuantity.Value == null && bT1.TcxlQuantity.Value != null) joist.TcxlQuantity = bT1.TcxlQuantity;
            if (joist.TcxlLengthFt.Value != null && bT1.TcxlLengthFt.Value != null)
                joist.AddError("Base Type TCXL length ft. interferes with original; using original ");
            if (joist.TcxlLengthFt.Value == null && bT1.TcxlLengthFt.Value != null) joist.TcxlLengthFt = bT1.TcxlLengthFt;
            if (joist.TcxlLengthIn.Value != null && bT1.TcxlLengthIn.Value != null)
                joist.AddError("Base Type TCXL length in. interferes with original; using original ");
            if (joist.TcxlLengthIn.Value == null && bT1.TcxlLengthIn.Value != null) joist.TcxlLengthIn = bT1.TcxlLengthIn;
            if (joist.TcxrQuantity.Value != null && bT1.TcxrQuantity.Value != null)
                joist.AddError("Base Type TCXR quantity interferes with original; using original ");
            if (joist.TcxrQuantity.Value == null && bT1.TcxrQuantity.Value != null) joist.TcxrQuantity = bT1.TcxrQuantity;
            if (joist.TcxrLengthFt.Value != null && bT1.TcxrLengthFt.Value != null)
                joist.AddError("Base Type TCXR length ft. interferes with original; using original ");
            if (joist.TcxrLengthFt.Value == null && bT1.TcxrLengthFt.Value != null) joist.TcxrLengthFt = bT1.TcxrLengthFt;
            if (joist.TcxrLengthIn.Value != null && bT1.TcxrLengthIn.Value != null)
                joist.AddError("Base Type TCXR length in. interferes with original; using original ");
            if (joist.TcxrLengthIn.Value == null && bT1.TcxrLengthIn.Value != null) joist.TcxrLengthIn = bT1.TcxrLengthIn;
            if (joist.SeatDepthLE.Value != null && bT1.SeatDepthLE.Value != null)
                joist.AddError("Base Type LE seat depth interferes with original; using original ");
            if (joist.SeatDepthLE.Value == null && bT1.SeatDepthLE.Value != null) joist.SeatDepthLE = bT1.SeatDepthLE;
            if (joist.SeatDepthRE.Value != null && bT1.SeatDepthRE.Value != null)
                joist.AddError("Base Type RE seat depth interferes with original; using original ");
            if (joist.SeatDepthRE.Value == null && bT1.SeatDepthRE.Value != null) joist.SeatDepthRE = bT1.SeatDepthRE;
            if (joist.BcxQuantity.Value != null && bT1.BcxQuantity.Value != null)
                joist.AddError("Base Type BCX quantity interferes with original; using original ");
            if (joist.BcxQuantity.Value == null && bT1.BcxQuantity.Value != null) joist.BcxQuantity = bT1.BcxQuantity;
            if (joist.Uplift.Value != null && bT1.Uplift.Value != null)
                joist.AddError("Base Type uplift interferes with original; using original ");
            if (joist.Uplift.Value == null && bT1.Uplift.Value != null) joist.Uplift = bT1.Uplift;
            if (joist.Erfos.Text != null && bT1.Erfos.Text != null)
                joist.AddError("Base Type erfos interferes with original; using original ");
            if (joist.Erfos.Text == null && bT1.Erfos.Text != null) joist.Erfos = bT1.Erfos;
            if (joist.DeflectionTL.Value != null && bT1.DeflectionTL.Value != null)
                joist.AddError("Base Type TL deflection interferes with original; using original ");
            if (joist.DeflectionTL.Value == null && bT1.DeflectionTL.Value != null) joist.DeflectionTL = bT1.DeflectionTL;
            if (joist.DeflectionLL.Value != null && bT1.DeflectionLL.Value != null)
                joist.AddError("Base Type LL deflection interferes with original; using original ");
            if (joist.DeflectionLL.Value == null && bT1.DeflectionLL.Value != null) joist.DeflectionLL = bT1.DeflectionLL;
            if (joist.WnSpacing.Text != null && bT1.WnSpacing.Text != null)
                joist.AddError("Base Type WN spacing interferes with original; using original ");
            if (joist.WnSpacing.Text == null && bT1.WnSpacing.Text != null) joist.WnSpacing = bT1.WnSpacing;


            //ADD THE LOADS
            foreach (var load in bT1.Loads)
            {
                var coppiedLoad = DeepClone(load);
                joist.Loads.Add(coppiedLoad);
            }


            //ADD THE NOTES
            foreach (var note in bT1.Notes) joist.Notes.Add(note);
        }

        public void AddQuantitiesFromBB(Takeoff bbTakeoff)
        {
            var blueBeamJoists =
              from s in bbTakeoff.Sequences
              from j in s.Joists
              select j;


            var rg = new Regex(@"\d+");

            var bbJoistTups =
              blueBeamJoists
                .GroupBy(x => x.Mark.Text)
                .Select(g => new Tuple<string, int?>(rg.Match(g.Key).Value, g.Sum(x => x.Quantity.Value)));


            Worksheet marksSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["Marks"];
            var lastUsedRow = marksSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            object[,] array = marksSheet.get_Range("A6", "C" + lastUsedRow).Value2;

            for (var i = 1; i <= array.GetLength(0); i++)
            {
                var marksColValue = array[i, 1];
                if (marksColValue != null)
                {
                    var mark = (string)marksColValue;
                    var bbMatchedJoists = bbJoistTups.Where(joist => joist.Item1 == mark);
                    if (bbMatchedJoists.Any())
                    {
                        var bbJoist = bbMatchedJoists.First();
                        var bbQty = bbJoist.Item2;
                        array[i, 3] = bbQty;
                    }
                    else
                    {
                        MessageBox.Show(string.Format("Takeoff Mark {0} is not in the BlueBeam markups.\r\n\r\n", mark));
                    }
                }
            }

            marksSheet.get_Range("A6", "C" + lastUsedRow).Value2 = array;
        }

        public class Sequence
        {
            public StringWithUpdateCheck Name { get; set; }
            public List<Joist> Joists { get; set; }
            public List<Bridging> Bridging { get; set; }

            public bool SeperateSeismic { get; set; } = false;

            public double? SDS { get; set; }
        }
    }
}