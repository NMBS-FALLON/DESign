using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace DESign_Sales_Excel_Add_in.Worksheet_Values
{
    public class Takeoff
    {
        public class Sequence
        {
            public StringWithUpdateCheck Name { get; set; }
            public List<Joist> Joists { get; set; }
            public List<Bridging> Bridging { get; set; }
        }
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


            // Create a range for the 'BaseLine' tab 
            Excel.Range baseTypesRange = baseTypesWS.UsedRange;

            //Create an object array containing all information from the 'Base Types' tab, in the form of a multidimensional array [row, column]
            object[,] baseTypesCells = (object[,])baseTypesRange.Value2;

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
            int firstBaseTypeRow = 0;

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
                foreach (int rowsForThisBaseType in rowsPerBaseTypeList)
                {
                    BaseType baseType = new BaseType();

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
                    baseType.BcxQuantity = new IntWithUpdateCheck { Value = (int?)(double?)baseTypesCells[rowCount, 13], IsUpdated = isUpdated[rowCount - 1, 12] };
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
                        load.Load1DistanceFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 19], IsUpdated = isUpdated[rowCount + i - 1, 18] };
                        load.Load1DistanceIn = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 20], IsUpdated = isUpdated[rowCount + i - 1, 19] };
                        load.Load2Value = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 21], IsUpdated = isUpdated[rowCount + i - 1, 20] };
                        load.Load2DistanceFt = new DoubleWithUpdateCheck { Value = (double?)baseTypesCells[rowCount + i, 22], IsUpdated = isUpdated[rowCount + i - 1, 21] };
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
            }
            
            ///////////////////

            // Create a range for the 'Marks' tab
            Excel.Range marksRange = marksWS.UsedRange;

            // Create an object array containing all information from the 'Marks' tab, in the form of a multidimensional array [rows, column]
            object[,] marksCells = (object[,])marksRange.Value2;

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
            int firstMarkRow = 0;

            i = 4;
            while (firstLineReached == false)
            {
                if (marksCells[i, 1] != null)
                {
                    firstLineReached = true;
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
                if (i == marksCells.GetLength(0) - 1)
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
                joist.BcxQuantity = new IntWithUpdateCheck { Value = (int?)(double?)marksCells[rowCount, 15], IsUpdated = isUpdated[rowCount - 1, 14] };
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
                    load.Load1DistanceFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 21], IsUpdated = isUpdated[rowCount + i - 1, 20] };
                    load.Load1DistanceIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 22], IsUpdated = isUpdated[rowCount + i - 1, 21] };
                    load.Load2Value = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 23], IsUpdated = isUpdated[rowCount + i - 1, 22] };
                    load.Load2DistanceFt = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 24], IsUpdated = isUpdated[rowCount + i - 1, 23] };
                    load.Load2DistanceIn = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 25], IsUpdated = isUpdated[rowCount + i - 1, 24] };
                    load.CaseNumber = new DoubleWithUpdateCheck { Value = (double?)marksCells[rowCount + i, 26], IsUpdated = isUpdated[rowCount + i - 1, 25] };
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

            //Seperate Sequences
            List<Sequence> sequences = new List<Sequence>();
            var sequenceQuery = from jst in joistLines
                                where jst.BaseTypesOnMark.Count == 0 && jst.Quantity.Value == null && jst.Description.Text == null
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
                

                if (joistLines[0].Quantity.Value != null || joistLines[0].Description.Text != null || joistLines[0].BaseTypesOnMark.Count != 0)
                {
                    MessageBox.Show("Please name your first sequence");
                }
                else
                {
                    Sequence sequence = new Sequence();
                    sequence.Name = new StringWithUpdateCheck { Text = "" };

                    
                    int jstIndex = 0;

                    for(int joistIndex = jstIndex; joistIndex<joistLines.Count; joistIndex ++)
                    {
                        if(joistLines[joistIndex].Quantity.Value == null && joistLines[joistIndex].Description.Text == null && joistLines[joistIndex].BaseTypesOnMark.Count == 0)
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
            takeoff.BaseTypes = baseTypes;
            takeoff.Sequences = sequences;

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


                        //ADD THE LOADS
                        foreach (var bT in matchedBaseType)
                        {
                            //ADD VALUES    ???DO I NEED TO CHECK ANYTHING THAT MAY BE UPDATED??? IF SO HOW TO IMPLEMENT?

                            if (joist.Description.Text != null && bT.Description.Text != null) { MessageBox.Show(string.Format("Mark {0}: Base Type description interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.Description.Text == null && bT.Description.Text != null) { joist.Description = bT.Description; }
                            if (joist.BaseLengthFt.Value != null && bT.BaseLengthFt.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type base length ft. interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.BaseLengthFt.Value == null && bT.BaseLengthFt.Value != null) { joist.BaseLengthFt = bT.BaseLengthFt; }
                            if (joist.BaseLengthIn.Value != null && bT.BaseLengthIn.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type base length in. interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.BaseLengthIn.Value == null && bT.BaseLengthIn.Value != null) { joist.BaseLengthIn = bT.BaseLengthIn; }
                            if (joist.TcxlQuantity.Value != null && bT.TcxlQuantity.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type TCXL quantity interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.TcxlQuantity.Value == null && bT.TcxlQuantity.Value != null) { joist.TcxlQuantity = bT.TcxlQuantity; }
                            if (joist.TcxlLengthFt.Value != null && bT.TcxlLengthFt.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type TCXL length ft. interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.TcxlLengthFt.Value == null && bT.TcxlLengthFt.Value != null) { joist.TcxlLengthFt = bT.TcxlLengthFt; }
                            if (joist.TcxlLengthIn.Value != null && bT.TcxlLengthIn.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type TCXL length in. interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.TcxlLengthIn.Value == null && bT.TcxlLengthIn.Value != null) { joist.TcxlLengthIn = bT.TcxlLengthIn; }
                            if (joist.TcxrQuantity.Value != null && bT.TcxrQuantity.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type TCXR quantity interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.TcxrQuantity.Value == null && bT.TcxrQuantity.Value != null) { joist.TcxrQuantity = bT.TcxrQuantity; }
                            if (joist.TcxrLengthFt.Value != null && bT.TcxrLengthFt.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type TCXR length ft. interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.TcxrLengthFt.Value == null && bT.TcxrLengthFt.Value != null) { joist.TcxrLengthFt = bT.TcxrLengthFt; }
                            if (joist.TcxrLengthIn.Value != null && bT.TcxrLengthIn.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type TCXR length in. interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.TcxrLengthIn.Value == null && bT.TcxrLengthIn.Value != null) { joist.TcxrLengthIn = bT.TcxrLengthIn; }
                            if (joist.SeatDepthLE.Value != null && bT.SeatDepthLE.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type LE seat depth interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.SeatDepthLE.Value == null && bT.SeatDepthLE.Value != null) { joist.SeatDepthLE = bT.SeatDepthLE; }
                            if (joist.SeatDepthRE.Value != null && bT.SeatDepthRE.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type RE seat depth interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.SeatDepthRE.Value == null && bT.SeatDepthRE.Value != null) { joist.SeatDepthRE = bT.SeatDepthRE; }
                            if (joist.BcxQuantity.Value != null && bT.BcxQuantity.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type BCX quantity interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.BcxQuantity.Value == null && bT.BcxQuantity.Value != null) { joist.BcxQuantity = bT.BcxQuantity; }
                            if (joist.Uplift.Value != null && bT.Uplift.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type uplift interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.Uplift.Value == null && bT.Uplift.Value != null) { joist.Uplift = bT.Uplift; }
                            if (joist.Erfos.Text != null && bT.Erfos.Text != null) { MessageBox.Show(string.Format("Mark {0}: Base Type erfos interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.Erfos.Text == null && bT.Erfos.Text != null) { joist.Erfos = bT.Erfos; }
                            if (joist.DeflectionTL.Value != null && bT.DeflectionTL.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type TL deflection interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.DeflectionTL.Value == null && bT.DeflectionTL.Value != null) { joist.DeflectionTL = bT.DeflectionTL; }
                            if (joist.DeflectionLL.Value != null && bT.DeflectionLL.Value != null) { MessageBox.Show(string.Format("Mark {0}: Base Type LL deflection interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.DeflectionLL.Value == null && bT.DeflectionLL.Value != null) { joist.DeflectionLL = bT.DeflectionLL; }
                            if (joist.WnSpacing.Text != null && bT.WnSpacing.Text != null) { MessageBox.Show(string.Format("Mark {0}: Base Type WN spacing interferes with original; using original ", joist.Mark.Text)); }
                            if (joist.WnSpacing.Text == null && bT.WnSpacing.Text != null) { joist.WnSpacing = bT.WnSpacing; }



                            //ADD THE LOADS
                            foreach (Load load in bT.Loads)
                            {
                                joist.Loads.Add(load);

                            }
                            //ADD THE NOTES
                            foreach (StringWithUpdateCheck note in bT.Notes)
                            {
                                joist.Notes.Add(note);
                            }
                        }
                    }
                }
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
            System.IO.File.WriteAllBytes(excelPath, Properties.Resources.BLANK_SALES_BOM);

            Excel.Application oXL2 = new Excel.Application();
            Excel.Workbooks workbooks = oXL2.Workbooks;
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
                        CellInsert(sheet, loadRow, 20, load.Load1DistanceFt.Value, load.Load1DistanceFt.IsUpdated);
                        CellInsert(sheet, loadRow, 21, load.Load1DistanceIn.Value, load.Load1DistanceIn.IsUpdated);
                        CellInsert(sheet, loadRow, 22, load.Load2Value.Value, load.Load2Value.IsUpdated);
                        CellInsert(sheet, loadRow, 23, load.Load2DistanceFt.Value, load.Load2DistanceFt.IsUpdated);
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

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsm)|*.xlsm";
            saveFileDialog.ShowDialog();
            if (saveFileDialog.FileName != "")
            {
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

            }
        }

        public void SeperateSeismic(double sds = 0.00)
        {
            foreach (Sequence sequence in Sequences)
            {
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

                            // SET SEISMIC LC TO 6. 
                            var listOfLCs = from load in joist.Loads
                                            select Convert.ToInt32(load.CaseNumber.Value);

                            int seismicLC = 3;
                            if (listOfLCs.Contains(3) == true)
                            {
                                MessageBox.Show(String.Format("MARK {0}: LC 3 MUST BE AVAILABLE FOR SEISMIC SEPERATION; ENDING PROGRAM",
                                    joist.Mark.Text));
                            }

                            // Move seismic loads to seismic load case

                            foreach (Load load in joist.Loads)
                            {
                                if (load.LoadInfoCategory.Text == "SM")
                                {
                                    load.CaseNumber.Value = seismicLC;
                                }
                            }

                            // Copy all other positive loads from LC1 to LC3. 
                            //ISSUES: no important loads can be in any other load case than LC1. 
                            List<Load> newLoads = new List<Load>();
                            Load copiedLoad = new Load();
                            foreach (Load load in joist.Loads)
                            {
                                if ((load.CaseNumber.Value == 1 || load.CaseNumber.Value == null) && load.Load1Value.Value >= 0)
                                {

                                    copiedLoad = DeepClone(load);
                                    copiedLoad.CaseNumber.Value = seismicLC;
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
                            uDL.Load1DistanceFt = new DoubleWithUpdateCheck { Value = null };
                            uDL.Load1DistanceIn = new DoubleWithUpdateCheck { Value = null };
                            uDL.Load2Value = new DoubleWithUpdateCheck { Value = null };
                            uDL.Load2DistanceFt = new DoubleWithUpdateCheck { Value = null };
                            uDL.Load2DistanceIn = new DoubleWithUpdateCheck { Value = null };
                            uDL.LoadNote = new StringWithUpdateCheck { Text = null };
                            uDL.CaseNumber = new DoubleWithUpdateCheck { Value = seismicLC };
                            joist.Loads.Add(uDL);

                            //ADD JOIST U SM 
                            Load uSM = new Load();
                            uSM.LoadInfoType = new StringWithUpdateCheck { Text = "U" };
                            uSM.LoadInfoCategory = new StringWithUpdateCheck { Text = "SM" };
                            uSM.LoadInfoPosition = new StringWithUpdateCheck { Text = "TC" };
                            uSM.Load1Value = new DoubleWithUpdateCheck { Value = 0.14 * sds * joist.UDL };
                            uSM.Load1DistanceFt = new DoubleWithUpdateCheck { Value = null };
                            uSM.Load1DistanceIn = new DoubleWithUpdateCheck { Value = null };
                            uSM.Load2Value = new DoubleWithUpdateCheck { Value = null };
                            uSM.Load2DistanceFt = new DoubleWithUpdateCheck { Value = null };
                            uSM.Load2DistanceIn = new DoubleWithUpdateCheck { Value = null };
                            uSM.LoadNote = new StringWithUpdateCheck { Text = null };
                            uSM.CaseNumber = new DoubleWithUpdateCheck { Value = seismicLC };
                            joist.Loads.Add(uSM);

                        }
                        else
                        {
                            string message = String.Format("MARK {0} IS NOT GIVEN IN TL/LL FORMAT; SEISMIC LC WILL NOT BE SEPERTATED", joist.Mark.Text);
                            MessageBox.Show(message);
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
    }
}
