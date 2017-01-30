using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace DESign_Sales_Excel_Add_in.Worksheet_Values
{
    public class Takeoff
    {
        public List<BaseType> BaseTypes { get; set; }
        public List<Joist> Joists { get; set; }
        public List<Bridging> Bridging { get; set; }

        public Takeoff ImportTakeoff()
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
            while (firstMarkReached == false)
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
                    load.LoadInfoType = new StringWithUpdateCheck { Text = (string)marksCells[rowCount + i, 17] };
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
                    if (load.IsNull == false)
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


            // ADD BASE TYPES TO JOISTS

            foreach (var joist in takeoff.Joists)
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
                        foreach(StringWithUpdateCheck note in bT.Notes)
                        {
                            joist.Notes.Add(note);
                        }


                    }
                    

                    
                }
            }



            // RETURN COMPLETE TAKEOFF
            return takeoff;

        }

        public void CreateOriginalTakeoff(Takeoff takeoff)
        {
            string excelPath = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllBytes(excelPath, Properties.Resources.BLANK_SALES_BOM);

            Excel.Application oXL = new Excel.Application();
            Excel.Workbooks workbooks = oXL.Workbooks;
            Excel.Workbook workbook = workbooks.Open(excelPath);
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = workbook.ActiveSheet;

            oXL.Visible = false;
            workbook.SaveAs(Environment.GetFolderPath(
                Environment.SpecialFolder.Desktop) + "/NEW TAKEOFF",
                Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
            sheet = workbook.Worksheets["J (1)"];
            int sheetIndex = sheet.Index;

            int row = 7;
            int pageRowCounter = 0;
            int sheetCount = 1;
            for (int markCounter = 0; markCounter < takeoff.Joists.Count;)
            {
                Joist joist = takeoff.Joists[markCounter];


                pageRowCounter = pageRowCounter + Math.Max(joist.Loads.Count, joist.Notes.Count) + 3;
                if (pageRowCounter > 34)
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
                CellInsert(sheet, row, 1, joist.Mark.Text);
                CellInsert(sheet, row, 2, joist.Quantity.Value);
                CellInsert(sheet, row, 3, joist.Description.Text);
                CellInsert(sheet, row, 4, joist.BaseLengthFt.Value);
                CellInsert(sheet, row, 5, joist.BaseLengthIn.Value);
                CellInsert(sheet, row, 6, joist.TcxlQuantity.Value);
                CellInsert(sheet, row, 7, joist.TcxlLengthFt.Value);
                CellInsert(sheet, row, 8, joist.TcxlLengthIn.Value);
                CellInsert(sheet, row, 9, joist.TcxrQuantity.Value);
                CellInsert(sheet, row, 10, joist.TcxrLengthFt.Value);
                CellInsert(sheet, row, 11, joist.TcxrLengthIn.Value);
                CellInsert(sheet, row, 12, joist.SeatDepthLE.Value);
                CellInsert(sheet, row, 13, joist.SeatDepthRE.Value);
                CellInsert(sheet, row, 14, joist.BcxQuantity.Value);
                CellInsert(sheet, row, 15, joist.Uplift.Value);

                int loadRow = row;
                foreach(Load load in joist.Loads)
                {
                    CellInsert(sheet, loadRow, 16, load.LoadInfoType.Text);
                    CellInsert(sheet, loadRow, 17, load.LoadInfoCategory.Text);
                    CellInsert(sheet, loadRow, 18, load.LoadInfoPosition.Text);
                    CellInsert(sheet, loadRow, 19, load.Load1Value.Value);
                    CellInsert(sheet, loadRow, 20, load.Load1DistanceFt.Value);
                    CellInsert(sheet, loadRow, 21, load.Load1DistanceIn.Value);
                    CellInsert(sheet, loadRow, 22, load.Load2Value.Value);
                    CellInsert(sheet, loadRow, 23, load.Load2DistanceFt.Value);
                    CellInsert(sheet, loadRow, 24, load.Load2DistanceFt.Value);
                    CellInsert(sheet, loadRow, 25, load.CaseNumber.Text);

                    loadRow++;
                }

                int noteRow = row;
                foreach(StringWithUpdateCheck note in joist.Notes)
                {
                    CellInsert(sheet, noteRow, 26, note);
                    //sheet.Cells[noteRow, 26] = note;
                    noteRow++;
                }

                markCounter++;
                row = row + Math.Max(Math.Max(joist.Loads.Count, joist.Notes.Count), 1) + 3;

            SkipLoop:
                ;
            }

            oXL.Visible = true;
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(oXL);
            GC.Collect();

        }
        private void CellInsert(Excel.Worksheet sheet, int row, int column, object o)
        {
            if (o == null) { }
            else
            {
                sheet.Cells[row, column] = o;
            }
        }
    }
}
