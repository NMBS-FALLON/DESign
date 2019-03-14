using System;
using System.Collections.Generic;

using System.Linq;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;

using System.Runtime.InteropServices;


namespace DESign_BOT
{

    public partial class FormNMBS_AB : Form
    {
        ExcelDataExtraction ExcelDataExtraction = new ExcelDataExtraction();
        StringManipulation sm = new StringManipulation();

        public FormNMBS_AB()
        {
            InitializeComponent();
            dataGridView1.Rows.Add();
            dataGridView1.Rows[0].Cells[0].Value = "U.N.O. (JOISTS)";
            dataGridView1.Rows.Add();
            dataGridView1.Rows[1].Cells[0].Value = "U.N.O. (GIRDERS)";
        }

        public struct FormNote
        {
            public string Note;
            public string A;
            public string B;
            public string Spacing;
            public string HC;

            public FormNote (string note, string a, string b, string spacing, string hc)
            {
                Note = note;
                A = a;
                B = b;
                Spacing = spacing;
                HC = hc;
            }
        }

        public struct NoteInfo
        {
            public string Mark;
            public string A;
            public string B;
            public string Spacing;
            public string HC;
            public List<string> Errors;

            public NoteInfo(string mark, string a, string b, string spacing, string hc, List<string> errors)
            {
                Mark = mark;
                A = a;
                B = b;
                Spacing = spacing;
                HC = hc;
                Errors = errors;
            }
        }

        public string[] GetNoteArray(string notes)
        {
            Char[] delimChars = { '[', ',', ']', ' ' };
            string[] notesArray = notes.Split(delimChars, StringSplitOptions.RemoveEmptyEntries);
            return notesArray;

        }


        private void btnBOMtoExcel_Click(object sender, EventArgs e)
        {
            // Get all notes from the dataGridView
            // ----------------------------------------------------------------------------------------------------------------------------
            var notesFromForm = new List<FormNote>();

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                var noteObj = dataGridView1.Rows[i].Cells[0].Value;
                var aObj = dataGridView1.Rows[i].Cells[1].Value;
                var bObj = dataGridView1.Rows[i].Cells[2].Value;
                var spacingObj = dataGridView1.Rows[i].Cells[3].Value;
                var hcObj = dataGridView1.Rows[i].Cells[4].Value;

                string note = noteObj == null ? null : noteObj.ToString();
                string a = aObj == null ? null : aObj.ToString();
                string b = bObj == null ? null : bObj.ToString();
                string spacing = spacingObj == null ? null : spacingObj.ToString();
                string hc = hcObj == null ? null : hcObj.ToString();

                notesFromForm.Add(new FormNote(note, a, b, spacing, hc));
            }
            // ----------------------------------------------------------------------------------------------------------------------------



            // Create a list of all noteInfos called allNoteInfo
            // ----------------------------------------------------------------------------------------------------------------------------
            var bomOWSJs = ExcelDataExtraction.getBOMOWSJs();
            var allNoteInfo = new List<NoteInfo>();

            foreach (var owsj in bomOWSJs)
            {
                string mark = owsj.Mark;
                string a = null;
                string b = null;
                string spacing = null;
                string hc = null;
                var errors = new List<string>();


                // Apply applicable notes
                // ----------------------------------------------------------------------------------------------------------------------------
                var notesArray = GetNoteArray(owsj.Notes);
                foreach (var note in notesArray)
                {
                    var applicableNotes = notesFromForm.Where(n => n.Note == note);
                    foreach (var applicableNote in applicableNotes)
                    {
                        // Add error if a info value has already been set and this applicableNote also has a value for that info
                        if (a != null && applicableNote.A != null) { errors.Add("Multiple 'A' values set on Mark " + mark); }
                        if (b != null && applicableNote.B != null) { errors.Add("Multiple 'B' values set on Mark " + mark); }
                        if (spacing != null && applicableNote.Spacing != null) { errors.Add("Multiple 'Spacing' values set on Mark " + mark); }
                        if (hc != null && applicableNote.HC != null) { errors.Add("Multiple 'HC' values set on Mark " + mark); }

                        // set info values to the applicableNote value if the applicableNote value is not null
                        a = applicableNote.A != null ? applicableNote.A : a;
                        b = applicableNote.B != null ? applicableNote.B : b;
                        spacing = applicableNote.Spacing != null ? applicableNote.Spacing : spacing;
                        hc = applicableNote.HC != null ? applicableNote.HC : hc;
                    }
                }
                // ----------------------------------------------------------------------------------------------------------------------------


                // Apply Default Values
                // ----------------------------------------------------------------------------------------------------------------------------
                var defaultIE = owsj.JorG == ExcelDataExtraction.JorG.Joist ?
                                    notesFromForm.Where(n => n.Note == "U.N.O. (JOISTS)")
                                    : notesFromForm.Where(n => n.Note == "U.N.O. (GIRDERS)");
                var hasDefault = defaultIE.Count() > 0;

                if (hasDefault)
                {
                    // set info values to the default value if the current info value is null
                    var defaultInfo = defaultIE.First();
                    a = a == null ? defaultInfo.A : a;
                    b = b == null ? defaultInfo.B : b;
                    spacing = spacing == null ? defaultInfo.Spacing : spacing;
                    hc = hc == null ? defaultInfo.HC : hc;
                }
                // ----------------------------------------------------------------------------------------------------------------------------

                //add current noteInfo to allNoteInfo
                var noteInfo = new NoteInfo(mark, a, b, spacing, hc, errors);
                allNoteInfo.Add(noteInfo);

            }
            // ----------------------------------------------------------------------------------------------------------------------------

            // Convert 'allNoteInfo' into an object[,]
            // ----------------------------------------------------------------------------------------------------------------------------
            object[,] allNoteInfoArray = new object[allNoteInfo.Count, 5];

            for (int i = 0; i < allNoteInfo.Count; i++)
            {
                var noteInfo = allNoteInfo[i];
                allNoteInfoArray[i, 0] = noteInfo.Mark;
                allNoteInfoArray[i, 1] = noteInfo.A;
                allNoteInfoArray[i, 2] = noteInfo.B;
                allNoteInfoArray[i, 3] = noteInfo.Spacing;
                allNoteInfoArray[i, 4] = noteInfo.HC;
            }
            // ----------------------------------------------------------------------------------------------------------------------------


            // create array of errors;
            // ----------------------------------------------------------------------------------------------------------------------------
            var allErrors = allNoteInfo
                            .SelectMany(n => n.Errors);

            object[,] allErrorsArray = new object[allErrors.Count(), 1];
            var errorCount = 0;
            foreach (var error in allErrors)
            {
                allErrorsArray[errorCount, 0] = error;
                errorCount++;
            }
            // ----------------------------------------------------------------------------------------------------------------------------


            // Create Excel workbook and 'paste' in allNoteInfoArray and allErrorsArray

            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks oWorkBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Range oRng = null;
            try
            {


                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.

                string excelPath = System.IO.Path.GetTempFileName();

                System.IO.File.WriteAllBytes(excelPath, Properties.Resources.BLANK_AB_SHEET);

                oWorkBooks = oXL.Workbooks;
                oWB = oWorkBooks.Open(excelPath);


                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                var numMarks = allNoteInfoArray.GetLength(0);
                oRng = oSheet.get_Range("B7", "F" + (numMarks + 6).ToString());

                oRng.Value2 = allNoteInfoArray;

                var numErrors = allErrorsArray.GetLength(0);
                oRng = oSheet.get_Range("J7", "J" + (numErrors + 6).ToString());
                oRng.Value2 = allErrorsArray;

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.ShowDialog();
                if (saveFileDialog.FileName != "")
                {
                    oWB.CheckCompatibility = false;
                    oWB.SaveAs(saveFileDialog.FileName);
                }

                Marshal.ReleaseComObject(oRng);
                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oWorkBooks);
                Marshal.ReleaseComObject(oXL); ;
                System.GC.Collect();
                

            }

            catch (Exception theException)
            {
                Marshal.ReleaseComObject(oRng);
                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oWorkBooks);
                Marshal.ReleaseComObject(oXL); ;
                System.GC.Collect();
                

                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, "Line:");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");

            } 

        }

        private void datagridview1_SelectionChanged(object sender, EventArgs e)
        {
            this.dataGridView1.ClearSelection();
        }


        private void formBOMtoExcel_Load(object sender, EventArgs e)
        {

        }


    }
}
