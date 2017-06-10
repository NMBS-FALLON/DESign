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
            dataGridView1.Rows[0].Cells[0].Value = "U.N.O.";
        }


        private void btnBOMtoExcel_Click(object sender, EventArgs e)
        {
            List<List<string>> BOMMarksAndNotes = ExcelDataExtraction.BOMMarksAndNotes();

            List<string> BOMMarks = BOMMarksAndNotes[0];
            List<string> BOMNotes = BOMMarksAndNotes[1];

            List<string> formNotes = new List<string>();
            List<string> formAs = new List<string>();
            List<string> formBs = new List<string>();
            List<string> formSpacings = new List<string>();


            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value != null)
                {
                    formNotes.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                    if (dataGridView1.Rows[i].Cells[1].Value != null)
                    {
                        formAs.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    }
                    else if (dataGridView1.Rows[i].Cells[1].Value == null)
                        formAs.Add("");

                    if (dataGridView1.Rows[i].Cells[2].Value != null)
                    {
                        formBs.Add(dataGridView1.Rows[i].Cells[2].Value.ToString());
                    }
                    else if (dataGridView1.Rows[i].Cells[2].Value == null)
                    {
                        formBs.Add("");
                    }
                    if (dataGridView1.Rows[i].Cells[3].Value != null)
                    {
                        formSpacings.Add(dataGridView1.Rows[i].Cells[3].Value.ToString());
                    }
                    else if (dataGridView1.Rows[i].Cells[3].Value == null)
                    {
                        formSpacings.Add("");
                    }

                }

            }

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;
            try
            {


                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.

                string excelPath = System.IO.Path.GetTempFileName();

                System.IO.File.WriteAllBytes(excelPath, Properties.Resources.BLANK_AB_SHEET);

                oWB = oXL.Workbooks.Open(excelPath);


                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                oRng = oSheet.get_Range("B5", Missing.Value);


                //Get a new workbook.


                for (int i = 0; i < BOMMarks.Count; i++)
                {
                    int cellNumber = 6 + i;
                    oRng = oSheet.get_Range("B" + cellNumber, Missing.Value);
                    oRng.Value = BOMMarks[i];
                }
                Excel.Range oRngAs;
                Excel.Range oRngBs;
                Excel.Range oRngSpacing;

                for (int i = 0; i < BOMMarks.Count; i++)
                {

                    int cellNumber = 6 + i;
                    oRngAs = oSheet.get_Range("C" + cellNumber, Missing.Value);
                    oRngBs = oSheet.get_Range("D" + cellNumber, Missing.Value);
                    oRngSpacing = oSheet.get_Range("E" + cellNumber, Missing.Value);

                    string[] alpha = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S" };

                    Char[] delimChars = { '[', ',', ']', ' ' };
                    string[] BOMNotesArray = BOMNotes[i].Split(delimChars, StringSplitOptions.RemoveEmptyEntries);
                    for (int k = 0; k < formNotes.Count; k++)
                    {
                        int alphaIndex = 5;
                        if (BOMNotesArray.Contains(formNotes[k]))
                        {
                            if (formAs[k].ToString() != "")
                            {

                                if ((string)oRngAs.Text == "")
                                {
                                    oRngAs.Value = "'" + formAs[k].ToString();
                                }
                                else
                                {
                                    oRngAs = oSheet.get_Range(alpha[alphaIndex] + cellNumber, Missing.Value);
                                    oRngAs.Value = "'" + formAs[k].ToString();
                                    alphaIndex++;
                                }
                            }
                            if (formBs[k].ToString() != "")
                            {
                                if ((string)oRngBs.Text == "")
                                {
                                    oRngBs.Value = "'" + formBs[k].ToString();
                                }
                                else
                                {
                                    oRngBs = oSheet.get_Range(alpha[alphaIndex] + cellNumber, Missing.Value);
                                    oRngBs.Value = "'" + formBs[k].ToString();
                                    alphaIndex++;
                                }

                            }
                            if (formSpacings[k].ToString() != "")
                            {
                                if ((string)oRngSpacing.Text == "")
                                {
                                    oRngSpacing.Value = "'" + formSpacings[k].ToString();
                                }
                                else
                                {
                                    oRngSpacing = oSheet.get_Range(alpha[alphaIndex] + cellNumber, Missing.Value);
                                    oRngSpacing.Value = "'" + formSpacings[k].ToString();
                                    alphaIndex++;
                                }

                            }
                        }
                    }
                   
                }

                Excel.Range last = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = last.Row;
                oRng = oSheet.get_Range("B6", "E" + lastUsedRow);

                object[,] stringJoistMarks = (object[,])oRng.Value2;

                for (int row = 1; row <= stringJoistMarks.GetLength(0); row++)
                {
                    if (dataGridView1.Rows[0].Cells[0].Value == "U.N.O.")
                    {
                        if (formAs[0] != null)
                        {
                            if (stringJoistMarks[row, 2] == null)
                            {
                                stringJoistMarks[row, 2] = "'" + formAs[0];
                            }
                        }
                        if (formBs[0] != null)
                        {
                            if (stringJoistMarks[row, 3] == null)
                            {
                                stringJoistMarks[row, 3] = "'" + formBs[0];
                            }
                        }
                        if (formSpacings[0] != null)
                        {
                            if (stringJoistMarks[row, 4] == null)
                            {
                                    stringJoistMarks[row, 4] = "'" + formSpacings[0];                                
                            }
                        }
                    }
                }
                oRng.Value2 = stringJoistMarks;

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.ShowDialog();
                if (saveFileDialog.FileName != "")
                {
                    oWB.CheckCompatibility = false;
                    oWB.SaveAs(saveFileDialog.FileName);
                }


            }

            catch (Exception theException)
            {
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
