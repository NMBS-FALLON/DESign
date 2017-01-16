using System;
using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
//using System.Threading;
//using System.Collections.Concurrent;
//using NMBS_2;
using System.Runtime.InteropServices;


namespace DESign_BOT
{

    public partial class FormNMBS_AB : Form
    {
        ExcelDataExtraction ExcelDataExtraction = new ExcelDataExtraction();

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

                oWB = oXL.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                oRng = oSheet.get_Range("B3", Missing.Value);

                oSheet.get_Range("B2", Missing.Value).Value = "MARKS:";
                oSheet.get_Range("C2", Missing.Value).Value = "A:";
                oSheet.get_Range("D2", Missing.Value).Value = "B:";

                oSheet.get_Range("B2", Missing.Value).Font.Bold = true;
                oSheet.get_Range("C2", Missing.Value).Font.Bold = true;
                oSheet.get_Range("D2", Missing.Value).Font.Bold = true;

                for (int i = 0; i < BOMMarks.Count; i++)
                {
                    int cellNumber = 3 + i;
                    oRng = oSheet.get_Range("B" + cellNumber, Missing.Value);

                    oRng.Value = BOMMarks[i];
                }

                Excel.Range oRngAs;
                Excel.Range oRngBs;

                for (int i = 0; i < BOMMarks.Count; i++)
                {

                    int cellNumber = 3 + i;
                    oRngAs = oSheet.get_Range("C" + cellNumber, Missing.Value);
                    oRngBs = oSheet.get_Range("D" + cellNumber, Missing.Value);

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
                        }
                    }
                   
                }

                Excel.Range last = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = last.Row;
                oRng = oSheet.get_Range("B3", "D" + lastUsedRow);

                object[,] stringJoistMarks = (object[,])oRng.Value2;

                for (int row = 1; row <= stringJoistMarks.GetLength(0); row++)
                {
                    if (dataGridView1.Rows[0].Cells[0].Value == "U.N.O.")
                    {
                        if (formAs[0] != null)
                        {
                            if (stringJoistMarks[row, 2] == null)
                            {
                                stringJoistMarks[row, 2] = formAs[0];
                            }
                        }
                        if (formBs[0] != null)
                        {
                            if (stringJoistMarks[row, 3] == null)
                            {
                                stringJoistMarks[row, 3] = formBs[0];
                            }
                        }
                    }
                }
                oRng.Value2 = stringJoistMarks;
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
