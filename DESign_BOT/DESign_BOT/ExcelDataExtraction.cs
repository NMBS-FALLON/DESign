using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using DESign_BOT;


namespace DESign_BOT
{
    class ExcelDataExtraction
    {
        StringManipulation stringManipulation = new StringManipulation();


        public List<List<string>> exlNailerValues()
        {
            object[,] stringJoistMarks = null;

            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "|*.xlsx;*.xlsm";


            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelFileName = openFileDialog.FileName;

                Excel.Range oRng = null;
                Excel._Worksheet oSheet = null;
                Excel._Workbook oWB = null;
                Excel.Workbooks oWorkbooks = null;
                Excel.Application oXL = null;
                
                

                try
                {
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    oXL.Visible = false;

                    //Get a new workbook.

                    oWorkbooks = oXL.Workbooks;
                    oWB = oWorkbooks.Open(excelFileName);
                   
                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    oRng = oSheet.UsedRange;


                    oRng.get_Range("B3", Missing.Value);
                    oRng = oRng.get_End(Excel.XlDirection.xlToRight);
                    oRng = oRng.get_End(Excel.XlDirection.xlDown);
                    string downJoistMarks = oRng.get_Address(Excel.XlReferenceStyle.xlA1, Type.Missing);
                    oRng = oSheet.get_Range("B3", downJoistMarks);
                    stringJoistMarks = (object[,])oRng.Value2;

                    oWB.Close(0);
                    oXL.Quit();

                    Marshal.ReleaseComObject(oRng);
                    Marshal.ReleaseComObject(oSheet);
                    Marshal.ReleaseComObject(oWB);
                    Marshal.ReleaseComObject(oWorkbooks);
                    Marshal.ReleaseComObject(oXL);
                    GC.Collect();
                    

                }

                catch (Exception theException)
                {
                    
                    Marshal.ReleaseComObject(oRng);
                    Marshal.ReleaseComObject(oSheet);    
                    Marshal.ReleaseComObject(oWB);
                    Marshal.ReleaseComObject(oWorkbooks);
                    Marshal.ReleaseComObject(oXL);
                    GC.Collect();
                    
                    

                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, "Line:");
                    errorMessage = String.Concat(errorMessage, theException.Source);

                    MessageBox.Show(errorMessage, "Error");
                } 

            }


            List<string> AsfromExcel = new List<string>();
            List<string> BsfromExcel = new List<string>();
            List<string> JoistMarksfromExcel = new List<string>();



            for (int i = 1; i <= stringJoistMarks.GetLength(0); i++)
            {
                string stringA = stringManipulation.convertLengthStringtoHyphenLength(stringJoistMarks[i, 2].ToString());
                string stringB = stringManipulation.convertLengthStringtoHyphenLength(stringJoistMarks[i, 3].ToString());
                JoistMarksfromExcel.Add(stringJoistMarks[i, 1].ToString());
                AsfromExcel.Add(stringA);
                BsfromExcel.Add(stringB);
            }

            List<List<string>> exlNailerData = new List<List<string>>();

            exlNailerData.Add(JoistMarksfromExcel);
            exlNailerData.Add(AsfromExcel);
            exlNailerData.Add(BsfromExcel);



            return exlNailerData;



        }
        public enum JorG { Joist, Girder };

        public struct OWSJ
        {
            public string Mark;
            public string Notes;
            public JorG JorG;

            public OWSJ(string mark, JorG jORg, string notes)
            {
                Mark = mark;
                Notes = notes;
                JorG = jORg;
            }

        }

        public List<OWSJ> getBOMOWSJs()
        {
            OpenFileDialog openBOMFileDialog = new OpenFileDialog();

            openBOMFileDialog.Filter = "|*.xlsx;*.xlsm";


            //List<List<string>> BOMMarksAndNotes = new List<List<string>>();
            var bomMarks = new List<OWSJ>();



            if (openBOMFileDialog.ShowDialog() == DialogResult.OK)
            {
                //    string excelFileName = openBOMFileDialog.FileName;

                string excelFileName = System.IO.Path.GetTempFileName();
                Byte[] BOMinByteArray = System.IO.File.ReadAllBytes(openBOMFileDialog.FileName);
                System.IO.File.WriteAllBytes(excelFileName, BOMinByteArray);

                Excel.Application oXL = null;
                Excel.Workbooks oWorkBooks = null;
                Excel._Workbook oWB = null;
                Excel.Sheets oSheets = null;
                Excel._Worksheet oSheet = null;
                Excel.Range oRngMarks = null;
                Excel.Range oRngNotes = null;
                try
                {
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    oXL.Visible = false;

                    //Get a new workbook.

                    oWorkBooks = oXL.Workbooks;
                    oWB = oWorkBooks.Open(excelFileName);
                oSheets = oWB.Sheets;


                    List<int> joistWorksheetIndices = new List<int>();

                    for (int i = 1; i <= oWB.Sheets.Count; i++)
                    {
                        oSheet = (Excel.Worksheet)oSheets.get_Item(i);
                        string worksheetName = oSheet.Name;
                        if (worksheetName.Contains("J") == true && worksheetName.Contains("(") == true)
                        {
                            joistWorksheetIndices.Add(i);
                        }
                    }

                    for (int i = 0; i < joistWorksheetIndices.Count; i++)
                    {
                        oSheet = (Excel._Worksheet)oSheets.get_Item(joistWorksheetIndices[i]);
                        for (int k = 16; k < 46; k++)
                        {
                            oRngMarks = oSheet.get_Range("A" + k, Missing.Value);
                            oRngNotes = oSheet.get_Range("AA" + k, Missing.Value);

                            string stringoRngMarks = (string)oRngMarks.Text;
                            string stringoRngNotes = (string)oRngNotes.Text;
                            if (stringoRngMarks != "" && stringoRngMarks != "MARK")
                            {
                                bomMarks.Add(new OWSJ(stringoRngMarks, JorG.Joist, stringoRngNotes));
                            }
                        }
                    }

                    List<int> girderWorksheetIndices = new List<int>();

                    for (int i = 1; i <= oWB.Sheets.Count; i++)
                    {
                        oSheet = (Excel.Worksheet)oSheets.get_Item(i);
                        string workSheetName = oSheet.Name;
                        if (workSheetName.Contains("G") == true && workSheetName.Contains("(") == true)
                        {
                            girderWorksheetIndices.Add(i);
                        }
                    }

                    for (int i = 0; i < girderWorksheetIndices.Count; i++)
                    {
                        oSheet = (Excel._Worksheet)oSheets.get_Item(girderWorksheetIndices[i]);
                        for (int k = 14; k < 46; k++)
                        {
                            oRngMarks = oSheet.get_Range("A" + k, Missing.Value);
                            oRngNotes = oSheet.get_Range("Z" + k, Missing.Value);

                            string stringoRngMarks = (string)oRngMarks.Text;
                            string stringoRngNotes = (string)oRngNotes.Text;
                            if (stringoRngMarks != "" && stringoRngMarks != "MARK")
                            {
                                bomMarks.Add(new OWSJ(stringoRngMarks, JorG.Girder, stringoRngNotes));
                            }
                        }
                    }

                    Marshal.ReleaseComObject(oRngMarks);
                    Marshal.ReleaseComObject(oRngNotes);
                    Marshal.ReleaseComObject(oSheet);
                    Marshal.ReleaseComObject(oSheets);
                    Marshal.ReleaseComObject(oWB);
                    Marshal.ReleaseComObject(oWorkBooks);
                    Marshal.ReleaseComObject(oXL); ;
                    System.GC.Collect();
                    

                }

                catch (Exception theException)
                {
                    
                    Marshal.ReleaseComObject(oRngMarks);
                    Marshal.ReleaseComObject(oRngNotes);
                    Marshal.ReleaseComObject(oSheet);
                    Marshal.ReleaseComObject(oSheets);
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


            return bomMarks;
        }







    }
}
