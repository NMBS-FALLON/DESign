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


         public List<List<string>> exlNailerValues ()
        {
            object[,] stringJoistMarks = null;

            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "|*.xlsx;*.xlsm";

          
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelFileName = openFileDialog.FileName;
                
                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;

                try
                   {
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    oXL.Visible = false;

                    //Get a new workbook.


                    oWB = oXL.Workbooks.Open(excelFileName);
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
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);
                Marshal.ReleaseComObject(oSheet);

                }

                catch(Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage=String.Concat(errorMessage, theException.Message);
                    errorMessage=String.Concat(errorMessage, "Line:");
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
                 JoistMarksfromExcel.Add(stringJoistMarks[i,1].ToString());
                 AsfromExcel.Add(stringA);
                 BsfromExcel.Add(stringB);
             }

             List<List<string>> exlNailerData = new List<List<string>>();

             exlNailerData.Add(JoistMarksfromExcel);
             exlNailerData.Add(AsfromExcel);
             exlNailerData.Add(BsfromExcel);

             

             return exlNailerData;


         
    }
        public List<List<string>> BOMMarksAndNotes()
         {
             OpenFileDialog openBOMFileDialog = new OpenFileDialog();

             openBOMFileDialog.Filter = "|*.xlsx;*.xlsm";

              
              List<List<string>> BOMMarksAndNotes = new List<List<string>>();
              List<string> BOMjoistMarks = new List<string>();
              List<string> BOMjoistNotes = new List<string>();



             if (openBOMFileDialog.ShowDialog() == DialogResult.OK)
             {
             //    string excelFileName = openBOMFileDialog.FileName;

                 string excelFileName = System.IO.Path.GetTempFileName();
                 Byte[] BOMinByteArray = System.IO.File.ReadAllBytes(openBOMFileDialog.FileName);
                 System.IO.File.WriteAllBytes(excelFileName, BOMinByteArray);

                 Excel.Application oXL;
                 Excel._Workbook oWB;
                 Excel._Worksheet oSheet;
                 Excel.Range oRngMarks = null;
                 Excel.Range oRngNotes = null;
                 try
                 {
                     //Start Excel and get Application object.
                     oXL = new Excel.Application();
                     oXL.Visible = false;

                     //Get a new workbook.

                     oWB = oXL.Workbooks.Open(excelFileName);
                     Excel.Sheets sheet = oWB.Worksheets;
                     
                    Excel.Worksheet worksheet = null;


                    List<int> joistWorksheetIndices = new List<int>();

                     for (int i = 1; i <= oWB.Sheets.Count; i++)
                     {
                         worksheet = (Excel.Worksheet)sheet.get_Item(i);
                         string worksheetName = worksheet.Name;
                         if (worksheetName.Contains("J") == true && worksheetName.Contains("(") == true)
                         {
                             joistWorksheetIndices.Add(i);
                         }
                     }


                     for (int i = 0; i < joistWorksheetIndices.Count; i++)
                     {
                         oSheet = (Excel._Worksheet)sheet.get_Item(joistWorksheetIndices[i]);
                         for (int k = 16; k < 46; k++)
                         {
                             oRngMarks = oSheet.get_Range("A" + k, Missing.Value);
                             oRngNotes = oSheet.get_Range("AA" + k, Missing.Value);

                             string stringoRngMarks = (string)oRngMarks.Text;
                             string stringoRngNotes = (string)oRngNotes.Text;
                             if (stringoRngMarks!= "" && stringoRngMarks != "MARK")
                             {
                                 BOMjoistMarks.Add(stringoRngMarks);
                                 BOMjoistNotes.Add(stringoRngNotes);
                             }
                         }
                     }

                     oWB.Close(0);
                     oXL.Quit();
                     Marshal.ReleaseComObject(oWB);
                     Marshal.ReleaseComObject(oXL);
                     Marshal.ReleaseComObject(sheet);
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


             BOMMarksAndNotes.Add(BOMjoistMarks);
             BOMMarksAndNotes.Add(BOMjoistNotes);

            
             return BOMMarksAndNotes;
         }



            

        

    }
}
