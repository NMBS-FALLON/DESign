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


namespace DESign_WordAddIn
{
    class ExcelDataExtraction
    {

        StringManipulation StringManipulation = new StringManipulation();

        public class NailerInformation
        {
            
            internal List<string> As { get; set; }
            internal List<string> Bs { get; set; }
            internal List<string> Marks { get; set; }
            private string pattern = "Staggered";
            internal string Pattern { get { return pattern; } set { pattern = value; } }
            internal string Initials { get; set; }
            internal List<string> Spacing { get; set; }
            internal List<string> WoodLengths { get; set; }
            
        }

        public NailerInformation exlNailerValues ()
        {
            NailerInformation nailerInformation = new NailerInformation();
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
                    oSheet = oWB.ActiveSheet;


                    Excel.Range last = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

                    stringJoistMarks = oSheet.get_Range("B6", "E" + last.Row).Value2;

                    nailerInformation.Initials = Convert.ToString(oSheet.Cells[2, 3].Value2);
                    nailerInformation.Pattern = Convert.ToString(oSheet.Cells[3, 3].Value2);


                    oWB.Close(0);
                    oXL.Quit();
                    Marshal.ReleaseComObject(oWB);
                    Marshal.ReleaseComObject(oXL);
                    Marshal.ReleaseComObject(oSheet);

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

          
             List<string> AsfromExcel = new List<string>();
             List<string> BsfromExcel = new List<string>();
             List<string> JoistMarksfromExcel = new List<string>();
            List<string> spacings = new List<string>();

                                      

             for (int i = 1; i <= stringJoistMarks.GetLength(0); i++)
             {
                 string stringA = StringManipulation.convertLengthStringtoHyphenLength(stringJoistMarks[i, 2].ToString());
                 string stringB = StringManipulation.convertLengthStringtoHyphenLength(stringJoistMarks[i, 3].ToString());
                string spacing = stringJoistMarks[i, 4].ToString();
                 
                 JoistMarksfromExcel.Add(stringJoistMarks[i,1].ToString());
                spacings.Add(spacing);
                 AsfromExcel.Add(stringA);
                 BsfromExcel.Add(stringB);
             }

            

            nailerInformation.Marks = JoistMarksfromExcel;
            nailerInformation.As = AsfromExcel;
            nailerInformation.Bs = BsfromExcel;
            nailerInformation.Spacing = spacings;



            return nailerInformation;


         
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
                 string excelFileName = openBOMFileDialog.FileName;

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
