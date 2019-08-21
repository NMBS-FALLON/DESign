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
using System.Diagnostics;
using DESign_BOT;

namespace DESign_BOT
{

    class ClassInsertBOMData
    {

        ClassExtractBOMData classExtractBOMData = new ClassExtractBOMData();
        public void createNMBSBOM()
        {


            List<List<object[]>> nucorBOMdata = new List<List<object[]>>();

            nucorBOMdata = classExtractBOMData.NucorBOMJoistInfo();

            List<object[]> nucorJoistData = nucorBOMdata[0];
            List<object[]> nucorGirderData = nucorBOMdata[1];

            OpenFileDialog openBOMFileDialog = new OpenFileDialog();

            openBOMFileDialog.Title = "SELECT EMPTY NMBS BOM";


            //    openBOMFileDialog.Filter = "|*.xlsx;*.xlsm";



            if (openBOMFileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelFileName = openBOMFileDialog.FileName;

                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;

                try
                {
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    oXL.Visible = true;

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


                    int numberOfSheetsNeeded = (nucorJoistData.Count() + 29) / 30;
                    int numberOfSheetsRemaining = numberOfSheetsNeeded;
                    int joistCounter = 0;


                    for (int pageCountIndex = 0; pageCountIndex < numberOfSheetsNeeded; pageCountIndex++)
                    {
                        oSheet = (Excel._Worksheet)sheet.get_Item(joistWorksheetIndices[pageCountIndex]);
                        if (numberOfSheetsRemaining - 1 > 0)
                        {
                            object[,] thisPageJoistData = new object[30, 26];
                            for (int joist_ROW = 0; joist_ROW < 30; joist_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 26; data_Column++)
                                {
                                    thisPageJoistData[joist_ROW, data_Column] = nucorJoistData[joistCounter][data_Column];
                                }
                                joistCounter++;

                            }
                            oSheet.get_Range("A16", "Z45").Value2 = thisPageJoistData;
                            numberOfSheetsRemaining = numberOfSheetsRemaining - 1;

                        }
                        else
                        {
                            int extraRowsNeeded = nucorJoistData.Count() - joistCounter;
                            object[,] thisPageJoistData = new object[extraRowsNeeded, 26];
                            for (int joist_ROW = 0; joist_ROW < extraRowsNeeded; joist_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 26; data_Column++)
                                {
                                    thisPageJoistData[joist_ROW, data_Column] = nucorJoistData[joistCounter][data_Column];
                                }
                                joistCounter++;

                            }
                            oSheet.get_Range("A16", "Z" + (15 + extraRowsNeeded)).Value2 = thisPageJoistData;
                        }

                    }





                    SaveFileDialog saveExcelNotesDialog = new SaveFileDialog();

                    saveExcelNotesDialog.Title = "File Location For New NMBS BOM";

                    saveExcelNotesDialog.Filter = "ExcelX|*.xlsx|ExcelMacro|*.xlsm";



                    if (saveExcelNotesDialog.ShowDialog() == DialogResult.OK)
                    {
                        string saveExcelFileName = saveExcelNotesDialog.FileName;
                        oWB.SaveAs(saveExcelFileName);
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


        }
        public void createNMBSBOM2()
        {

            List<List<object[]>> nucorBOMdata = new List<List<object[]>>();

            nucorBOMdata = classExtractBOMData.NucorBOMJoistInfo();

            List<object[]> nucorJoistData = nucorBOMdata[0];
            List<object[]> nucorGirderData = nucorBOMdata[1];

            OpenFileDialog openBOMFileDialog = new OpenFileDialog();

            openBOMFileDialog.Title = "SELECT EMPTY NMBS BOM";

            //    openBOMFileDialog.Filter = "|*.xlsx;*.xlsm";



            if (openBOMFileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelFileName = openBOMFileDialog.FileName;

                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;

                try
                {
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    oXL.Visible = true;

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



                    int numberOfSheetsNeeded = (nucorJoistData.Count() + 29) / 30;
                    int numberOfSheetsRemaining = numberOfSheetsNeeded;
                    int joistCounter = 0;


                    for (int pageCountIndex = 0; pageCountIndex < numberOfSheetsNeeded; pageCountIndex++)
                    {
                        oSheet = (Excel._Worksheet)sheet.get_Item(joistWorksheetIndices[pageCountIndex]);
                        if (numberOfSheetsRemaining - 1 > 0)
                        {
                            object[,] thisPageJoistData = new object[30, 26];
                            for (int joist_ROW = 0; joist_ROW < 30; joist_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 26; data_Column++)
                                {
                                    thisPageJoistData[joist_ROW, data_Column] = nucorJoistData[joistCounter][data_Column];
                                }
                                joistCounter++;

                            }
                            oSheet.get_Range("A16", "Z45").Value2 = thisPageJoistData;
                            numberOfSheetsRemaining = numberOfSheetsRemaining - 1;

                        }
                        else
                        {
                            int extraRowsNeeded = nucorJoistData.Count() - joistCounter;
                            object[,] thisPageJoistData = new object[extraRowsNeeded, 26];
                            for (int joist_ROW = 0; joist_ROW < extraRowsNeeded; joist_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 26; data_Column++)
                                {
                                    thisPageJoistData[joist_ROW, data_Column] = nucorJoistData[joistCounter][data_Column];
                                }
                                joistCounter++;

                            }
                            oSheet.get_Range("A16", "Z" + (15 + extraRowsNeeded)).Value2 = thisPageJoistData;
                        }

                    }

                    //

                    List<int> girderWorksheetIndices = new List<int>();

                    for (int i = 1; i <= oWB.Sheets.Count; i++)
                    {
                        worksheet = (Excel.Worksheet)sheet.get_Item(i);
                        string worksheetName = worksheet.Name;
                        if (worksheetName.Contains("G") == true && worksheetName.Contains("(") == true)
                        {
                            girderWorksheetIndices.Add(i);
                        }
                    }



                    int numberOfGirderSheetsNeeded = (nucorGirderData.Count() + 31) / 32;
                    int numberOfGirderSheetsRemaining = numberOfGirderSheetsNeeded;
                    int girderCounter = 0;


                    for (int pageCountIndex = 0; pageCountIndex < numberOfGirderSheetsNeeded; pageCountIndex++)
                    {
                        oSheet = (Excel._Worksheet)sheet.get_Item(girderWorksheetIndices[pageCountIndex]);
                        if (numberOfGirderSheetsRemaining - 1 > 0)
                        {
                            object[,] thisPageGirderData = new object[32, 36];
                            for (int girder_ROW = 0; girder_ROW < 32; girder_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 36; data_Column++)
                                {
                                    thisPageGirderData[girder_ROW, data_Column] = nucorGirderData[girderCounter][data_Column];
                                }
                                girderCounter++;

                            }
                            oSheet.get_Range("A14", "AJ45").Value2 = thisPageGirderData;
                            numberOfGirderSheetsRemaining = numberOfGirderSheetsRemaining - 1;

                        }
                        else
                        {
                            int extraRowsNeeded = nucorGirderData.Count() - girderCounter;
                            object[,] thisPageGirderData = new object[extraRowsNeeded, 36];
                            for (int girder_ROW = 0; girder_ROW < extraRowsNeeded; girder_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 36; data_Column++)
                                {
                                    thisPageGirderData[girder_ROW, data_Column] = nucorGirderData[girderCounter][data_Column];
                                }
                                girderCounter++;

                            }
                            oSheet.get_Range("A14", "AJ" + (13 + extraRowsNeeded)).Value2 = thisPageGirderData;
                        }

                    }




                    /*

                                        SaveFileDialog saveExcelNotesDialog = new SaveFileDialog();

                                        saveExcelNotesDialog.Title = "File Location For New NMBS BOM";

                                        saveExcelNotesDialog.Filter = "ExcelX|*.xlsx|ExcelMacro|*.xlsm";



                                        if (saveExcelNotesDialog.ShowDialog() == DialogResult.OK)
                                        {
                                            string saveExcelFileName = saveExcelNotesDialog.FileName;
                                            oWB.SaveAs(saveExcelFileName);
                                        }
                    */
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
        }
        /*       public void openExcelTemplateFromResources ()
               {
                   string excelTemplateFromResources = System.IO.Path.GetTempFileName(); 

                   System.IO.File.WriteAllBytes(excelTemplateFromResources,Properties.Resources.NMBS_BOM_EMPTY);

                   Excel.Application excelApplication = new Excel.Application();

                   Excel._Workbook excelWorkbook;

                   excelWorkbook = excelApplication.Workbooks.Open(excelTemplateFromResources);

                   excelApplication.Visible = true;

                   List<int> joistWorksheetIndices = new List<int>();
                   Excel.Worksheet worksheet = null;

                   for (int i = 1; i <= excelWorkbook.Sheets.Count; i++)
                   {
                       worksheet = (Excel.Worksheet)sheet.get_Item(i);
                       string worksheetName = worksheet.Name;
                       if (worksheetName.Contains("J") == true && worksheetName.Contains("(") == true)
                       {
                           joistWorksheetIndices.Add(i);
                       }
                   }
               }
         */

        public void createNMBSBOM3()
        {

            List<List<object[]>> nucorBOMdata = new List<List<object[]>>();

            nucorBOMdata = classExtractBOMData.NucorBOMJoistInfo();

            List<object[]> nucorJoistData = nucorBOMdata[0];
            List<object[]> nucorGirderData = nucorBOMdata[1];


//            if (nucorJoistData.Count() != 0 | nucorGirderData.Count() != 0)
//            {

                //      string excelFileName = Properties.Resources.NMBS_BOM_EMPTY;

                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;

                try
                {
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    oXL.Visible = false;

                    //Get a new workbook.

                    string excelPath = System.IO.Path.GetTempFileName();

                    System.IO.File.WriteAllBytes(excelPath, Properties.Resources.NMBS_BOM_EMPTY);

                    oWB = oXL.Workbooks.Open(excelPath);

                    Excel.Sheets sheet = oWB.Worksheets;

                    Excel.Worksheet worksheet = null;



                    int numberOfSheetsNeeded = (nucorJoistData.Count() + 29) / 30;
                    int numberOfSheetsRemaining = numberOfSheetsNeeded;
                    int joistCounter = 0;

                    for (int i = 0; i < numberOfSheetsNeeded - 1; i++)
                    {
                        worksheet = (Excel.Worksheet)sheet.get_Item(16);
                        string worksheetName = worksheet.Name;
                        worksheet.Copy(Type.Missing, After: (Excel.Worksheet)sheet.get_Item(16 + i));
                    }

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


                    for (int pageCountIndex = 0; pageCountIndex < numberOfSheetsNeeded; pageCountIndex++)
                    {
                        oSheet = (Excel._Worksheet)sheet.get_Item(joistWorksheetIndices[pageCountIndex]);
                        if (numberOfSheetsRemaining - 1 > 0)
                        {
                            object[,] thisPageJoistData = new object[30, 26];
                            for (int joist_ROW = 0; joist_ROW < 30; joist_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 26; data_Column++)
                                {
                                    thisPageJoistData[joist_ROW, data_Column] = nucorJoistData[joistCounter][data_Column];
                                }
                                joistCounter++;

                            }
                            oSheet.get_Range("A16", "Z45").Value2 = thisPageJoistData;
                            numberOfSheetsRemaining = numberOfSheetsRemaining - 1;

                        }
                        else
                        {
                            int extraRowsNeeded = nucorJoistData.Count() - joistCounter;
                            object[,] thisPageJoistData = new object[extraRowsNeeded, 26];
                            for (int joist_ROW = 0; joist_ROW < extraRowsNeeded; joist_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 26; data_Column++)
                                {
                                    thisPageJoistData[joist_ROW, data_Column] = nucorJoistData[joistCounter][data_Column];
                                }
                                joistCounter++;

                            }
                            oSheet.get_Range("A16", "Z" + (15 + extraRowsNeeded)).Value2 = thisPageJoistData;
                        }

                    }

                    //

                    List<int> girderWorksheetIndices = new List<int>();


                    int numberOfGirderSheetsNeeded = (nucorGirderData.Count() + 31) / 32;
                    int numberOfGirderSheetsRemaining = numberOfGirderSheetsNeeded;
                    int girderCounter = 0;

                    for (int i = 0; i < numberOfGirderSheetsNeeded - 1; i++)
                    {
                        worksheet = (Excel.Worksheet)sheet.get_Item(15);
                        string worksheetName = worksheet.Name;
                        worksheet.Copy(Type.Missing, After: (Excel.Worksheet)sheet.get_Item(15 + i));
                    }

                    for (int i = 1; i <= oWB.Sheets.Count; i++)
                    {
                        worksheet = (Excel.Worksheet)sheet.get_Item(i);
                        string worksheetName = worksheet.Name;
                        if (worksheetName.Contains("G") == true && worksheetName.Contains("(") == true)
                        {
                            girderWorksheetIndices.Add(i);
                        }
                    }



                    for (int pageCountIndex = 0; pageCountIndex < numberOfGirderSheetsNeeded; pageCountIndex++)
                    {
                        oSheet = (Excel._Worksheet)sheet.get_Item(girderWorksheetIndices[pageCountIndex]);
                        if (numberOfGirderSheetsRemaining - 1 > 0)
                        {
                            object[,] thisPageGirderData = new object[32, 36];
                            for (int girder_ROW = 0; girder_ROW < 32; girder_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 36; data_Column++)
                                {
                                    thisPageGirderData[girder_ROW, data_Column] = nucorGirderData[girderCounter][data_Column];
                                }
                                girderCounter++;

                            }
                            oSheet.get_Range("A14", "AJ45").Value2 = thisPageGirderData;
                            numberOfGirderSheetsRemaining = numberOfGirderSheetsRemaining - 1;

                        }
                        else
                        {
                            int extraRowsNeeded = nucorGirderData.Count() - girderCounter;
                            object[,] thisPageGirderData = new object[extraRowsNeeded, 36];
                            for (int girder_ROW = 0; girder_ROW < extraRowsNeeded; girder_ROW++)
                            {
                                for (int data_Column = 0; data_Column < 36; data_Column++)
                                {
                                    thisPageGirderData[girder_ROW, data_Column] = nucorGirderData[girderCounter][data_Column];
                                }
                                girderCounter++;

                            }
                            oSheet.get_Range("A14", "AJ" + (13 + extraRowsNeeded)).Value2 = thisPageGirderData;
                        }

                    }


                    oXL.Visible = true;

                    /*

                                        SaveFileDialog saveExcelNotesDialog = new SaveFileDialog();

                                        saveExcelNotesDialog.Title = "File Location For New NMBS BOM";

                                        saveExcelNotesDialog.Filter = "ExcelX|*.xlsx|ExcelMacro|*.xlsm";



                                        if (saveExcelNotesDialog.ShowDialog() == DialogResult.OK)
                                        {
                                            string saveExcelFileName = saveExcelNotesDialog.FileName;
                                            oWB.SaveAs(saveExcelFileName);
                                        }
                    */
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

        public void Blank_AB()
        {
             Excel.Application oXL;
             Excel._Workbook oWB;
             Excel._Worksheet oSheet;

     
             //Start Excel and get Application object.
             oXL = new Excel.Application();
             oXL.Visible = true;

             //Get a new workbook.

             string excelPath = System.IO.Path.GetTempFileName();

             System.IO.File.WriteAllBytes(excelPath, Properties.Resources.BLANK_AB_SHEET);

             oWB = oXL.Workbooks.Open(excelPath);
        }

        public void NucorBOM_AB()
        {

            List<List<object[]>> nucorBOMdata = new List<List<object[]>>();

            nucorBOMdata = classExtractBOMData.NucorBOMJoistInfo();

            if (nucorBOMdata[0].Count() > 0 || nucorBOMdata[1].Count > 0)
            {
                

                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;


                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.

                string excelPath = System.IO.Path.GetTempFileName();

                System.IO.File.WriteAllBytes(excelPath, Properties.Resources.BLANK_AB_SHEET);

                oWB = oXL.Workbooks.Open(excelPath);


                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                oRng = oSheet.get_Range("B5", Missing.Value);


                List<object[]> nucorJoistData = nucorBOMdata[0];
                List<object[]> nucorGirderData = nucorBOMdata[1];
                object[,] Marks = new object[nucorJoistData.Count() + nucorGirderData.Count(), 1];
                object[,] Nailer_A_Array = new object[nucorJoistData.Count() + nucorGirderData.Count(), 1];
                object[,] Nailer_B_Array = new object[nucorJoistData.Count() + nucorGirderData.Count(), 1];
                object[,] Nailer_Space_Array = new object[nucorJoistData.Count() + nucorGirderData.Count(), 1];
                object[,] HoldClear = new object[nucorJoistData.Count() + nucorGirderData.Count(), 1];

                for (int i = 0; i < nucorJoistData.Count(); i++)
                {
                    Marks[i, 0] = nucorJoistData[i][0];

                    if (nucorJoistData[i][26] == null | (nucorJoistData[i][26] ?? String.Empty).ToString() == "")
                    {
                        Nailer_A_Array[i, 0] = 0;
                    }
                    else
                    {
                        Nailer_A_Array[i, 0] = nucorJoistData[i][26];
                    }

                    if (nucorJoistData[i][27] == null | (nucorJoistData[i][27] ?? String.Empty).ToString() == "")
                    {
                        Nailer_B_Array[i, 0] = 0;
                    }
                    else
                    {
                        Nailer_B_Array[i, 0] = nucorJoistData[i][27];
                    }
                    if (nucorJoistData[i][28] == null | (nucorJoistData[i][28] ?? String.Empty).ToString() == "")
                    {
                        Nailer_Space_Array[i, 0] = 0;
                    }
                    else
                    {
                        Nailer_Space_Array[i, 0] = (object)(Convert.ToDouble(nucorJoistData[i][28])*2);
                    }
                    bool hcLeft = !(nucorJoistData[i][29] == null | (nucorJoistData[i][29] ?? String.Empty).ToString() == "");
                    bool hcRight = !(nucorJoistData[i][30] == null | (nucorJoistData[i][30] ?? String.Empty).ToString() == "");
                    if (hcLeft && hcRight)
                    {
                        HoldClear[i, 0] = "BOTH";
                    }
                    else if (hcLeft)
                    {
                        HoldClear[i, 0] = "LEFT";
                    }
                    else if (hcRight)
                    {
                        HoldClear[i, 0] = "RIGHT";
                    }
                    else
                    {
                        HoldClear[i, 0] = "";
                    }


                }

                int joistCount = nucorJoistData.Count();

                for (int i = 0; i < nucorGirderData.Count(); i++)
                {
                    int j = i + joistCount;
                    Marks[j, 0] = nucorGirderData[i][0];
                    Nailer_A_Array[j, 0] = "";
                    Nailer_B_Array[j, 0] = "";
                    Nailer_Space_Array[j, 0] = "";
                    
                    bool hcLeft = !(nucorGirderData[i][36] == null | (nucorGirderData[i][36] ?? String.Empty).ToString() == "");
                    bool hcRight = !(nucorGirderData[i][37] == null | (nucorGirderData[i][37] ?? String.Empty).ToString() == "");
                    if (hcLeft && hcRight)
                    {
                        HoldClear[j, 0] = "BOTH";
                    }
                    else if (hcLeft)
                    {
                        HoldClear[j, 0] = "LEFT";
                    }
                    else if (hcRight)
                    {
                        HoldClear[j, 0] = "RIGHT";
                    }
                    else
                    {
                        HoldClear[j, 0] = "";
                    }


                }

                int lastRow = 5 + Nailer_A_Array.Length;
                oSheet.get_Range("B7", "B" + lastRow).Value2 = Marks;
                oSheet.get_Range("C7", "C" + lastRow).Value2 = Nailer_A_Array;
                oSheet.get_Range("D7", "D" + lastRow).Value2 = Nailer_B_Array;
                oSheet.get_Range("E7", "E" + lastRow).Value2 = Nailer_Space_Array;
                oSheet.get_Range("F7", "F" + lastRow).Value2 = HoldClear;

                oXL.Visible = true;

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.ShowDialog();
                if (saveFileDialog.FileName != "")
                {
                    oWB.CheckCompatibility = false;
                    oWB.SaveAs(saveFileDialog.FileName);
                }


                



                /*           }
                           catch
                           {

                           }
               */
            }
        }

    }

}




