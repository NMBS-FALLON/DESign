using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;
using System.Data;

namespace DESign_BASE
{
    public class ExtractJoistDetails
    {
        public Job JobFromShoporderJoistDetails()
        {
            Job job = new Job();
            List<Joist> allJoists = new List<Joist>();
            List<Girder> allGirders = new List<Girder>();

            OpenFileDialog openBOMFileDialog = new OpenFileDialog();
            openBOMFileDialog.Title = "SELECT JOIST DETAILS";
            //openBOMFileDialog.Filter = "Excel WorkBook|*.xls";
            if (openBOMFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = openBOMFileDialog.FileName;
                job.Number = fileName.Split(new string[] { " -", ".xls" }, StringSplitOptions.RemoveEmptyEntries)[1];

                string excelFileName = System.IO.Path.GetTempFileName();
                Byte[] BOMinByteArray = System.IO.File.ReadAllBytes(openBOMFileDialog.FileName);
                System.IO.File.WriteAllBytes(excelFileName, BOMinByteArray);
                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;
                Excel.Range oRng;

                oXL = new Excel.Application();
                oXL.Visible = false;

                oWB = oXL.Workbooks.Open(excelFileName);
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                oRng = oSheet.UsedRange;

             //   oRng.get_Range("A1", Missing.Value);
             //   oRng = oRng.get_End(Excel.XlDirection.xlToRight);
             //   oRng = oRng.get_End(Excel.XlDirection.xlDown);
             //   string downJoistMarks = oRng.get_Address(Excel.XlReferenceStyle.xlA1, Type.Missing);
             //   oRng = oSheet.get_Range("A1", downJoistMarks);
                var joistDetailArray = (object[,])oRng.Value2;

                int joistDetailArrayRows = joistDetailArray.GetLength(0);

                for (int row = 2; row <= joistDetailArrayRows; row++)
                {
                    string joistDescription = (string)joistDetailArray[row, 1];
                    if (joistDescription.Contains("G") == true)
                    {
                        Girder girder = new Girder();
                        girder.Mark = (string)joistDetailArray[row, 1];
                        girder.Quantity = Convert.ToInt32(joistDetailArray[row, 2]);
                        girder.Description = (string)joistDetailArray[row, 3];
                        girder.BaseLength = Convert.ToDouble(joistDetailArray[row, 4]);
                        girder.JoistType = (string)joistDetailArray[row, 5];
                        girder.SeatsBDL = Convert.ToDouble(joistDetailArray[row, 6]);
                        girder.SeatsBDR = Convert.ToDouble(joistDetailArray[row, 7]);
                        girder.TCXL = Convert.ToDouble(joistDetailArray[row, 8]);
                        girder.TCXR = Convert.ToDouble(joistDetailArray[row, 9]);
                        girder.BCXL = Convert.ToDouble(joistDetailArray[row, 10]);
                        girder.BCXR = Convert.ToDouble(joistDetailArray[row, 11]);
                        string TCandBC = (string)joistDetailArray[row, 12];
                        girder.TC = TCandBC.Split('/')[0];
                        girder.BC = TCandBC.Split('/')[1];
                        girder.MaterialCost = Convert.ToDouble(joistDetailArray[row, 13]);
                        girder.WeightInLBS = Convert.ToDouble(joistDetailArray[row, 14]);
                        girder.TotalLH = Convert.ToDouble(joistDetailArray[row, 15]);
                        girder.BLDecimal = Convert.ToDouble(joistDetailArray[row, 19]);
                        girder.Time = Convert.ToDouble(joistDetailArray[row, 20]);
                        girder.UseWood = Convert.ToBoolean(joistDetailArray[row, 21]);
                        allGirders.Add(girder);

                    }
                    else
                    {
                        Joist joist = new Joist();
                        joist.Mark = (string)joistDetailArray[row, 1];
                        joist.Quantity = Convert.ToInt32(joistDetailArray[row, 2]);
                        joist.Description = (string)joistDetailArray[row, 3];
                        joist.BaseLength = Convert.ToDouble(joistDetailArray[row, 4]);
                        joist.JoistType = (string)joistDetailArray[row, 5];
                        joist.SeatsBDL = Convert.ToDouble(joistDetailArray[row, 6]);
                        joist.SeatsBDR = Convert.ToDouble(joistDetailArray[row, 7]);
                        joist.TCXL = Convert.ToDouble(joistDetailArray[row, 8]);
                        joist.TCXR = Convert.ToDouble(joistDetailArray[row, 9]);
                        joist.BCXL = Convert.ToDouble(joistDetailArray[row, 10]);
                        joist.BCXR = Convert.ToDouble(joistDetailArray[row, 11]);
                        string TCandBC = (string)joistDetailArray[row, 12];
                        joist.TC = TCandBC.Split('/')[0];
                        joist.BC = TCandBC.Split('/')[1];
                        joist.MaterialCost = Convert.ToDouble(joistDetailArray[row, 13]);
                        joist.WeightInLBS = Convert.ToDouble(joistDetailArray[row, 14]);
                        joist.TotalLH = Convert.ToDouble(joistDetailArray[row, 15]);
                        joist.BLDecimal = Convert.ToDouble(joistDetailArray[row, 19]);
                        joist.Time = Convert.ToDouble(joistDetailArray[row, 20]);
                        joist.UseWood = Convert.ToBoolean(joistDetailArray[row, 21]);
                        allJoists.Add(joist);
                    }
                    
                    
                }
                allJoists = allJoists.OrderBy(x => x.StrippedNumber).ToList();
                allGirders = allGirders.OrderBy(x => x.StrippedNumber).ToList();
                job.Joists = allJoists;
                job.Girders = allGirders;
            }
            return job;
        }
    }
}
