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
        static public Job JobFromShoporderJoistDetails()
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

                var headerRow =
                        Enumerable.Range(0, joistDetailArray.GetLength(1))
                            .Select(x => joistDetailArray[1, x+1])
                            .ToArray();

                var Mark_Index = Array.IndexOf(headerRow, "Mark") + 1;
                var Quantity_Index = Array.IndexOf(headerRow, "Quantity") + 1;
                var Description_Index = Array.IndexOf(headerRow, "Description") + 1;
                var Base_Length_Index = Array.IndexOf(headerRow, "Base Length") + 1;
                var Joist_Type_Index = Array.IndexOf(headerRow, "Joist Type") + 1;
                var Seats_BDL_Index = Array.IndexOf(headerRow, "Seats BDL") + 1;
                var Seats_BDR_Index = Array.IndexOf(headerRow, "Seats BDR") + 1;
                var TCXL_Index = Array.IndexOf(headerRow, "TCXL") + 1;
                var TCXR_Index = Array.IndexOf(headerRow, "TCXR") + 1;
                var BCXL_Index = Array.IndexOf(headerRow, "BCXL") + 1;
                var BCXR_Index = Array.IndexOf(headerRow, "BCXR") + 1;
                var Chords_Index = Array.IndexOf(headerRow, "Chords") + 1;
                var BCP_Index = Array.IndexOf(headerRow, "BCP") + 1;
                var Material_Index = Array.IndexOf(headerRow, "Material") + 1;
                var Weight_Index = Array.IndexOf(headerRow, "Weight") + 1;
                var Total_LH_Index = Array.IndexOf(headerRow, "Total LH") + 1;
                var Bridging_Index = Array.IndexOf(headerRow, "Bridging") + 1;
                var TL_Deflection_Index = Array.IndexOf(headerRow, "TL Deflection") + 1;
                var LL_Deflection_Index = Array.IndexOf(headerRow, "LL Deflection") + 1;
                var BL_Decimal_Index = Array.IndexOf(headerRow, "BL Decimal") + 1;
                var Time_Index = Array.IndexOf(headerRow, "Time") + 1;
                var UseWoodNailerTC_Index = Array.IndexOf(headerRow, "UseWoodNailerTC") + 1;
                var TCLength_Index = Array.IndexOf(headerRow, "TCLength") + 1;
                var BCLength_Index = Array.IndexOf(headerRow, "BCLength") + 1;
                var LineName_Index = Array.IndexOf(headerRow, "LineName") + 1;
                var TCMaxBridging_Index = Array.IndexOf(headerRow, "TCMaxBridging_ING") + 1;
                var BCMaxBridging_Index = Array.IndexOf(headerRow, "BCMaxBridging_ING") + 1;
                var ReactionLEmax_Index = Array.IndexOf(headerRow, "ReactionLEmax") + 1;
                var ReactionLEmin_Index = Array.IndexOf(headerRow, "ReactionLEmin") + 1;
                var ReactionREmax_Index = Array.IndexOf(headerRow, "ReactionREmax") + 1;
                var ReactionREmin_Index = Array.IndexOf(headerRow, "ReactionREmin") + 1;





                int joistDetailArrayRows = joistDetailArray.GetLength(0);

                for (int row = 2; row <= joistDetailArrayRows; row++)
                {
                    string joistDescription = (string)joistDetailArray[row, Description_Index];
                    if (joistDescription.Contains("G") == true)
                    {
                        Girder girder = new Girder();
                        girder.Mark = (string)joistDetailArray[row, Mark_Index];
                        girder.Quantity = Convert.ToInt32(joistDetailArray[row, Quantity_Index]);
                        girder.Description = (string)joistDetailArray[row, Description_Index];
                        girder.BaseLength = Convert.ToDouble(joistDetailArray[row, Base_Length_Index]);
                        girder.JoistType = (string)joistDetailArray[row, Joist_Type_Index];
                        girder.SeatsBDL = Convert.ToDouble(joistDetailArray[row, Seats_BDL_Index]);
                        girder.SeatsBDR = Convert.ToDouble(joistDetailArray[row, Seats_BDR_Index]);
                        girder.TCXL = Convert.ToDouble(joistDetailArray[row, TCXL_Index]);
                        girder.TCXR = Convert.ToDouble(joistDetailArray[row, TCXR_Index]);
                        girder.BCXL = Convert.ToDouble(joistDetailArray[row, BCXL_Index]);
                        girder.BCXR = Convert.ToDouble(joistDetailArray[row, BCXR_Index]);
                        string TCandBC = (string)joistDetailArray[row, Chords_Index];
                        girder.TC = TCandBC.Split('/')[0];
                        girder.BC = TCandBC.Split('/')[1];
                        girder.MaterialCost = Convert.ToDouble(joistDetailArray[row, Material_Index]);
                        girder.WeightInLBS = Convert.ToDouble(joistDetailArray[row, Weight_Index]);
                        girder.TotalLH = Convert.ToDouble(joistDetailArray[row, Total_LH_Index]);
                        girder.BLDecimal = Convert.ToDouble(joistDetailArray[row, BL_Decimal_Index]);
                        girder.Time = Convert.ToDouble(joistDetailArray[row, Time_Index]);
                        girder.UseWood = Convert.ToBoolean(joistDetailArray[row, UseWoodNailerTC_Index]);
                        girder.DecimalTcMaxBridgingSpacing = ((string)joistDetailArray[row, TCMaxBridging_Index]).Split(' ')[0] ;
                        girder.DecimalBcMaxBridgingSpacing = ((string)joistDetailArray[row, BCMaxBridging_Index]).Split(' ')[0];
                        allGirders.Add(girder);

                    }
                    else
                    {
                        Joist joist = new Joist();
                        joist.Mark = (string)joistDetailArray[row, Mark_Index];
                        joist.Quantity = Convert.ToInt32(joistDetailArray[row, Quantity_Index]);
                        joist.Description = (string)joistDetailArray[row, Description_Index];
                        joist.BaseLength = Convert.ToDouble(joistDetailArray[row, Base_Length_Index]);
                        joist.JoistType = (string)joistDetailArray[row, Joist_Type_Index];
                        joist.SeatsBDL = Convert.ToDouble(joistDetailArray[row, Seats_BDL_Index]);
                        joist.SeatsBDR = Convert.ToDouble(joistDetailArray[row, Seats_BDR_Index]);
                        joist.TCXL = Convert.ToDouble(joistDetailArray[row, TCXL_Index]);
                        joist.TCXR = Convert.ToDouble(joistDetailArray[row, TCXR_Index]);
                        joist.BCXL = Convert.ToDouble(joistDetailArray[row, BCXL_Index]);
                        joist.BCXR = Convert.ToDouble(joistDetailArray[row, BCXR_Index]);
                        string TCandBC = (string)joistDetailArray[row, Chords_Index];
                        joist.TC = TCandBC.Split('/')[0];
                        joist.BC = TCandBC.Split('/')[1];
                        joist.MaterialCost = Convert.ToDouble(joistDetailArray[row, Material_Index]);
                        joist.WeightInLBS = Convert.ToDouble(joistDetailArray[row, Weight_Index]);
                        joist.TotalLH = Convert.ToDouble(joistDetailArray[row, Total_LH_Index]);
                        joist.BLDecimal = Convert.ToDouble(joistDetailArray[row, BL_Decimal_Index]);
                        joist.Time = Convert.ToDouble(joistDetailArray[row, Time_Index]);
                        joist.UseWood = Convert.ToBoolean(joistDetailArray[row, UseWoodNailerTC_Index]);
                        joist.DecimalTcMaxBridgingSpacing = ((string)joistDetailArray[row, TCMaxBridging_Index]).Split(' ')[0];
                        joist.DecimalBcMaxBridgingSpacing = ((string)joistDetailArray[row, BCMaxBridging_Index]).Split(' ')[0];
                        allJoists.Add(joist);
                    }
                    
                    
                }
                allJoists = allJoists.OrderBy(x => x.StrippedNumber).ToList();
                allGirders = allGirders.OrderBy(x => x.StrippedNumber).ToList();
                job.Joists = allJoists;
                job.Girders = allGirders;

                oWB.Close();
                oXL.Quit();
            }
            return job;
        }
    }
}
