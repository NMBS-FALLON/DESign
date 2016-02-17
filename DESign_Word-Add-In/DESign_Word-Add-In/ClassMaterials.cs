using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel=Microsoft.Office.Interop.Excel;
using Word=Microsoft.Office.Interop.Word;


using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;



namespace DESign_WordAddIn
{
    class ClassMaterials
    {
        StringManipulation StringManipulation = new StringManipulation();
        public List<Tuple<string, double, double, double>> angleData()
        {
            var angleData = new List<Tuple<string, double, double, double>>
               {
            Tuple.Create("1010",1.0,1.0,0.109),
            Tuple.Create("1011",1.0,1.0,0.115),
            Tuple.Create("1210",1.25,1.25,0.109),
            Tuple.Create("1212",1.25,1.25,0.125),
            Tuple.Create("1510",1.5,1.5,0.109),
            Tuple.Create("1511",1.5,1.5,0.115),
            Tuple.Create("1512",1.5,1.5,0.123),
            Tuple.Create("15130",1.5,1.5,0.13),
            Tuple.Create("1513",1.5,1.5,0.137),
            Tuple.Create("1515",1.5,1.5,0.155),
            Tuple.Create("1714",1.75,1.75,0.143),
            Tuple.Create("1715",1.75,1.75,0.155),
            Tuple.Create("2012",2.0,2.0,0.125),
            Tuple.Create("2015",2.0,2.0,0.156),
            Tuple.Create("2016",2.0,2.0,0.163),
            Tuple.Create("2018",2.0,2.0,0.188),
            Tuple.Create("2021",2.0,2.0,0.216),
            Tuple.Create("2024",2.0,2.0,0.248),
            Tuple.Create("203025",2.0,3.0,0.25),
            Tuple.Create("2521",2.5,2.5,0.212),
            Tuple.Create("2523",2.5,2.5,0.23),
            Tuple.Create("2525",2.5,2.5,0.25),
            Tuple.Create("302025",3.0,2.0,0.25),
            Tuple.Create("3022",3.0,3.0,0.227),
            Tuple.Create("3025",3.0,3.0,0.25),
            Tuple.Create("3028",3.0,3.0,0.281),
            Tuple.Create("3031",3.0,3.0,0.313),
            Tuple.Create("3528",3.5,3.5,0.287),
            Tuple.Create("3531",3.5,3.5,0.313),
            Tuple.Create("3534",3.5,3.5,0.344),
            Tuple.Create("4037",4.0,4.0,0.375),
            Tuple.Create("4043",4.0,4.0,0.438),
            Tuple.Create("4050",4.0,4.0,0.5),
            Tuple.Create("503550",5.0,3.5,0.5),
            Tuple.Create("5043",5.0,5.0,0.438),
            Tuple.Create("5050",5.0,5.0,0.5),
            Tuple.Create("407050",4.0,7.0,0.5),
            Tuple.Create("6050",6.0,6.0,0.5),
            Tuple.Create("6056",6.0,6.0,0.563),
            Tuple.Create("6062",6.0,6.0,0.625),
            Tuple.Create("6075",6.0,6.0,0.75),
            Tuple.Create("8050",8.0,8.0,0.5),
            Tuple.Create("8062",8.0,8.0,0.625),
            Tuple.Create("6010",6.0,6.0,1.0),
            Tuple.Create("8075",8.0,8.0,0.75),
            Tuple.Create("8010",8.0,8.0,1.0),
            Tuple.Create("15172",1.5,1.5,0.172),
            Tuple.Create("406037",4.0,6.0,0.375),
            Tuple.Create("1012",1.0,1.0,0.125),
            Tuple.Create("1517",1.5,1.5,0.17),
            Tuple.Create("1518",1.5,1.5,0.188),
            Tuple.Create("1717",1.75,1.75,0.17),
            Tuple.Create("A12B",1.156,1.156,0.09),
            Tuple.Create("A14B",1.109,1.109,0.102),
            Tuple.Create("A16B",1.375,1.375,0.102),
            Tuple.Create("A18B",1.375,1.375,0.118),
            Tuple.Create("A20B",1.437,1.437,0.124),
            Tuple.Create("A22B",1.5,1.5,0.129),
            Tuple.Create("A24B",1.594,1.594,0.133),
            Tuple.Create("A26B",1.656,1.656,0.142),
            Tuple.Create("A28B",1.735,1.735,0.15),
            Tuple.Create("A30B",1.797,1.797,0.158),
            Tuple.Create("A32B",1.906,1.906,0.158),
            Tuple.Create("A34A",1.938,1.938,0.176),
            Tuple.Create("A36B",2.078,2.078,0.188),
            Tuple.Create("A38B",2.219,2.219,0.199),
            Tuple.Create("A40B",2.375,2.375,0.218),
            Tuple.Create("A42A",2.625,2.625,0.209),
            Tuple.Create("A44A",2.875,2.875,0.209),
            Tuple.Create("A46A",2.5938,2.5938,0.25),
            Tuple.Create("A48A",3.0625,3.0625,0.227),
            Tuple.Create("A50A",3.0938,3.0938,0.25),
            Tuple.Create("A26B19",1.9375,1.3745,0.142),
            Tuple.Create("A28B19",1.9375,1.5325,0.15),
            Tuple.Create("A30B19",1.9375,1.6565,0.158),
            Tuple.Create("A36B19",1.9375,2.2185,0.188),
            Tuple.Create("A38B19",1.9375,2.5005,0.199),
            Tuple.Create("A40B19",1.9375,2.8165,0.218),
            Tuple.Create("A40B30",3.0,1.754,0.218),
            Tuple.Create("A42A30",3.0,2.25,0.209),
            Tuple.Create("A44A30",3.0,2.75,0.209),
            Tuple.Create("A46A30",3.0,2.1876,0.25),
            Tuple.Create("A48A30",3.0,3.125,0.227),
            Tuple.Create("A50A30",3.0,3.1876,0.25),
            Tuple.Create("A26B18",1.875,1.437,0.142),
            Tuple.Create("A28B18",1.875,1.595,0.15),
            Tuple.Create("A30B18",1.875,1.719,0.158),
            Tuple.Create("A34A18",1.875,2.001,0.176),
            Tuple.Create("A36B18",1.875,2.281,0.188),
            Tuple.Create("A38B18",1.875,2.563,0.199),
            Tuple.Create("A40B18",1.875,2.879,0.218),
            Tuple.Create("A42A29",2.9375,2.3125,0.209),
            Tuple.Create("A44A29",2.9375,2.8125,0.209),
            Tuple.Create("A46A29",2.9375,2.2501,0.25),
            Tuple.Create("A48A29",2.9375,3.1875,0.227),


               };

            return angleData;

        }
        public void getChordAnglesfromExcel()
        {
            object[,] stringAngleData = null;

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


                    oRng.get_Range("A1", Missing.Value);
                    oRng = oRng.get_End(Excel.XlDirection.xlToRight);
                    oRng = oRng.get_End(Excel.XlDirection.xlDown);
                    string downJoistMarks = oRng.get_Address(Excel.XlReferenceStyle.xlA1, Type.Missing);
                    oRng = oSheet.get_Range("A1", downJoistMarks);
                    stringAngleData = (object[,])oRng.Value2;

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

            List<string> angleNames = new List<string>();
            List<string> angleHorizontalLeg = new List<string>();
            List<string> angleVerticalLeg = new List<string>();
            List<string> angleThickness = new List<string>();

            for (int i = 2; i <= stringAngleData.GetLength(0) - 1; i++)
            {
                angleNames.Add(stringAngleData[i, 1].ToString());

                string stringHorLegLength = null;
                string stringVertLegLength = null;
                string stringThickness = null;

                string angleDimensions = stringAngleData[i, 2].ToString();
                string[] angleDimensionsArray = angleDimensions.Split(new string[] { " x ", " X " }, StringSplitOptions.RemoveEmptyEntries);
                if (angleDimensionsArray.Count() == 2)
                {
                    stringHorLegLength = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                    stringVertLegLength = (12 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                    stringThickness = (12 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[1])).ToString();

                }
                else
                {
                    stringHorLegLength = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                    stringVertLegLength = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[1])).ToString();
                    stringThickness = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[2])).ToString();
                }

                if (stringHorLegLength.Contains(".") == false) { angleHorizontalLeg.Add(stringHorLegLength + ".0"); }
                else { angleHorizontalLeg.Add(stringHorLegLength); }

                if (stringVertLegLength.Contains(".") == false) { angleVerticalLeg.Add(stringVertLegLength + ".0"); }
                else { angleVerticalLeg.Add(stringVertLegLength); }

                if (stringThickness.Contains(".") == false) { angleThickness.Add(stringThickness + ".0"); }
                else { angleThickness.Add(stringThickness); }


            }

            Word.Application wordApp = new Word.Application();

            wordApp.Visible = true;

            Word.Document wordDoc = wordApp.Documents.Add();

            Word.Selection wordSelection = wordApp.Selection;

            wordSelection.HomeKey(Word.WdUnits.wdStory, 0);

            for (int i = 0; i < angleNames.Count(); i++)
            {
                wordSelection.Text = string.Format("Tuple.Create(\"{0}\",{1},{2},{3}),\r\n", angleNames[i], angleHorizontalLeg[i], angleVerticalLeg[i], angleThickness[i]);
                wordSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }



        }

        public void NMBSAngles()
        {
            StringManipulation StringManipulation = new StringManipulation();

            object[,] stringAngleData = null;

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


                    oRng.get_Range("A1", Missing.Value);
                    oRng = oRng.get_End(Excel.XlDirection.xlToRight);
                    oRng = oRng.get_End(Excel.XlDirection.xlDown);
                    string downJoistMarks = oRng.get_Address(Excel.XlReferenceStyle.xlA1, Type.Missing);
                    oRng = oSheet.get_Range("A1", downJoistMarks);
                    stringAngleData = (object[,])oRng.Value2;

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

            List<Angle> NMbsAngles = new List<Angle>();



            for (int i = 2; i <= stringAngleData.GetLength(0) - 1; i++)
            {
                Angle angle = new Angle();

                
                angle.name = stringAngleData[i, 1].ToString();

                if (angle.name.Contains("A") == false)
                {
                    string angleDimensions = stringAngleData[i, 2].ToString();
                    string[] angleDimensionsArray = angleDimensions.Split(new string[] { " x ", " X " }, StringSplitOptions.RemoveEmptyEntries);
                    if (angleDimensionsArray.Count() == 2)
                    {
                        angle.horizontalLeg = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                        angle.verticalLeg = (12 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                        angle.thickness = (12 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[1])).ToString();

                    }
                    else
                    {
                        angle.horizontalLeg = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                        angle.verticalLeg = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[1])).ToString();
                        angle.thickness = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[2])).ToString();
                    }

                    if (angle.horizontalLeg.Contains(".") == false) { angle.horizontalLeg = angle.horizontalLeg + ".0"; }
                    else { }

                    if (angle.verticalLeg.Contains(".") == false) { angle.verticalLeg = angle.verticalLeg + ".0"; }
                    else { }

                    if (angle.thickness.Contains(".") == false) { angle.thickness = angle.thickness + ".0"; }
                    else { }

                    angle.radius = "0";
                    angle.area = stringAngleData[i, 3].ToString();
                    angle.weight = stringAngleData[i, 4].ToString();
                    angle.rx = stringAngleData[i, 5].ToString();
                    angle.rz = stringAngleData[i, 6].ToString();
                    angle.y = stringAngleData[i, 7].ToString();
                    angle.x = stringAngleData[i, 8].ToString();
                    angle.lx = stringAngleData[i, 9].ToString();
                    angle.ly = stringAngleData[i, 10].ToString();
                    angle.Q = stringAngleData[i, 11].ToString();

                    NMbsAngles.Add(angle);
                }
                else
                {
                    string angleDimensions = stringAngleData[i, 2].ToString();
                    string[] angleDimensionsArray = angleDimensions.Split(new string[] { " x ", " X " }, StringSplitOptions.RemoveEmptyEntries);
                    if (angleDimensionsArray.Count() == 2)
                    {
                        angle.horizontalLeg = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                        angle.verticalLeg = (12 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                        angle.thickness = (12 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[1])).ToString();

                    }
                    else
                    {
                        angle.horizontalLeg = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                        angle.verticalLeg = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[1])).ToString();
                        angle.thickness = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[2])).ToString();
                    }

                    if (angle.horizontalLeg.Contains(".") == false) { angle.horizontalLeg = angle.horizontalLeg + ".0"; }
                    else { }

                    if (angle.verticalLeg.Contains(".") == false) { angle.verticalLeg = angle.verticalLeg + ".0"; }
                    else { }

                    if (angle.thickness.Contains(".") == false) { angle.thickness = angle.thickness + ".0"; }
                    else { }

                    angle.radius = stringAngleData[i, 3].ToString();
                    angle.area = stringAngleData[i, 4].ToString();
                    angle.weight = stringAngleData[i, 5].ToString();
                    angle.rx = stringAngleData[i, 6].ToString();
                    angle.rz = stringAngleData[i, 7].ToString();
                    angle.y = stringAngleData[i, 8].ToString();
                    angle.x = stringAngleData[i, 9].ToString();
                    angle.lx = stringAngleData[i, 10].ToString();
                    angle.ly = stringAngleData[i, 11].ToString();
                    angle.Q = stringAngleData[i, 12].ToString();

                    NMbsAngles.Add(angle);
                }
            }
            //  return NMBSangles;

            Word.Application wordApp = new Word.Application();

            wordApp.Visible = true;

            Word.Document wordDoc = wordApp.Documents.Add();

            Word.Selection wordSelection = wordApp.Selection;

            wordSelection.HomeKey(Word.WdUnits.wdStory, 0);

            for (int i = 0; i < NMbsAngles.Count(); i++)
            {

                wordSelection.Text = string.Format("Tuple.Create(\"{0}\",{1},{2},{3}),\r\n", NMbsAngles[i].name, NMbsAngles[i].horizontalLeg, NMbsAngles[i].verticalLeg, NMbsAngles[i].thickness);
                wordSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }



        }
        public void getChordAnglesfromExcel2()
        {
            object[,] stringAngleData = null;

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


                    oRng.get_Range("A1", Missing.Value);
                    oRng = oRng.get_End(Excel.XlDirection.xlToRight);
                    oRng = oRng.get_End(Excel.XlDirection.xlDown);
                    string downJoistMarks = oRng.get_Address(Excel.XlReferenceStyle.xlA1, Type.Missing);
                    oRng = oSheet.get_Range("A1", downJoistMarks);
                    stringAngleData = (object[,])oRng.Value2;

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

            List<string> angleNames = new List<string>();
            List<string> angleHorizontalLeg = new List<string>();
            List<string> angleVerticalLeg = new List<string>();
            List<string> angleThickness = new List<string>();
            List<string> angleArea = new List<string>();
            List<string> angleWeight = new List<string>();
            List<string> angleRx = new List<string>();
            List<string> angleRz = new List<string>();
            List<string> angleY = new List<string>();
            List<string> angleX = new List<string>();
            List<string> angleLx = new List<string>();
            List<string> angleLy = new List<string>();
            List<string> angleQ = new List<string>();


            for (int i = 2; i <= stringAngleData.GetLength(0) - 1; i++)
            {
                angleNames.Add(stringAngleData[i, 1].ToString());

                string stringHorLegLength = null;
                string stringVertLegLength = null;
                string stringThickness = null;

                string angleDimensions = stringAngleData[i, 2].ToString();
                string[] angleDimensionsArray = angleDimensions.Split(new string[] { " x ", " X " }, StringSplitOptions.RemoveEmptyEntries);
                if (angleDimensionsArray.Count() == 2)
                {
                    stringHorLegLength = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                    stringVertLegLength = (12 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                    stringThickness = (12 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[1])).ToString();

                }
                else
                {
                    stringHorLegLength = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[0])).ToString();
                    stringVertLegLength = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[1])).ToString();
                    stringThickness = (12.0 * StringManipulation.ConvertLengthtoDecimal(angleDimensionsArray[2])).ToString();
                }

                if (stringHorLegLength.Contains(".") == false) { angleHorizontalLeg.Add(stringHorLegLength + ".0"); }
                else { angleHorizontalLeg.Add(stringHorLegLength); }

                if (stringVertLegLength.Contains(".") == false) { angleVerticalLeg.Add(stringVertLegLength + ".0"); }
                else { angleVerticalLeg.Add(stringVertLegLength); }

                if (stringThickness.Contains(".") == false) { angleThickness.Add(stringThickness + ".0"); }
                else { angleThickness.Add(stringThickness); }


            }

            Word.Application wordApp = new Word.Application();

            wordApp.Visible = true;

            Word.Document wordDoc = wordApp.Documents.Add();

            Word.Selection wordSelection = wordApp.Selection;

            wordSelection.HomeKey(Word.WdUnits.wdStory, 0);

            for (int i = 0; i < angleNames.Count(); i++)
            {
                wordSelection.Text = string.Format("Tuple.Create(\"{0}\",{1},{2},{3}),\r\n", angleNames[i], angleHorizontalLeg[i], angleVerticalLeg[i], angleThickness[i]);
                wordSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }



        }

    }

    public class Angle
    {
        public string name;
        public string horizontalLeg;
        public string verticalLeg;
        public string thickness;
        public string radius;
        public string area;
        public string weight;
        public string rx;
        public string rz;
        public string y;
        public string x;
        public string lx;
        public string ly;
        public string Q;

       // public List<Angle> NMBSAngles()



    }




}
