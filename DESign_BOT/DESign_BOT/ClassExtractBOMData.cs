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
    class ClassExtractBOMData
    {

        StringManipulation stringManipulation = new StringManipulation();

        public List<List<object[]>> NucorBOMJoistInfo()
        {
            
            OpenFileDialog openBOMFileDialog = new OpenFileDialog();

            openBOMFileDialog.Title = "SELECT NUCOR BOM";
            openBOMFileDialog.Multiselect = true;

     //       openBOMFileDialog.Filter = "Excel WorkBook|*.xlsx;Excel Macro-Workbook|*.xlsm;97-2003 Workboook|*.xls;|*.xml;|*.xltx";

            List<List<object[]>> NucorBOMCompleteInfo = new List<List<object[]>>();

            if (openBOMFileDialog.ShowDialog() == DialogResult.OK)
            {

                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;


                oXL = new Excel.Application();
                oXL.Visible = false;


                List<object[]> nucorJoistData = new List<object[]>();

                List<object[]> nucorGirderData = new List<object[]>();


                foreach (String file in openBOMFileDialog.FileNames)
                {

                    string excelFileName = System.IO.Path.GetTempFileName();

                    Byte[] BOMinByteArray = System.IO.File.ReadAllBytes(file);

                    System.IO.File.WriteAllBytes(excelFileName, BOMinByteArray);

                    oWB = oXL.Workbooks.Open(excelFileName);


                    Excel.Sheets sheet = oWB.Worksheets;

                    Excel.Worksheet worksheet = null;

                    int visibleSheetCount = 0;

                    


                    List<int> joistWorksheetIndices = new List<int>();

                    for (int i = 1; i <= oWB.Sheets.Count; i++)
                    {
                        try
                        {
                            worksheet = (Excel.Worksheet)sheet.get_Item(i);
                            object workSheetTitle = worksheet.get_Range("AS1", Missing.Value).Value2;
                            string workSheetTitleString = (workSheetTitle ?? String.Empty).ToString();
                            if (workSheetTitleString.ToUpper().Contains("JOIST") && workSheetTitleString.Contains("BILL") == true)
                            {
                                joistWorksheetIndices.Add(i);
                            }
                        }
                        catch { }
                        
                    }


                    List<object[,]> listJoistSheetMultiArray = new List<object[,]>();
                    object[,] joistSheetMultiArray = null;
                    for (int i = 0; i < joistWorksheetIndices.Count; i++)
                    {
                        object mark, quantity, designation, overallLengthFeet, overallLengthInches, TCXLfeet, TCXLinches, TCXLtype, TCXRfeet, TCXRinches,
                               TCXRtype, BCXLfeet, BCXLinches, BCXLtype, BCXRfeet, BCXRinches, BCXRtype, bearingDepthL, bearingDepthR, baySlope, baySlopeHE, LEseatSlope, LEseatHL,
                               REseatSlope, REseatHL, holeLeftFeet, holeLeftInches, holeRightFeet, holeRightInches, holeSize, holeGage, NailerHoldBackLE, NailerHoldBackRE, NailerSpacing,
                               KnifePlateLE, KnifePlateRE, Adload, NetUplift, LeAxialW, LeAxialE, LeAxialEm, ReAxialW, ReAxialE, ReAxialEm;
                        mark = quantity = designation = overallLengthFeet = overallLengthInches = TCXLfeet = TCXLinches = TCXLtype = TCXRfeet = TCXRinches =
                               TCXRtype = BCXLfeet = BCXLinches = BCXLtype = BCXRfeet = BCXRinches = BCXRtype = bearingDepthL = bearingDepthR = baySlope = baySlopeHE = LEseatSlope = LEseatHL =
                               REseatSlope = REseatHL = holeLeftFeet = holeLeftInches = holeRightFeet = holeRightInches = holeSize = holeGage = NailerHoldBackLE = NailerHoldBackRE = NailerSpacing =
                               KnifePlateLE = KnifePlateRE = Adload = NetUplift = LeAxialW = LeAxialE =LeAxialEm = ReAxialW = ReAxialE = ReAxialEm =null;

                        oSheet = (Excel._Worksheet)sheet.get_Item(joistWorksheetIndices[i]);

                        joistSheetMultiArray = (object[,])oSheet.get_Range("A12", "GI124").Value2;

                        for (int j = 1; j <= 9; j = j + 2)
                        {
                            if (joistSheetMultiArray[j, 1] != null)
                            {
                                mark = joistSheetMultiArray[j, 1];
                                quantity = joistSheetMultiArray[j, 9];

                                if (joistSheetMultiArray[j, 15] != null && joistSheetMultiArray[j, 24].ToString() != null && joistSheetMultiArray[j, 29].ToString() != null)
                                {
                                    string stringDesignation = joistSheetMultiArray[j, 15].ToString() + joistSheetMultiArray[j, 24].ToString() + joistSheetMultiArray[j, 29].ToString();
                                    designation = stringDesignation;
                                }


                                overallLengthFeet = joistSheetMultiArray[j, 47];
                                overallLengthInches = joistSheetMultiArray[j, 53];
                                TCXLfeet = joistSheetMultiArray[j, 60];
                                TCXLinches = joistSheetMultiArray[j, 66];
                                TCXLtype = joistSheetMultiArray[j, 73];
                                TCXRfeet = joistSheetMultiArray[j, 76];
                                TCXRinches = joistSheetMultiArray[j, 82];
                                TCXRtype = joistSheetMultiArray[j, 89];
                                BCXLfeet = joistSheetMultiArray[j, 109];
                                BCXLinches = joistSheetMultiArray[j, 114];
                                BCXLtype = joistSheetMultiArray[j, 106];
                                BCXRfeet = joistSheetMultiArray[j, 124];
                                BCXRinches = joistSheetMultiArray[j, 129];
                                BCXRtype = joistSheetMultiArray[j, 121];
                                bearingDepthL = joistSheetMultiArray[j, 137];
                                bearingDepthR = joistSheetMultiArray[j, 144];
                                baySlope = joistSheetMultiArray[j, 151];
                                baySlopeHE = joistSheetMultiArray[j, 161];
                                LEseatSlope = joistSheetMultiArray[j, 164];
                                LEseatHL = joistSheetMultiArray[j, 174];
                                REseatSlope = joistSheetMultiArray[j, 177];
                                REseatHL = joistSheetMultiArray[j, 187];
                                holeLeftFeet = joistSheetMultiArray[j + 13, 43];
                                holeLeftInches = joistSheetMultiArray[j + 13, 48];
                                holeRightFeet = joistSheetMultiArray[j + 13, 80];
                                holeRightInches = joistSheetMultiArray[j + 13, 85];
                                holeSize = joistSheetMultiArray[j + 13, 57];
                                holeGage = joistSheetMultiArray[j + 13, 70];

                                KnifePlateLE = joistSheetMultiArray[j + 102, 159];
                                KnifePlateRE = joistSheetMultiArray[j + 102, 164];



                                NailerHoldBackLE = joistSheetMultiArray[j + 102, 169];
                                NailerHoldBackRE = joistSheetMultiArray[j + 102, 174];
                                NailerSpacing = joistSheetMultiArray[j + 102, 179];

                                NetUplift = joistSheetMultiArray[j + 26, 9];

                                Adload = joistSheetMultiArray[j + 26, 66];

                                LeAxialW = joistSheetMultiArray[j + 102, 59];
                                LeAxialE = joistSheetMultiArray[j + 102, 67];
                                LeAxialEm = joistSheetMultiArray[j + 102, 75];

                                ReAxialW = joistSheetMultiArray[j + 102, 88];
                                ReAxialE = joistSheetMultiArray[j + 102, 96];
                                ReAxialEm = joistSheetMultiArray[j + 102, 104];


                                // hole size is actually L= [i+13,57] +"X" + [i+13,64] w/ gage = [i+13,70]
                                //                       R= [i+13,94] +"X" + [i+13,101] w/ gage = [i+13,107]

                                //
                                if ((TCXLtype ?? String.Empty).ToString() == "F") { TCXLtype = "R"; }

                                if ((TCXRtype ?? String.Empty).ToString() == "F") { TCXRtype = "R"; }

                                if ((baySlope ?? String.Empty).ToString() != "")
                                {
                                    if ((baySlopeHE ?? String.Empty).ToString() == "L") { LEseatSlope = "-" + baySlope; REseatSlope = baySlope; }
                                    else { LEseatSlope = baySlope; REseatSlope = "-" + baySlope; }
                                }
                                else
                                {
                                    if ((LEseatSlope == null | (LEseatSlope ?? String.Empty).ToString() == "") && (REseatSlope == null | (REseatSlope ?? String.Empty).ToString() == ""))
                                    {
                                        LEseatSlope = REseatSlope = null;

                                    }
                                    else if ((LEseatSlope == null | (LEseatSlope ?? String.Empty).ToString() == ""))
                                    {
                                        if ((REseatHL ?? String.Empty).ToString() == "H") { REseatSlope = "-" + REseatSlope; }
                                        else { }
                                    }
                                    else if ((REseatSlope == null | (REseatSlope ?? String.Empty).ToString() == ""))
                                    {
                                        if ((LEseatHL ?? String.Empty).ToString() == "H") { LEseatSlope = "-" + LEseatSlope; }
                                        else { }
                                    }
                                    else
                                    {
                                        if ((REseatHL ?? String.Empty).ToString() == "H") { REseatSlope = "-" + REseatSlope; }
                                        else { }
                                        if ((LEseatHL ?? String.Empty).ToString() == "H") { LEseatSlope = "-" + LEseatSlope; }
                                        else { }


                                    }
                                }


                                object[] nucorJoistDataLine = new object[26];
                                nucorJoistDataLine = new object[] {mark, quantity, designation, overallLengthFeet, overallLengthInches,TCXLfeet,TCXLinches,TCXLtype,
                                                            TCXRfeet,TCXRinches,TCXRtype, bearingDepthL, bearingDepthR, BCXLfeet,BCXLinches,BCXLtype,BCXRfeet,BCXRinches,BCXRtype,
                                                            holeLeftFeet,holeLeftInches,holeRightFeet,holeRightInches,holeGage,
                                                           LEseatSlope,REseatSlope, NailerHoldBackLE, NailerHoldBackRE, NailerSpacing, KnifePlateLE, KnifePlateRE};

                                nucorJoistData.Add(nucorJoistDataLine);

                            }




                        }
                    }

                    List<int> girderWorksheetIndices = new List<int>();

                    for (int i = 1; i <= oWB.Sheets.Count; i++)
                    {
                        try
                        {
                            worksheet = (Excel.Worksheet)sheet.get_Item(i);
                            object workSheetTitle = worksheet.get_Range("AQ1", Missing.Value).Value2;
                            string workSheetTitleString = (workSheetTitle ?? String.Empty).ToString();
                            if (workSheetTitleString.ToUpper().Contains("GIRDER") && workSheetTitleString.Contains("BILL") == true)
                            {
                                girderWorksheetIndices.Add(i);
                            }
                        }
                        catch { }
                    }


                    List<object[,]> listGirderSheetMultiArray = new List<object[,]>();
                    object[,] girderSheetMultiArray = null;
                    for (int i = 0; i < girderWorksheetIndices.Count; i++)
                    {

                        object mark, quantity, designation, overallLengthFeet, overallLengthInches, TCwidth, TCXLfeet, TCXLinches, TCXLtype, TCXRfeet, TCXRinches,
                               TCXRtype, BCXLfeet, BCXLinches, BCXLtype, BCXRfeet, BCXRinches, BCXRtype, bearingDepthL, bearingDepthR, baySlope, baySlopeHE, LEseatSlope, LEseatHL,
                               REseatSlope, REseatHL, holeLeftFeet, holeLeftInches, holeRightFeet, holeRightInches, holeSize, holeGage, notes, unbrLen, NFB,
                               Aft, Ain, numbPanels, Panelft, Panelin, Bft, Bin, netUpliftKip, KnifePlateLE, KnifePlateRE, Adload, NetUplift, LeAxialW, LeAxialE, LeAxialEm, ReAxialW, ReAxialE, ReAxialEm;


                        mark = quantity = designation = overallLengthFeet = overallLengthInches = TCwidth = TCXLfeet = TCXLinches = TCXLtype = TCXRfeet = TCXRinches =
                               TCXRtype = BCXLfeet = BCXLinches = BCXLtype = BCXRfeet = BCXRinches = BCXRtype = bearingDepthL = bearingDepthR = baySlope = baySlopeHE = LEseatSlope = LEseatHL =
                               REseatSlope = REseatHL = holeLeftFeet = holeLeftInches = holeRightFeet = holeRightInches = holeSize = holeGage = notes = unbrLen = netUpliftKip = NFB =
                        Aft = Ain = numbPanels = Panelft = Panelin = Bft = Bin = KnifePlateLE = KnifePlateRE = Adload = NetUplift = LeAxialW = LeAxialE = LeAxialEm = ReAxialW = ReAxialE = ReAxialEm = null;

                        oSheet = (Excel._Worksheet)sheet.get_Item(girderWorksheetIndices[i]);

                        girderSheetMultiArray = (object[,])oSheet.get_Range("A12", "GI124").Value2;

                        for (int j = 1; j <= 9; j = j + 2)
                        {
                            if (girderSheetMultiArray[j, 1] != null)
                            {
                                mark = girderSheetMultiArray[j, 1];
                                quantity = girderSheetMultiArray[j, 9];

                                string girderLoad = (girderSheetMultiArray[j, 37] ?? String.Empty).ToString();
                                string[] girderLoadArray = girderLoad.Split(new string[] { "/" }, StringSplitOptions.None);
                                string girderTotalLoad = girderLoadArray[0];


                                overallLengthFeet = girderSheetMultiArray[j, 50];
                                overallLengthInches = girderSheetMultiArray[j, 56];
                                TCXLfeet = girderSheetMultiArray[j, 63];
                                TCXLinches = girderSheetMultiArray[j, 69];
                                TCXLtype = girderSheetMultiArray[j, 76];
                                TCXRfeet = girderSheetMultiArray[j, 79];
                                TCXRinches = girderSheetMultiArray[j, 85];
                                TCXRtype = girderSheetMultiArray[j, 92];
                                BCXLfeet = girderSheetMultiArray[j, 111];
                                BCXLinches = girderSheetMultiArray[j, 116];
                                BCXLtype = girderSheetMultiArray[j, 108];
                                BCXRfeet = girderSheetMultiArray[j, 126];
                                BCXRinches = girderSheetMultiArray[j, 131];
                                BCXRtype = girderSheetMultiArray[j, 123];
                                bearingDepthL = girderSheetMultiArray[j, 138];
                                bearingDepthR = girderSheetMultiArray[j, 145];
                                baySlope = girderSheetMultiArray[j, 152];
                                baySlopeHE = girderSheetMultiArray[j, 162];
                                LEseatSlope = girderSheetMultiArray[j, 165];
                                LEseatHL = girderSheetMultiArray[j, 175];
                                REseatSlope = girderSheetMultiArray[j, 178];
                                REseatHL = girderSheetMultiArray[j, 188];
                                holeLeftFeet = girderSheetMultiArray[j + 13, 33];
                                holeLeftInches = girderSheetMultiArray[j + 13, 38];
                                holeRightFeet = girderSheetMultiArray[j + 13, 70];
                                holeRightInches = girderSheetMultiArray[j + 13, 75];
                                holeSize = girderSheetMultiArray[j + 13, 47];
                                holeGage = girderSheetMultiArray[j + 13, 60];

                                Aft = girderSheetMultiArray[j + 39, 9];
                                Ain = girderSheetMultiArray[j + 39, 15];
                                numbPanels = girderSheetMultiArray[j + 39, 22];
                                Panelft = girderSheetMultiArray[j + 39, 30];
                                Panelin = girderSheetMultiArray[j + 39, 36];
                                Bft = girderSheetMultiArray[j + 39, 43];
                                Bin = girderSheetMultiArray[j + 39, 49];

                                KnifePlateLE = girderSheetMultiArray[j + 100, 159];
                                KnifePlateRE = girderSheetMultiArray[j + 100, 164];

                                NetUplift = joistSheetMultiArray[j + 26, 9];

                                Adload = joistSheetMultiArray[j + 26, 66];

                                LeAxialW = joistSheetMultiArray[j + 102, 59];
                                LeAxialE = joistSheetMultiArray[j + 102, 67];
                                LeAxialEm = joistSheetMultiArray[j + 102, 75];

                                ReAxialW = joistSheetMultiArray[j + 102, 88];
                                ReAxialE = joistSheetMultiArray[j + 102, 96];
                                ReAxialEm = joistSheetMultiArray[j + 102, 104];




                                double dblNetUpliftPLF = 0;
                                if (NetUplift != null | (NetUplift?? String.Empty).ToString() != "")
                                {
                                    try
                                    {
                                        dblNetUpliftPLF = Convert.ToDouble(NetUplift);
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Net Uplift is not in the correct form on mark " + (mark?.ToString()) + ". The converter will continue without net uplift on this mark but the user must fix the net uplift on this mark once it is complete");
                                    }
                                }

                                string hyphenJoistSpace = null;
                                string stringPanelft = ((Panelft ?? String.Empty).ToString());
                                string stringPanelin = ((Panelin ?? String.Empty).ToString());
                                bool isPanelftNull = (stringPanelft == null | stringPanelft == "");
                                bool isPanelinNull = (stringPanelin == null | stringPanelin == "");
                                if (isPanelftNull && isPanelinNull == true)
                                {
                                    hyphenJoistSpace = null;
                                }
                                else if (isPanelftNull == false && isPanelinNull == false)
                                {
                                    hyphenJoistSpace = stringPanelft + "-" + stringPanelin;
                                }
                                else if (isPanelftNull == false && isPanelinNull == true)
                                {
                                    hyphenJoistSpace = stringPanelft = "-0";
                                }
                                else if (isPanelftNull == true && isPanelinNull == false)
                                {
                                    hyphenJoistSpace = "0-" + stringPanelin;
                                }

                                string stringUpliftInKip = null;
                                if (hyphenJoistSpace != null)
                                {
                                    double dblHyphenJoistSpace = stringManipulation.hyphenLengthToDecimal(hyphenJoistSpace);
                                    double dblUpliftInKip = Math.Ceiling(((dblHyphenJoistSpace * dblNetUpliftPLF) / 1000.0) * 10) / 10;

                                    if (dblUpliftInKip != 0 | dblUpliftInKip != 0.0)
                                    {
                                        stringUpliftInKip = Convert.ToString(dblUpliftInKip);
                                    } 
                                }

                                if (hyphenJoistSpace == null && dblNetUpliftPLF != null && dblNetUpliftPLF != 0.0)
                                {
                                    MessageBox.Show("Mark " + mark?.ToString() + " has a net uplift but does not include a panel 'Space'. The converter will continue without net uplift on this mark but the user must fix the net uplift on this mark once it is complete.");
                                }

                                string stringDesignation = (girderSheetMultiArray[j, 15] ?? String.Empty).ToString()
                                 + (girderSheetMultiArray[j, 24] ?? String.Empty).ToString()
                                 + (girderSheetMultiArray[j, 30] ?? String.Empty).ToString()
                                 + (girderSheetMultiArray[j, 34] ?? String.Empty).ToString()
                                 + girderTotalLoad
                                 + (girderSheetMultiArray[j, 48] ?? String.Empty).ToString()
                                + stringUpliftInKip;

                                designation = stringDesignation;
                                // hole size is actually L= [i+13,57] +"X" + [i+13,64] w/ gage = [i+13,70]
                                //                       R= [i+13,94] +"X" + [i+13,101] w/ gage = [i+13,107]

                                //
                                if ((TCXLtype ?? String.Empty).ToString() == "F") { TCXLtype = "R"; }

                                if ((TCXRtype ?? String.Empty).ToString() == "F") { TCXRtype = "R"; }

                                if ((baySlope ?? String.Empty).ToString() != "")
                                {
                                    if ((baySlopeHE ?? String.Empty).ToString() == "L") { LEseatSlope = "-" + baySlope; REseatSlope = baySlope; }
                                    else { LEseatSlope = baySlope; REseatSlope = "-" + baySlope; }
                                }
                                else
                                {
                                    if ((LEseatSlope == null | (LEseatSlope ?? String.Empty).ToString() == "") && (REseatSlope == null | (REseatSlope ?? String.Empty).ToString() == ""))
                                    {
                                        LEseatSlope = REseatSlope = null;

                                    }
                                    else if ((LEseatSlope == null | (LEseatSlope ?? String.Empty).ToString() == ""))
                                    {
                                        if ((REseatHL ?? String.Empty).ToString() == "H") { REseatSlope = "-" + REseatSlope; }
                                        else { }
                                    }
                                    else if ((REseatSlope == null | (REseatSlope ?? String.Empty).ToString() == ""))
                                    {
                                        if ((LEseatHL ?? String.Empty).ToString() == "H") { LEseatSlope = "-" + LEseatSlope; }
                                        else { }
                                    }
                                    else
                                    {
                                        if ((REseatHL ?? String.Empty).ToString() == "H") { REseatSlope = "-" + REseatSlope; }
                                        else { }
                                        if ((LEseatHL ?? String.Empty).ToString() == "H") { LEseatSlope = "-" + LEseatSlope; }
                                        else { }


                                    }
                                }

                                object[] nucorGirderDataLine = new object[29];

                                nucorGirderDataLine = new object[] {mark, quantity, designation, overallLengthFeet, overallLengthInches, TCwidth, TCXLfeet,TCXLinches,TCXLtype,
                                                            TCXRfeet,TCXRinches,TCXRtype, bearingDepthL, bearingDepthR, BCXLfeet,BCXLinches,BCXRfeet,BCXRinches,
                                                            holeLeftFeet,holeLeftInches,holeRightFeet,holeRightInches,holeGage,
                                                           LEseatSlope,REseatSlope, notes, unbrLen, mark, NFB, Aft, Ain, numbPanels, Panelft, Panelin, Bft, Bin, KnifePlateLE, KnifePlateRE};
                                nucorGirderData.Add(nucorGirderDataLine);
                            }
                        }

                    }

                    oWB.Close(0);
                }
                
                NucorBOMCompleteInfo.Add(nucorJoistData);
                NucorBOMCompleteInfo.Add(nucorGirderData);

                oXL.Quit();
            }
            return NucorBOMCompleteInfo;
        }
        private static void Release(object obj)
        {
            // Errors are ignored per Microsoft's suggestion for this type of function:
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
            }
            catch
            {
            }
        }

    }
}
   

