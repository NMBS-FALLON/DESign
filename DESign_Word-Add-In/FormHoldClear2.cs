using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NMBS_2;
using Word = Microsoft.Office.Interop.Word;
using System.Globalization;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;

namespace NMBS_2
{
    public partial class FormHoldClear2 : Form
    {
        JoistCoverSheet JoistCoverSheet = new JoistCoverSheet();

        StringManipulation StringManipulation = new StringManipulation();



        public FormHoldClear2()
        {
            InitializeComponent();
        }

        List<CheckBox> cbLEList = new List<CheckBox>();
        List<CheckBox> cbREList = new List<CheckBox>();

        CheckBox cbAllLE = new CheckBox();
        CheckBox cbAllRE = new CheckBox();
        private void FormHoldClear2_Load(object sender, EventArgs e)
        {

            string clipboard = Clipboard.GetText();

            List<List<string>> joistData = JoistCoverSheet.JoistData();

            var labelMarkTitle = new Label();
            var labelLEBPL1 = new Label();
            var labelLEBPL2 = new Label();
            var labelREBPL1 = new Label();
            var labelREBPL2 = new Label();


            labelMarkTitle.Size = new System.Drawing.Size(60, 15);
            labelLEBPL1.Size = new System.Drawing.Size(80, 15);
            labelLEBPL2.Size = new System.Drawing.Size(60, 15);
            labelREBPL1.Size = new System.Drawing.Size(80, 15);
            labelREBPL2.Size = new System.Drawing.Size(60, 15);

            labelMarkTitle.Location = new Point(20, 60);
            labelLEBPL1.Location = new Point(85, 60);
            labelLEBPL2.Location = new Point(95, 75);
            labelREBPL1.Location = new Point(180, 60);
            labelREBPL2.Location = new Point(190, 75);

            labelMarkTitle.Text = "MARK";
            labelLEBPL1.Text = "LE HOLD";
            labelLEBPL2.Text = "CLEAR";
            labelREBPL1.Text = "RE HOLD";
            labelREBPL2.Text = "CLEAR";

            labelMarkTitle.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelLEBPL1.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelLEBPL2.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelREBPL1.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelREBPL2.Font = new Font("Times New Roman", 9, FontStyle.Bold);

            labelMarkTitle.TextAlign = ContentAlignment.MiddleLeft;
            labelLEBPL1.TextAlign = ContentAlignment.MiddleCenter;
            labelREBPL1.TextAlign = ContentAlignment.MiddleCenter;

            var labelAllMarks = new Label();

            labelAllMarks.Size = labelMarkTitle.Size;


            labelAllMarks.Location = new Point(20, 120);
            cbAllLE.Location = new Point(110, 120);
            cbAllRE.Location = new Point(210, 120);
            cbAllLE.Size = new System.Drawing.Size(20, 20);
            cbAllRE.Size = new System.Drawing.Size(20, 20);

            labelAllMarks.Text = "ALL";

            labelAllMarks.TextAlign = ContentAlignment.MiddleLeft;


            this.Controls.Add(labelMarkTitle);
            this.Controls.Add(labelLEBPL1);
            this.Controls.Add(labelLEBPL2);
            this.Controls.Add(labelREBPL1);
            this.Controls.Add(labelREBPL2);

            this.Controls.Add(labelAllMarks);
            this.Controls.Add(cbAllLE);
            this.Controls.Add(cbAllRE);


            List<string> joistMarks = joistData[0];

            int joistDataLength = joistMarks.Count();



            var labelMark = new Label[joistDataLength];
            var cbLEs = new CheckBox[joistDataLength];
            var cbREs = new CheckBox[joistDataLength];


            for (var i = 0; i < joistDataLength; i++)
            {
                var labelMarks = new Label();
                var cbLE = new CheckBox();
                var cbRE = new CheckBox();

                int Y = 150 + (i * 25);

                labelMarks.Text = joistMarks[i];
                labelMarks.Location = new Point(20, Y);

                labelMarks.Size = new System.Drawing.Size(50, 25);

                cbLE.Location = new Point(110, Y);
                cbLE.Size = new System.Drawing.Size(20, 20);

                cbRE.Location = new Point(210, Y);
                cbRE.Size = new System.Drawing.Size(20, 20);

                this.Controls.Add(cbLE);
                this.Controls.Add(labelMarks);
                this.Controls.Add(cbRE);

                cbLEList.Add(cbLE);
                cbREList.Add(cbRE);


                cbLEs[i] = cbLE;
                labelMark[i] = labelMarks;
                cbREs[i] = cbRE;
            }


        }

        public List<List<string>> joistDataByMark()
        {
            List<List<string>> shopOrderData = JoistCoverSheet.JoistData();
            List<List<string>> joistDataByMark = new List<List<string>>();
            for (int i = 0; i < shopOrderData[0].Count(); i++)
            {
                List<string> markData = new List<string>();
                markData.Add(shopOrderData[0][i]); //mark
                markData.Add(shopOrderData[1][i]); //quantity
                markData.Add(shopOrderData[4][i]); //TC size
                joistDataByMark.Add(markData);
            }
            return joistDataByMark;
        }
        public List<Tuple<string, bool, bool>> whichNeedHoldClears()
        {
            List<Tuple<string, bool, bool>> whichNeedHoldClears = new List<Tuple<string, bool, bool>>();
            List<List<string>> dataByMark = joistDataByMark();

            if (cbAllLE.Checked == true)
            {
                foreach (CheckBox cb in cbLEList)
                {
                    cb.Checked = true;
                }
            }

            if (cbAllRE.Checked == true)
            {
                foreach (CheckBox cb in cbREList)
                {
                    cb.Checked = true;
                }
            }

            for (int i = 0; i < dataByMark.Count(); i++)
            {
                var markHoldClears = new Tuple<string, bool, bool>(dataByMark[i][0], cbLEList[i].Checked, cbREList[i].Checked);
                whichNeedHoldClears.Add(markHoldClears);
            }

            return whichNeedHoldClears;

        }

        private List<List<string>> getHoldClearData()
        {
            List<List<string>> joistDataByMarks = joistDataByMark();

            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Find.Execute("COLOR CODE");
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            
            
            List<Tuple<string, bool, bool>> marksWithHoldClears = whichNeedHoldClears();
            List<List<string>> holdClearData = new List<List<string>>();
            
            for (int i=0; i<marksWithHoldClears.Count();i++)
            {
                
                string mark = marksWithHoldClears[i].Item1;
                selection.Find.Execute(mark);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                if (marksWithHoldClears[i].Item2 == true)
                {
                    selection.Find.Execute("BPL-L");
                    selection.HomeKey(Word.WdUnits.wdLine, 0);
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    string BPLline = selection.Text;
                    string[] BPLlineArray = BPLline.Split(new string[] { "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);
                    List<string> markLEBPLdata = new List<string>();
                    markLEBPLdata.Add(mark);
                    markLEBPLdata.Add(joistDataByMarks[i][1]); //qty
                    markLEBPLdata.Add(joistDataByMarks[i][2]); //TC size
                    markLEBPLdata.Add("BPL-L");                //base plate side
                    markLEBPLdata.Add(BPLlineArray[3]);        //BPL Length
                    if (BPLlineArray[5].Contains("|") == false)
                    {
                        markLEBPLdata.Add(BPLlineArray[5]);   //base plate depth outside
                        markLEBPLdata.Add(BPLlineArray[5]);   //base plate depth inside
                    }
                    if (BPLlineArray[5].Contains("|") == true)
                    {
                        if (BPLlineArray[4].Contains("BPL-L"))
                        {
                            string[] bpl = BPLlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                            markLEBPLdata.Add(bpl[0]);           //base plate depth outside
                            markLEBPLdata.Add(bpl[1]);           //base plate depth inside
                        }
                        if (BPLlineArray[4].Contains("BPL-R"))
                        {
                            string[] bpl = BPLlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                            markLEBPLdata.Add(bpl[1]);           //base plate depth outside
                            markLEBPLdata.Add(bpl[0]);           //base plate depth inside
                        }
                    }
                    markLEBPLdata.Add(BPLlineArray[6]);      // slot setback
                    markLEBPLdata.Add(BPLlineArray[7]);      // slot size
                    markLEBPLdata.Add(BPLlineArray[8]);      // slot gauge

                    holdClearData.Add(markLEBPLdata);
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }

                if (marksWithHoldClears[i].Item3 == true)
                {
                    selection.Find.Execute("BPL-R");
                    selection.HomeKey(Word.WdUnits.wdLine, 0);
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    string BPRline = selection.Text;
                    string[] BPRlineArray = BPRline.Split(new string[] { "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);
                    List<string> markREBPLdata = new List<string>();
                    markREBPLdata.Add(mark);
                    markREBPLdata.Add(joistDataByMarks[i][1]); //qty
                    markREBPLdata.Add(joistDataByMarks[i][2]); //TC size
                    markREBPLdata.Add("BPL-R");                //base plate side
                    markREBPLdata.Add(BPRlineArray[3]);        //BPL Length
                    if (BPRlineArray[5].Contains("|") == false)
                    {
                        markREBPLdata.Add(BPRlineArray[5]);   //base plate depth outside
                        markREBPLdata.Add(BPRlineArray[5]);   //base plate depth inside
                    }
                    if (BPRlineArray[5].Contains("|") == true)
                    {
                        if (BPRlineArray[4].Contains("BPL-L"))
                        {
                            string[] bpl = BPRlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                            markREBPLdata.Add(bpl[0]);           //base plate depth outside
                            markREBPLdata.Add(bpl[1]);           //base plate depth inside
                        }
                        if (BPRlineArray[4].Contains("BPL-R"))
                        {
                            string[] bpl = BPRlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                            markREBPLdata.Add(bpl[1]);           //base plate depth outside
                            markREBPLdata.Add(bpl[0]);           //base plate depth inside
                        }
                    }
                    markREBPLdata.Add(BPRlineArray[6]);      // slot setback
                    markREBPLdata.Add(BPRlineArray[7]);      // slot size
                    markREBPLdata.Add(BPRlineArray[8]);      // slot gauge

                    holdClearData.Add(markREBPLdata);
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }
            }
            return holdClearData;

        }

        private List<joistSeatInfo> getJoistSeatInfo()
        {
            List<List<string>> joistDataByMarks = joistDataByMark();

            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Find.Execute("COLOR CODE");
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);


            List<Tuple<string, bool, bool>> marksWithHoldClears = whichNeedHoldClears();
            List<joistSeatInfo> joistSeatInfo = new List<joistSeatInfo>();

            for (int i = 0; i < marksWithHoldClears.Count(); i++)
            {
                joistSeatInfo currentJoistSeatInfo;


                string mark = marksWithHoldClears[i].Item1;
                selection.HomeKey(Word.WdUnits.wdStory, 0);
                selection.Find.Execute("COLOR CODE");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selection.Find.Execute(mark);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                if (marksWithHoldClears[i].Item2 == true)
                {
                    currentJoistSeatInfo = new joistSeatInfo();

                    selection.Find.Execute("BPL-L");
                    selection.HomeKey(Word.WdUnits.wdLine, 0);
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    string BPLline = selection.Text;
                    string[] BPLlineArray = BPLline.Split(new string[] { "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);
                    List<string> markBPLdata = new List<string>();

                    currentJoistSeatInfo.mark = mark;
                    currentJoistSeatInfo.qty = Convert.ToInt16(joistDataByMarks[i][1]);
                    currentJoistSeatInfo.TC=joistDataByMarks[i][2]; 
                    currentJoistSeatInfo.bplSide="BPL-L";       
                    currentJoistSeatInfo.bplLength = BPLlineArray[3];        
                    if (BPLlineArray[5].Contains("|") == false)
                    {
                        currentJoistSeatInfo.bplOutsideDepth=StringManipulation.ConvertLengthtoDecimal("0-" +BPLlineArray[5]);   
                        currentJoistSeatInfo.bplInsideDepth=StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);  
                    }
                    if (BPLlineArray[5].Contains("|") == true)
                    {
                        string[] bpl = BPLlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                        currentJoistSeatInfo.bplOutsideDepth=StringManipulation.ConvertLengthtoDecimal("0-" +bpl[0]);        
                        currentJoistSeatInfo.bplInsideDepth=StringManipulation.ConvertLengthtoDecimal("0-" +bpl[1]);      
                    }
                    if (BPLlineArray.Count() > 6)
                    {
                        currentJoistSeatInfo.slotSetback = BPLlineArray[6];      // slot setback
                        currentJoistSeatInfo.slotSize = BPLlineArray[7];      // slot size
                        currentJoistSeatInfo.slotGauge = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[8]);      // slot gauge
                    }

                    joistSeatInfo.Add(currentJoistSeatInfo);
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }

                if (marksWithHoldClears[i].Item3 == true)
                {
                    currentJoistSeatInfo = new joistSeatInfo();
                    selection.Find.Execute("BPL-R");
                    selection.HomeKey(Word.WdUnits.wdLine, 0);
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    string BPLline = selection.Text;
                    string[] BPLlineArray = BPLline.Split(new string[] { "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);
                    List<string> markBPLdata = new List<string>();

                    currentJoistSeatInfo.mark = mark;
                    currentJoistSeatInfo.qty = Convert.ToInt16(joistDataByMarks[i][1]);
                    currentJoistSeatInfo.TC = joistDataByMarks[i][2];
                    currentJoistSeatInfo.bplSide = "BPL-R";
                    currentJoistSeatInfo.bplLength = BPLlineArray[3];
                    if (BPLlineArray[5].Contains("|") == false)
                    {
                        currentJoistSeatInfo.bplOutsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" +BPLlineArray[5]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);
                    }
                    if (BPLlineArray[5].Contains("|") == true)
                    {
                        string[] bpl = BPLlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                        currentJoistSeatInfo.bplOutsideDepth =StringManipulation.ConvertLengthtoDecimal("0-" + bpl[1]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + bpl[0]);
                    }
                    if (BPLlineArray.Count() > 6)
                    {
                        currentJoistSeatInfo.slotSetback = BPLlineArray[6];      // slot setback
                        currentJoistSeatInfo.slotSize = BPLlineArray[7];      // slot size
                        currentJoistSeatInfo.slotGauge = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[8]);      // slot gauge
                    }
                    joistSeatInfo.Add(currentJoistSeatInfo);
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }
            }
            return joistSeatInfo;

        }

        public List<HCSeatInfo> listHCSeatInfo (List<joistSeatInfo> listJoistSeatInfo)
        {
            List<HCSeatInfo> listHCSeatInfo = new List<HCSeatInfo>();

            foreach (joistSeatInfo thisJoistSeatInfo in listJoistSeatInfo)
            {
                HCSeatInfo thisHCSeatInfo = new HCSeatInfo();
                thisHCSeatInfo.mark = thisJoistSeatInfo.mark;
                thisHCSeatInfo.qty = thisJoistSeatInfo.qty*2;
                thisHCSeatInfo.TC = thisJoistSeatInfo.TC;
                thisHCSeatInfo.bplSide = thisJoistSeatInfo.bplSide;
                thisHCSeatInfo.bplLength = StringManipulation.cleanDecimalToHyphen(StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.bplLength));
                thisHCSeatInfo.bplOutsideDepth = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth);
                thisHCSeatInfo.bplInsideDepth = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth);
                thisHCSeatInfo.slotSetback = StringManipulation.cleanDecimalToHyphen(StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.slotSetback));
                thisHCSeatInfo.slotSize = thisJoistSeatInfo.slotSize;

                int diffInHCandTC = 0;

                if (thisJoistSeatInfo.TC == "3531" || thisJoistSeatInfo.TC == "3534")
                {
                    diffInHCandTC = 1;
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - 3.5 / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - 3.5 / 12.0);
                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= 7.5 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= 7.5)
                    {
                        thisHCSeatInfo.HCMaterial = "4037";
                    }
                    else
                    {
                        thisHCSeatInfo.HCMaterial = "406037 (HOR. LEG =4\")";
                    }



                }

                if (thisJoistSeatInfo.TC == "4037")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - 4.0 / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - 4.0 / 12.0);

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= 7.5 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= 7.5)
                    {
                        thisHCSeatInfo.HCMaterial = "3534";
                        diffInHCandTC = -1;
                    }
                    else
                    {
                        thisHCSeatInfo.HCMaterial = "406037 (HOR. LEG = 4\")";
                        diffInHCandTC = 0;
                    }
                }

                if (thisJoistSeatInfo.TC == "4043")
                {

                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - 4.0 / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - 4.0 / 12.0);

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= 7.5 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= 7.5)
                    {
                        thisHCSeatInfo.HCMaterial = "3534";
                        diffInHCandTC = -1;
                    }
                    else if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= 8.0 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= 8.0)
                    {
                        thisHCSeatInfo.HCMaterial = "4050";
                        thisHCSeatInfo.diffInHCandTC = 1;
                    }
                    else
                    {
                        thisHCSeatInfo.HCMaterial = "407050 (HOR. LEG = 4\")";
                        diffInHCandTC = 1;
                    }
                }
                if (thisJoistSeatInfo.TC == "4050")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - 4.0 / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - 4.0 / 12.0);

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= 8.0 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= 8.0)
                    {
                        thisHCSeatInfo.HCMaterial = "4050";
                        diffInHCandTC = 0;
                    }
                    else
                    {
                        thisHCSeatInfo.HCMaterial = "407050 (HOR. LEG = 4\")";
                        diffInHCandTC = 0;
                    }
                }
                if (thisJoistSeatInfo.TC == "5043")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - 5.0 / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - 5.0 / 12.0);

                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = 1;

                }
                if (thisJoistSeatInfo.TC == "5050" )
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - 5.0 / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - 5.0 / 12.0);

                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = 0;
                }
                if (thisJoistSeatInfo.TC == "6050")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - 6.0 / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - 6.0 / 12.0);

                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = 0;
                }
                if (thisJoistSeatInfo.TC == "6056")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - 6.0 / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - 6.0 / 12.0);
                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = -1;

                }
                if (thisJoistSeatInfo.TC == "6075")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - 6.0 / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - 6.0 / 12.0);
                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = -2;

                }

                int intGOL = Convert.ToInt16((thisJoistSeatInfo.slotGauge-(1.0/12.0))/2)*12* 16 + diffInHCandTC;
                double dblGOL = ((thisJoistSeatInfo.slotGauge - (1.0 / 12.0)) / 2.0) + Convert.ToDouble(diffInHCandTC) / (12*16.0);
                thisHCSeatInfo.GOL = StringManipulation.cleanDecimalToHyphen16ths(dblGOL);

                listHCSeatInfo.Add(thisHCSeatInfo);
            }
            return listHCSeatInfo;
        }

        public List<HCSeatInfo> organizedSeatInfoList ()
        {
            List<HCSeatInfo> seatList = listHCSeatInfo(getJoistSeatInfo());

            for (int index = 0; index < seatList.Count(); index++ )
            {
                if (seatList[index].bplSide == "BPL-L")
                {
                    seatList[index].mark = seatList[index].mark + " LE";
                }
                if (seatList[index].bplSide == "BPL-R")
                {
                    seatList[index].mark = seatList[index].mark + " RE";
                }
            }

                for (int index = 1; index < seatList.Count(); index++)
                {
                    if (seatList[index].mark.Substring(0, seatList[index].mark.Length - 3) == seatList[index - 1].mark.Substring(0, seatList[index-1].mark.Length - 3) &&
                        seatList[index].HCMaterial == seatList[index - 1].HCMaterial &&
                        seatList[index].bplLength == seatList[index - 1].bplLength &&
                        seatList[index].HCInsideHeight == seatList[index - 1].HCInsideHeight &&
                        seatList[index].HCOutsideHeight == seatList[index - 1].HCOutsideHeight &&
                        seatList[index].slotSetback == seatList[index - 1].slotSetback &&
                        seatList[index].GOL == seatList[index - 1].GOL)
                    {
                        seatList[index - 1].mark = seatList[index - 1].mark.Substring(0, seatList[index-1].mark.Length - 3) + " BE";
                        seatList[index - 1].qty = seatList[index - 1].qty + seatList[index].qty;
                        seatList.RemoveAt(index);
                        index--;

                    }

                }

            for (int index = 1; index < seatList.Count(); index++)
            {
                for (int index2 = 0; index2 < seatList.Count(); index2++)
                {
                    if (index2 != index)
                    {
                        if (seatList[index].HCMaterial == seatList[index2].HCMaterial &&
                            seatList[index].bplLength == seatList[index2].bplLength &&
                            seatList[index].HCInsideHeight == seatList[index2].HCInsideHeight &&
                            seatList[index].HCOutsideHeight == seatList[index2].HCOutsideHeight &&
                            seatList[index].slotSetback == seatList[index2].slotSetback &&
                            seatList[index].GOL == seatList[index2].GOL)
                        {
                            if (index < index2)
                            {
                                seatList[index2].mark = seatList[index].mark + ", " + seatList[index2].mark;
                            }
                            else
                            {
                                seatList[index2].mark = seatList[index2].mark + ", " + seatList[index].mark;
                            }
                            seatList[index2].qty = seatList[index2].qty + seatList[index].qty;
                            seatList.RemoveAt(index);
                            index--;
                            break;

                        }
                    }
                }

            }

            return seatList;
        }

        public void createHoldClearSKs ()
        {
            List<List<string>> joistData = JoistCoverSheet.JoistData();
   //         List<joistSeatInfo> joistSeatInfo = getJoistSeatInfo();
   //         List<HCSeatInfo> HCSeatInfoList = listHCSeatInfo(joistSeatInfo);
            List<HCSeatInfo> HCSeatInfoList = organizedSeatInfoList();
            
            Image SK1 = Properties.Resources.HCCombined;

            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Section section = selection.Sections.Add();
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            selection.PageSetup.LeftMargin = (float)50;
            selection.PageSetup.RightMargin = (float)50;

            Word.Table tableHoldClearCover = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 2, 6);

            tableHoldClearCover.Cell(1, 1).Range.Text = "JOB NAME: ";
            tableHoldClearCover.Cell(2, 1).Range.Text = "LOCATION: ";
            tableHoldClearCover.Cell(1, 3).Range.Text = "JOB #: ";
            tableHoldClearCover.Cell(2, 3).Range.Text = "LIST:  ";

            for (int i = 1; i <= 2; i++)
            {
                tableHoldClearCover.Cell(i, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableHoldClearCover.Cell(i, 1).Range.Underline = Word.WdUnderline.wdUnderlineSingle;
                tableHoldClearCover.Cell(i, 1).Range.Font.Bold = 1;
            }

            for (int i = 1; i <= 2; i++)
            {
                tableHoldClearCover.Cell(i, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableHoldClearCover.Cell(i, 3).Range.Underline = Word.WdUnderline.wdUnderlineSingle;
                tableHoldClearCover.Cell(i, 3).Range.Font.Bold = 1;
            }


            tableHoldClearCover.Cell(1, 2).Range.Text = joistData[8][0];
            tableHoldClearCover.Cell(2, 2).Range.Text = joistData[9][0];
            tableHoldClearCover.Cell(1, 4).Range.Text = joistData[7][0];
            tableHoldClearCover.Cell(2, 4).Range.Text = joistData[10][0];

            for (int i = 1; i <= 2; i++)
            {
                tableHoldClearCover.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }
            for (int i = 1; i <= 2; i++)
            {
                tableHoldClearCover.Cell(i, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }

            tableHoldClearCover.Cell(1, 5).Range.Text = "SHEET #:";
            tableHoldClearCover.Cell(1, 5).Borders.Enable = 1;
            tableHoldClearCover.Cell(2, 5).Borders.Enable = 1;
            tableHoldClearCover.Cell(1, 5).Range.Font.Bold = 1;
            tableHoldClearCover.Cell(2, 5).Range.Font.Bold = 1;





            selection.EndKey(Word.WdUnits.wdStory, 1);

            tableHoldClearCover.Columns[1].Width = 65;
            tableHoldClearCover.Columns[2].Width = 250;
            tableHoldClearCover.Columns[3].Width = 50;
            tableHoldClearCover.Columns[4].Width = 208;
            tableHoldClearCover.Columns[5].Width = 80;













            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.EndKey(Word.WdUnits.wdStory);
            selection.Text = "\r\n";
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            Clipboard.SetImage(SK1);


            selection.Paste();
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.Text = "\r\n";
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            Word.Table tblHC = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 3, 11);

            tblHC.Columns[1].Width = 45;
            tblHC.Columns[2].Width = 210;
            tblHC.Columns[3].Width = 45;
            tblHC.Columns[4].Width = 45;
            tblHC.Columns[5].Width = 45;
            tblHC.Columns[6].Width = 45;
            tblHC.Columns[7].Width = 45;
            tblHC.Columns[8].Width = 55;
            tblHC.Columns[9].Width = 60;
            tblHC.Columns[10].Width =45;
            tblHC.Columns[11].Width = 45;

            tblHC.Borders.Enable = 1;

            

            

            tblHC.Cell(1, 7).Merge(MergeTo: tblHC.Cell(1, 9));
            tblHC.Cell(1,7).Borders.Enable = 1;

            tblHC.Cell(1, 7).Range.Text = "SLOTS";
            tblHC.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            tblHC.Cell(2, 1).Range.Text = "Qty";
            tblHC.Cell(2, 2).Range.Text = "Mark(s)";
            tblHC.Cell(2, 3).Range.Text = "MAT.";
            tblHC.Cell(2, 4).Range.Text = "H";
            tblHC.Cell(2, 5).Range.Text = "A";
            tblHC.Cell(2, 6).Range.Text = "B";
            tblHC.Cell(2, 7).Range.Text = "E";
            tblHC.Cell(2, 8).Range.Text = "GOL";
            tblHC.Cell(2, 9).Range.Text = "S";
            tblHC.Cell(2, 10).Select();
            selection.Text = "out";
            selection.Font.Subscript = 1;
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.Text = "d";
            selection.Font.Subscript = 0;
            tblHC.Cell(2, 11).Select();
            selection.Text = "in";
            selection.Font.Subscript = 1;
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.Text = "d";
            selection.Font.Subscript = 0;

            for (int col = 1; col <= 11; col++ )
            {
                tblHC.Cell(2, col).Range.Font.Size = 8;
                tblHC.Cell(2,col).Borders.Enable = 1;
                tblHC.Cell(2, col).Range.Font.Bold = 1;
                tblHC.Cell(2, col).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            }

            int row = 3;

            foreach (HCSeatInfo hcSeatInfo in HCSeatInfoList )
            {
                tblHC.Cell(row, 1).Range.Text = hcSeatInfo.qty.ToString();
                tblHC.Cell(row, 2).Range.Text = hcSeatInfo.mark;
                tblHC.Cell(row, 3).Range.Text = hcSeatInfo.HCMaterial;
                tblHC.Cell(row, 4).Range.Text = hcSeatInfo.bplLength;
                tblHC.Cell(row, 5).Range.Text = hcSeatInfo.HCOutsideHeight;
                tblHC.Cell(row, 6).Range.Text = hcSeatInfo.HCInsideHeight;
                tblHC.Cell(row, 7).Range.Text = hcSeatInfo.slotSetback;
                tblHC.Cell(row, 8).Range.Text = hcSeatInfo.GOL;
                tblHC.Cell(row, 9).Range.Text = hcSeatInfo.slotSize;
                tblHC.Cell(row, 10).Range.Text = hcSeatInfo.bplOutsideDepth;
                tblHC.Cell(row, 11).Range.Text = hcSeatInfo.bplInsideDepth;


                tblHC.Rows.Add();
                row = row + 1;
                
            }

            
            for (int row1 = 2; row1 <= tblHC.Rows.Count; row1++ )
            {
                for (int col1 =1; col1<=tblHC.Columns.Count; col1++)
                {
                    if (col1 != 2)
                    {
                        tblHC.Cell(row1, col1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }
            }




                this.Close();
        }




        private void btnCreateHoldClears_Click(object sender, EventArgs e)
        {
            createHoldClearSKs();
        }
        public Image textOnImage(Image image, List<Tuple<string, int, int, int>> listOfText)
        {
            Bitmap bitMapImage = new
                System.Drawing.Bitmap(image);
            Graphics graphicImage = Graphics.FromImage(bitMapImage);

            graphicImage.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            foreach (Tuple<string,int,int,int> thisText in listOfText)
            {
                Font font1 = new Font("Calibri", thisText.Item2, FontStyle.Regular);
                Point point1 = new Point(thisText.Item3, thisText.Item4);
                graphicImage.DrawString(thisText.Item1, font1, SystemBrushes.WindowText, point1);
            }



            bitMapImage = BitmapTo1Bpp(bitMapImage);

            return bitMapImage;
        }

        public static Bitmap BitmapTo1Bpp(Bitmap img)
        {
            int w = img.Width;
            int h = img.Height;
            Bitmap bmp = new Bitmap(w, h, PixelFormat.Format1bppIndexed);
            BitmapData data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadWrite, PixelFormat.Format1bppIndexed);
            byte[] scan = new byte[(w + 7) / 8];
            for (int y = 0; y < h; y++)
            {
                for (int x = 0; x < w; x++)
                {
                    if (x % 8 == 0) scan[x / 8] = 0;
                    Color c = img.GetPixel(x, y);
                    if (c.GetBrightness() >= 0.5) scan[x / 8] |= (byte)(0x80 >> (x % 8));
                }
                Marshal.Copy(scan, 0, (IntPtr)((long)data.Scan0 + data.Stride * y), scan.Length);
            }
            bmp.UnlockBits(data);
            return bmp;
        }

   

    }


    
    public class joistSeatInfo
    {
        public string mark;
        public int qty;
        public string TC;
        public string bplSide;
        public string bplLength;
        public double bplOutsideDepth;
        public double bplInsideDepth;
        public string slotSetback;
        public string slotSize;
        public double slotGauge;
    }
    public class HCSeatInfo
    {
        public string mark;
        public int qty;
        public string TC;
        public string HCMaterial;
        public int diffInHCandTC;
        public string bplSide;
        public string bplLength;
        public string bplOutsideDepth;
        public string HCOutsideHeight;
        public string bplInsideDepth;
        public string HCInsideHeight;
        public string slotSetback;
        public string slotSize;
        public string GOL;
    }

    


}
