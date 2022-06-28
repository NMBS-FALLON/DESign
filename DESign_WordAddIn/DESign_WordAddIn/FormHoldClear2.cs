using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DESign_WordAddIn;
using Word = Microsoft.Office.Interop.Word;
using System.Globalization;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Linq;
using System.Xml.Linq;
using System.Reflection;
using System.Threading;
using DESign_BASE;
using System.Collections.Generic;


namespace DESign_WordAddIn
{
    public partial class FormHoldClear2 : Form
    {
        JoistCoverSheet JoistCoverSheet = new JoistCoverSheet();

        StringManipulation StringManipulation = new StringManipulation();

        DESign_BASE.QueryAngleData QueryAngleData = new DESign_BASE.QueryAngleData();
        List<DESign_BASE.Angle> anglesFromSql = QueryAngleData.AnglesFromSql("Fallon");

        

        public FormHoldClear2()
        {
            InitializeComponent();
        }

        List<CheckBox> cbLEList = new List<CheckBox>();
        List<CheckBox> cbREList = new List<CheckBox>();

        CheckBox cbAllLE = new CheckBox();
        CheckBox cbAllRE = new CheckBox();
        ComboBox allDetailOveride = new ComboBox();

        List<List<string>> joistData;
        List<ComboBox> detailComboBoxs = new List<ComboBox>();
        DataTable overideDetails = new DataTable();
        bool noOverides;

        private void FormHoldClear2_Load(object sender, EventArgs e)
        {
            overideDetails.Columns.Add("Mark", typeof(string));
            overideDetails.Columns.Add("Hold Clear Type", typeof(string));

            string clipboard = Clipboard.GetText();

            joistData = JoistCoverSheet.JoistData();

            var labelMarkTitle = new Label();
            var labelAutoFilledTitle = new Label();
            var labelLEBPL1 = new Label();
            var labelREBPL1 = new Label();

            labelMarkTitle.Size = new System.Drawing.Size(60, 15);



            labelMarkTitle.AutoSize = true;
            labelAutoFilledTitle.AutoSize = true;
            labelLEBPL1.AutoSize = true;
            labelREBPL1.AutoSize = true;

            
            labelMarkTitle.Location = new Point(20, 60);
            labelAutoFilledTitle.Location = new Point(85, 60);
            labelLEBPL1.Location = new Point(185, 60);
            labelREBPL1.Location = new Point(280, 60);


            labelMarkTitle.Text = "MARK";
            labelAutoFilledTitle.Text = "AUTO-\nFILLED?";
            labelLEBPL1.Text = "LE HOLD\nCLEAR";
            labelREBPL1.Text = "RE HOLD\nCLEAR";


            labelMarkTitle.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelAutoFilledTitle.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelLEBPL1.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelREBPL1.Font = new Font("Times New Roman", 9, FontStyle.Bold);

            labelMarkTitle.TextAlign = ContentAlignment.MiddleLeft;
            labelAutoFilledTitle.TextAlign = ContentAlignment.MiddleCenter;
            labelLEBPL1.TextAlign = ContentAlignment.MiddleCenter;
            labelREBPL1.TextAlign = ContentAlignment.MiddleCenter;

            var labelOveride = new Label();
            labelOveride.AutoSize = true;
            labelOveride.Location = new Point(385, 60);
            labelOveride.Text = "DETAIL\nOVERIDE";
            labelOveride.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelOveride.TextAlign = ContentAlignment.MiddleLeft;

            var labelAllMarks = new Label();
            labelAllMarks.Size = labelMarkTitle.Size;
            labelAllMarks.Location = new Point(20, 120);
            cbAllLE.Location = new Point(210, 120);
            cbAllRE.Location = new Point(310, 120);
            cbAllLE.Size = new System.Drawing.Size(20, 20);
            cbAllRE.Size = new System.Drawing.Size(20, 20);
            labelAllMarks.Text = "ALL";
            labelAllMarks.TextAlign = ContentAlignment.MiddleLeft;



            allDetailOveride.DrawMode = System.Windows.Forms.DrawMode.Normal;
            allDetailOveride.Location = new Point(385, 120);
            allDetailOveride.Size = new System.Drawing.Size(70, 20);

            string[] allHCdetails = new string[] { "","Butted", "Gapped", "1/4\" Plate", "1/2\" Plate" };

            allDetailOveride.DataSource = allHCdetails;
            allDetailOveride.DropDownWidth = 70;
            allDetailOveride.DropDownStyle = ComboBoxStyle.DropDownList;
            allDetailOveride.Enabled = true;



            this.Controls.Add(labelMarkTitle);
            this.Controls.Add(labelAutoFilledTitle);
            this.Controls.Add(labelLEBPL1);
            this.Controls.Add(labelREBPL1);

            this.Controls.Add(labelAllMarks);
            this.Controls.Add(cbAllLE);
            this.Controls.Add(cbAllRE);
            this.Controls.Add(allDetailOveride);

            this.Controls.Add(labelOveride);


            List<string> joistMarks = joistData[0];

            int joistDataLength = joistMarks.Count();
            Dictionary<string, ExcelDataExtraction.HoldClearInformation> allHoldClearInformation = new Dictionary<string, ExcelDataExtraction.HoldClearInformation>();


            var labelMark = new Label[joistDataLength];
            var cbLEs = new CheckBox[joistDataLength];
            var cbREs = new CheckBox[joistDataLength];

            string noteInfoPath = Globals.ThisAddIn.Application.ActiveDocument.Path.ToString() + "\\Note Info.xlsx";
            if (System.IO.File.Exists(noteInfoPath))
            {
                ExcelDataExtraction excelDataExtraction = new ExcelDataExtraction();
                allHoldClearInformation = excelDataExtraction.getHoldClearInfo(noteInfoPath);
            }


            for (var i = 0; i < joistDataLength; i++)
            {
                string mark = joistMarks[i];
                bool hcLeft = false;
                bool hcRight = false;
                var cbAutoFilled = new CheckBox();
                cbAutoFilled.AutoCheck = false;

                if (allHoldClearInformation.ContainsKey(mark))
                {
                    hcLeft = allHoldClearInformation[mark].HCLeft;
                    hcRight = allHoldClearInformation[mark].HCRight;
                    cbAutoFilled.Checked = true;
                }

                var labelMarks = new Label();
                var cbLE = new CheckBox();
                cbLE.Checked = hcLeft;
                var cbRE = new CheckBox();
                cbRE.Checked = hcRight;
                var detailComboBox = new ComboBox();

                int Y = 150 + (i * 25);

                labelMarks.Text = mark;
                labelMarks.Location = new Point(20, Y);
                labelMarks.Size = new System.Drawing.Size(50, 25);

                cbAutoFilled.Location = new Point(110, Y);
                cbAutoFilled.Size = new System.Drawing.Size(20, 20);

                cbLE.Location = new Point(210, Y);
                cbLE.Size = new System.Drawing.Size(20, 20);

                cbRE.Location = new Point(310, Y);
                cbRE.Size = new System.Drawing.Size(20, 20);

                detailComboBox.DrawMode = System.Windows.Forms.DrawMode.Normal;
                detailComboBox.Location = new Point(385, Y);
                detailComboBox.Size = new System.Drawing.Size(70, 20);


                string[] HCdetails = new string[] { "", "Butted", "Gapped", "1/4\" Plate", "1/2\" Plate" };

                detailComboBox.DataSource = HCdetails;
                detailComboBox.DropDownWidth = 70;
                detailComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                detailComboBox.Enabled = true;


                this.Controls.Add(labelMarks);
                this.Controls.Add(cbAutoFilled);
                this.Controls.Add(cbLE);
                this.Controls.Add(cbRE);
                this.Controls.Add(detailComboBox);

                cbLEList.Add(cbLE);
                cbREList.Add(cbRE);
                detailComboBoxs.Add(detailComboBox);


                cbLEs[i] = cbLE;
                labelMark[i] = labelMarks;
                cbREs[i] = cbRE;

            }


        }

        public List<List<string>> joistDataByMark()
        {
            List<List<string>> joistDataByMark = new List<List<string>>();
            for (int i = 0; i < joistData[0].Count(); i++)
            {
                List<string> markData = new List<string>();
                markData.Add(joistData[0][i]); //mark
                markData.Add(joistData[1][i]); //quantity
                markData.Add(joistData[4][i]); //TC size
                joistDataByMark.Add(markData);
            }
            return joistDataByMark;
        }
        public List<Tuple<string, bool, bool>> whichNeedHoldClears(List<List<string>> dataByMark)
        {
            List<Tuple<string, bool, bool>> whichNeedHoldClears = new List<Tuple<string, bool, bool>>();

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
            
            noOverides = true;

            for(int i=0; i<dataByMark.Count(); i++)
            {
                if (allDetailOveride.Text == "")
                {
                    overideDetails.Rows.Add(dataByMark[i][0], detailComboBoxs[i].Text);
                }
                else
                {
                    overideDetails.Rows.Add(dataByMark[i][0], allDetailOveride.Text);
                    noOverides = false;
                }
                if (detailComboBoxs[i].Text !="")
                {
                    noOverides = false;
                }
            }

            return whichNeedHoldClears;

        }

        private AllSeatInfo getAllSeatInfo(List<List<string>> joistDataByMarks, List<Tuple<string, bool, bool>> marksWithHoldClears)
        {


            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Find.Execute("COLOR CODE");
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);



            List<JoistSeatInfo> holdClear = new List<JoistSeatInfo>();
            List<JoistSeatInfo> tPlate = new List<JoistSeatInfo>();
            List<JoistSeatInfo> bcg = new List<JoistSeatInfo>();
            List<JoistSeatInfo> standard = new List<JoistSeatInfo>();


            for (int i = 0; i < joistDataByMarks.Count(); i++)
            {

                JoistSeatInfo currentJoistSeatInfo;

                string mark = joistDataByMarks[i][0];
                //                string markWithHC = marksWithHoldClears[i].Item1;
                selection.HomeKey(Word.WdUnits.wdStory, 0);
                selection.Find.Execute("COLOR CODE");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selection.Find.Execute(mark+ "    ");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selection.Find.Execute("TOTAL LENGTH");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.MoveDown(Word.WdUnits.wdLine, 1);
                selection.HomeKey(Word.WdUnits.wdLine);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.EndKey(Extend: Word.WdUnits.wdLine);
                string tcxLine = selection.Text;
                string[] tcxLineArray = tcxLine.Split(new string[] { "                               ", "                              ", "                             ", "                            ", "                           ", "                          ", "                         ", "                        ", "                       ", "                      ", "                     ", "                    ", "                   ", "                  ", "                 ", "                ", "               ", "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.Find.Execute("CLEAR ");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selection.EndKey(Extend: Word.WdUnits.wdLine);
                string clearLine = selection.Text;
                string[] clearLineArray = clearLine.Split(new string[] { "CLEAR ", "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.Find.Execute("SETBACK");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.HomeKey(Extend: Word.WdUnits.wdLine);
                int charsToSetback = selection.Text.Length;
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                string[] BPLlineArray = null;

                if (1 == 1)
                {
                    currentJoistSeatInfo = new JoistSeatInfo();
                    currentJoistSeatInfo.tcxL = StringManipulation.hyphenLengthToDecimal(tcxLineArray[1]);
                    currentJoistSeatInfo.tcxR = StringManipulation.hyphenLengthToDecimal(tcxLineArray[3]);
                    currentJoistSeatInfo.ota = StringManipulation.hyphenLengthToDecimal(tcxLineArray[2]);
                    currentJoistSeatInfo.oal = StringManipulation.hyphenLengthToDecimal(tcxLineArray[0]);
                    currentJoistSeatInfo.clearLeft = StringManipulation.hyphenLengthToDecimal(clearLineArray[0]);
                    currentJoistSeatInfo.clearRight = StringManipulation.hyphenLengthToDecimal(clearLineArray[1]);
                    selection.Find.Execute("BPL-L");
                    selection.HomeKey(Word.WdUnits.wdLine, 0);
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    string BPLline = selection.Text;
                    if (BPLline.Length > charsToSetback)
                    {
                        string toSetback = BPLline.Substring(0, charsToSetback);
                        string afterSetback = BPLline.Substring(charsToSetback, BPLline.Length - charsToSetback);
                        BPLline = toSetback + "  " + afterSetback;
                    }
                    if (BPLline.Contains(" S ")) { currentJoistSeatInfo.seatType = "S"; }
                    if (BPLline.Contains(" R ")) { currentJoistSeatInfo.seatType = "R"; }
                    BPLlineArray = BPLline.Split(new string[] { "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);

                    currentJoistSeatInfo.mark = mark;
                    currentJoistSeatInfo.qty = Convert.ToInt16(joistDataByMarks[i][1]);
                    currentJoistSeatInfo.TC = joistDataByMarks[i][2];
                    currentJoistSeatInfo.bplSide = BPLlineArray[1];
                    currentJoistSeatInfo.bplLength = BPLlineArray[3];
                    if (BPLlineArray[5].Contains("|") == false)
                    {
                        currentJoistSeatInfo.bplOutsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);
                    }
                    if (BPLlineArray[5].Contains("|") == true)
                    {
                        string[] bpl = BPLlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                        currentJoistSeatInfo.bplOutsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + bpl[0]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + bpl[1]);
                    }
                    if (BPLlineArray.Count() > 6)
                    {
                        currentJoistSeatInfo.slotSetback = BPLlineArray[6];      // slot setback
                        currentJoistSeatInfo.slotSize = BPLlineArray[7];      // slot size
                        currentJoistSeatInfo.slotGauge = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[8]);      // slot gauge

                        string[] bpl2 = BPLlineArray[7].Split(new string[] { " x " }, StringSplitOptions.RemoveEmptyEntries);
                        double slotLength = StringManipulation.ConvertLengthtoDecimal("0-" + bpl2[1])*12.0;
                        double slotSetback = 0.0;
                        if (BPLlineArray[6].Contains("-") == true)
                        {
                            slotSetback = StringManipulation.ConvertLengthtoDecimal(BPLlineArray[6])*12.0;
                        }
                        else
                        {
                            slotSetback = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[6])*12.0;
                        }
                        if(slotSetback<slotLength/2.0 +0.75)
                        {
                            MessageBox.Show(mark + " " + currentJoistSeatInfo.bplSide + ":\nSTIFFENER PLATE INTERFERES WITH SLOT, PLEASE CORRECT MANUALLY");
                        }
                    }

                    if (marksWithHoldClears[i].Item2 == true)
                    {
                        holdClear.Add(currentJoistSeatInfo);
                    }
                    else if (marksWithHoldClears[i].Item2 == false && BPLlineArray[2] == "BCG")
                    {
                        bcg.Add(currentJoistSeatInfo);
                    }
                    else if (marksWithHoldClears[i].Item2 == false && BPLlineArray[2] == "T-PLATE")
                    {
                        tPlate.Add(currentJoistSeatInfo);
                    }
                    else
                    {
                        standard.Add(currentJoistSeatInfo);
                    }
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                if (BPLlineArray[1] == "BPL-L1")
                {
                    currentJoistSeatInfo = new JoistSeatInfo();
                    currentJoistSeatInfo.tcxL = StringManipulation.hyphenLengthToDecimal(tcxLineArray[1]);
                    currentJoistSeatInfo.tcxR = StringManipulation.hyphenLengthToDecimal(tcxLineArray[3]);
                    currentJoistSeatInfo.ota = StringManipulation.hyphenLengthToDecimal(tcxLineArray[2]);
                    currentJoistSeatInfo.oal = StringManipulation.hyphenLengthToDecimal(tcxLineArray[0]);
                    currentJoistSeatInfo.clearLeft = StringManipulation.hyphenLengthToDecimal(clearLineArray[0]);
                    currentJoistSeatInfo.clearRight = StringManipulation.hyphenLengthToDecimal(clearLineArray[1]);
                    selection.Find.Execute("BPL-L");
                    selection.HomeKey(Word.WdUnits.wdLine, 0);
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    string BPLline = selection.Text;
                    if (BPLline.Length > charsToSetback)
                    {
                        string toSetback = BPLline.Substring(0, charsToSetback);
                        string afterSetback = BPLline.Substring(charsToSetback, BPLline.Length - charsToSetback);
                        BPLline = toSetback + "  " + afterSetback;
                    }
                    if (BPLline.Contains(" S ")) { currentJoistSeatInfo.seatType = "S"; }
                    if (BPLline.Contains(" R ")) { currentJoistSeatInfo.seatType = "R"; }
                    BPLlineArray = BPLline.Split(new string[] { "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);

                    currentJoistSeatInfo.mark = mark;
                    currentJoistSeatInfo.qty = Convert.ToInt16(joistDataByMarks[i][1]);
                    currentJoistSeatInfo.TC = joistDataByMarks[i][2];
                    currentJoistSeatInfo.bplSide = BPLlineArray[1];
                    currentJoistSeatInfo.bplLength = BPLlineArray[3];
                    if (BPLlineArray[5].Contains("|") == false)
                    {
                        currentJoistSeatInfo.bplOutsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);
                    }
                    if (BPLlineArray[5].Contains("|") == true)
                    {
                        string[] bpl = BPLlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                        currentJoistSeatInfo.bplOutsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + bpl[0]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + bpl[1]);
                    }
                    if (BPLlineArray.Count() > 6)
                    {
                        currentJoistSeatInfo.slotSetback = BPLlineArray[6];      // slot setback
                        currentJoistSeatInfo.slotSize = BPLlineArray[7];      // slot size
                        currentJoistSeatInfo.slotGauge = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[8]);      // slot gauge

                        string[] bpl2 = BPLlineArray[7].Split(new string[] { " x " }, StringSplitOptions.RemoveEmptyEntries);
                        double slotLength = StringManipulation.ConvertLengthtoDecimal("0-" + bpl2[1]) * 12.0;
                        double slotSetback = 0.0;
                        if (BPLlineArray[6].Contains("-") == true)
                        {
                            slotSetback = StringManipulation.ConvertLengthtoDecimal(BPLlineArray[6]) * 12.0;
                        }
                        else
                        {
                            slotSetback = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[6]) * 12.0;
                        }
                        if (slotSetback < slotLength / 2.0 + 0.75)
                        {
                            MessageBox.Show(mark + " " + currentJoistSeatInfo.bplSide + ":\nSTIFFENER PLATE INTERFERES WITH SLOT, PLEASE CORRECT MANUALLY");
                        }

                    }


                    if (marksWithHoldClears[i].Item2 == true)
                    {
                        holdClear.Add(currentJoistSeatInfo);
                    }
                    else if (marksWithHoldClears[i].Item2 == false && BPLlineArray[2] == "BCG")
                    {
                        bcg.Add(currentJoistSeatInfo);
                    }
                    else if (marksWithHoldClears[i].Item2 == false && BPLlineArray[2] == "T-PLATE")
                    {
                        tPlate.Add(currentJoistSeatInfo);
                    }
                    else
                    {
                        standard.Add(currentJoistSeatInfo);
                    }

                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                }
                if (1 == 1)
                {
                    currentJoistSeatInfo = new JoistSeatInfo();
                    currentJoistSeatInfo.tcxL = StringManipulation.hyphenLengthToDecimal(tcxLineArray[1]);
                    currentJoistSeatInfo.tcxR = StringManipulation.hyphenLengthToDecimal(tcxLineArray[3]);
                    currentJoistSeatInfo.ota = StringManipulation.hyphenLengthToDecimal(tcxLineArray[2]);
                    currentJoistSeatInfo.oal = StringManipulation.hyphenLengthToDecimal(tcxLineArray[0]);
                    currentJoistSeatInfo.clearLeft = StringManipulation.hyphenLengthToDecimal(clearLineArray[0]);
                    currentJoistSeatInfo.clearRight = StringManipulation.hyphenLengthToDecimal(clearLineArray[1]);
                    selection.Find.Execute("BPL-R");
                    selection.HomeKey(Word.WdUnits.wdLine, 0);
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    string BPLline = selection.Text;
                    if (BPLline.Length > charsToSetback)
                    {
                        string toSetback = BPLline.Substring(0, charsToSetback);
                        string afterSetback = BPLline.Substring(charsToSetback, BPLline.Length - charsToSetback);
                        BPLline = toSetback + "  " + afterSetback;
                    }
                    if (BPLline.Contains(" S ")) { currentJoistSeatInfo.seatType = "S"; }
                    if (BPLline.Contains(" R ")) { currentJoistSeatInfo.seatType = "R"; }
                    BPLlineArray = BPLline.Split(new string[] { "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);

                    currentJoistSeatInfo.mark = mark;
                    currentJoistSeatInfo.qty = Convert.ToInt16(joistDataByMarks[i][1]);
                    currentJoistSeatInfo.TC = joistDataByMarks[i][2];
                    currentJoistSeatInfo.bplSide = BPLlineArray[1];
                    currentJoistSeatInfo.bplLength = BPLlineArray[3];
                    if (BPLlineArray[5].Contains("|") == false)
                    {
                        currentJoistSeatInfo.bplOutsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);
                    }
                    if (BPLlineArray[5].Contains("|") == true)
                    {
                        string[] bpl = BPLlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                        currentJoistSeatInfo.bplOutsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + bpl[1]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + bpl[0]);
                    }
                    if (BPLlineArray.Count() > 6)
                    {
                        currentJoistSeatInfo.slotSetback = BPLlineArray[6];      // slot setback
                        currentJoistSeatInfo.slotSize = BPLlineArray[7];      // slot size
                        currentJoistSeatInfo.slotGauge = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[8]);      // slot gauge

                        string[] bpl2 = BPLlineArray[7].Split(new string[] { " x " }, StringSplitOptions.RemoveEmptyEntries);

                        double slotLength = bpl2[1].Contains(" ") ?
                                             StringManipulation.ConvertLengthtoDecimal("0-" + bpl2[1]) * 12.0 :
                                             StringManipulation.ConvertLengthtoDecimal("0-0 " + bpl2[1]) * 12.0;
                        double slotSetback = 0.0;
                        if (BPLlineArray[6].Contains("-") == true)
                        {
                            slotSetback = StringManipulation.ConvertLengthtoDecimal(BPLlineArray[6]) * 12.0;
                        }
                        else
                        {
                            slotSetback = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[6]) * 12.0;
                        }
                        if (slotSetback < slotLength / 2.0 + 0.75)
                        {
                            MessageBox.Show(mark + " " + currentJoistSeatInfo.bplSide + ":\nSTIFFENER PLATE INTERFERES WITH SLOT, PLEASE CORRECT MANUALLY");
                        }
                    }


                    if (marksWithHoldClears[i].Item3 == true)
                    {
                        holdClear.Add(currentJoistSeatInfo);
                    }
                    else if (marksWithHoldClears[i].Item3 == false && BPLlineArray[2] == "BCG")
                    {
                        bcg.Add(currentJoistSeatInfo);
                    }
                    else if (marksWithHoldClears[i].Item3 == false && BPLlineArray[2] == "T-PLATE")
                    {
                        tPlate.Add(currentJoistSeatInfo);
                    }
                    else
                    {
                        standard.Add(currentJoistSeatInfo);
                    }
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                if (BPLlineArray[1] == "BPL-R1")
                {
                    currentJoistSeatInfo = new JoistSeatInfo();
                    currentJoistSeatInfo.tcxL = StringManipulation.hyphenLengthToDecimal(tcxLineArray[1]);
                    currentJoistSeatInfo.tcxR = StringManipulation.hyphenLengthToDecimal(tcxLineArray[3]);
                    currentJoistSeatInfo.ota = StringManipulation.hyphenLengthToDecimal(tcxLineArray[2]);
                    currentJoistSeatInfo.oal = StringManipulation.hyphenLengthToDecimal(tcxLineArray[0]);
                    currentJoistSeatInfo.clearLeft = StringManipulation.hyphenLengthToDecimal(clearLineArray[0]);
                    currentJoistSeatInfo.clearRight = StringManipulation.hyphenLengthToDecimal(clearLineArray[1]);
                    selection.Find.Execute("BPL-R");
                    selection.HomeKey(Word.WdUnits.wdLine, 0);
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    string BPLline = selection.Text;
                    if (BPLline.Length > charsToSetback)
                    {
                        string toSetback = BPLline.Substring(0, charsToSetback);
                        string afterSetback = BPLline.Substring(charsToSetback, BPLline.Length - charsToSetback);
                        BPLline = toSetback + "  " + afterSetback;
                    }
                    if (BPLline.Contains(" S ")) { currentJoistSeatInfo.seatType = "S"; }
                    if (BPLline.Contains(" R ")) { currentJoistSeatInfo.seatType = "R"; }
                    BPLlineArray = BPLline.Split(new string[] { "              ", "             ", "            ", "           ", "          ", "         ", "        ", "       ", "      ", "     ", "    ", "   ", "  ", "\u00A0", "\u000B" }, StringSplitOptions.RemoveEmptyEntries);

                    currentJoistSeatInfo.mark = mark;
                    currentJoistSeatInfo.qty = Convert.ToInt16(joistDataByMarks[i][1]);
                    currentJoistSeatInfo.TC = joistDataByMarks[i][2];
                    currentJoistSeatInfo.bplSide = BPLlineArray[1];
                    currentJoistSeatInfo.bplLength = BPLlineArray[3];
                    if (BPLlineArray[5].Contains("|") == false)
                    {
                        currentJoistSeatInfo.bplOutsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[5]);
                    }
                    if (BPLlineArray[5].Contains("|") == true)
                    {
                        string[] bpl = BPLlineArray[5].Split(new string[] { "  ", " | ", " |", "| ", "|" }, StringSplitOptions.RemoveEmptyEntries);
                        currentJoistSeatInfo.bplOutsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + bpl[1]);
                        currentJoistSeatInfo.bplInsideDepth = StringManipulation.ConvertLengthtoDecimal("0-" + bpl[0]);
                    }
                    if (BPLlineArray.Count() > 6)
                    {
                        currentJoistSeatInfo.slotSetback = BPLlineArray[6];      // slot setback
                        currentJoistSeatInfo.slotSize = BPLlineArray[7];      // slot size
                        currentJoistSeatInfo.slotGauge = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[8]);      // slot gauge

                        string[] bpl2 = BPLlineArray[7].Split(new string[] { " x " }, StringSplitOptions.RemoveEmptyEntries);
                        double slotLength = StringManipulation.ConvertLengthtoDecimal("0-" + bpl2[1]) * 12.0;
                        double slotSetback = 0.0;
                        if (BPLlineArray[6].Contains("-") == true)
                        {
                            slotSetback = StringManipulation.ConvertLengthtoDecimal(BPLlineArray[6]) * 12.0;
                        }
                        else
                        {
                            slotSetback = StringManipulation.ConvertLengthtoDecimal("0-" + BPLlineArray[6]) * 12.0;
                        }
                        if (slotSetback < slotLength / 2.0 + 0.75)
                        {
                            MessageBox.Show(mark + " " + currentJoistSeatInfo.bplSide + ":\nSTIFFENER PLATE INTERFERES WITH SLOT, PLEASE CORRECT MANUALLY");
                        }
                    }

                    if (marksWithHoldClears[i].Item3 == true)
                    {
                        holdClear.Add(currentJoistSeatInfo);
                    }
                    else if (marksWithHoldClears[i].Item3 == false && BPLlineArray[2] == "BCG")
                    {
                        bcg.Add(currentJoistSeatInfo);
                    }
                    else if (marksWithHoldClears[i].Item3 == false && BPLlineArray[2] == "T-PLATE")
                    {
                        tPlate.Add(currentJoistSeatInfo);
                    }
                    else
                    {
                        standard.Add(currentJoistSeatInfo);
                    }
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                }

            }

            AllSeatInfo allSeatInfo = new AllSeatInfo();
            allSeatInfo.HoldClear = holdClear;
            allSeatInfo.BCG = bcg;
            allSeatInfo.Tplate = tPlate;
            allSeatInfo.Standard = standard;

            var allSeatInfoForKCheck = holdClear.Concat(bcg).Concat(tPlate).Concat(standard).ToList();

            foreach (var hcSeat in holdClear)
            {
                var oppositeSeat = allSeatInfoForKCheck.Where(s => s.mark == hcSeat.mark && s.bplSide != hcSeat.bplSide).First();
                if (hcSeat.seatType == "S")
                {
                    if (hcSeat.bplInsideDepth == oppositeSeat.bplInsideDepth)
                    {

                        if (Math.Abs(hcSeat.oal - hcSeat.tcxL - hcSeat.tcxR - hcSeat.ota) > 0.001)
                        {
                            MessageBox.Show("Mark " + hcSeat.mark + " may have a 'K' value at " + hcSeat.bplSide + "; please confirm seat configuration.");
                        }
                    }
                }
            }
            

            return allSeatInfo;

        }

        public List<List<HCSeatInfo>> listHCSeatInfo(List<JoistSeatInfo> listJoistSeatInfo)
        {
            List<HCSeatInfo> buttedHC = new List<HCSeatInfo>();
            List<HCSeatInfo> gappedHC = new List<HCSeatInfo>();
            List<HCSeatInfo> plattedHCP04 = new List<HCSeatInfo>();
            List<HCSeatInfo> plattedHCP08 = new List<HCSeatInfo>();

            int i = 0;
            foreach (JoistSeatInfo thisJoistSeatInfo in listJoistSeatInfo)
            {

                double tcVLeg = QueryAngleData.DblVleg(anglesFromSql, thisJoistSeatInfo.TC);
                double tcThickness = QueryAngleData.DblThickness(anglesFromSql, thisJoistSeatInfo.TC);

                HCSeatInfo thisHCSeatInfo = new HCSeatInfo();
                thisHCSeatInfo.mark = thisJoistSeatInfo.mark;
                thisHCSeatInfo.qty = thisJoistSeatInfo.qty * 2;
                thisHCSeatInfo.TC = thisJoistSeatInfo.TC;
                thisHCSeatInfo.seatType = thisJoistSeatInfo.seatType;
                thisHCSeatInfo.bplSide = thisJoistSeatInfo.bplSide;
                thisHCSeatInfo.bplLength = StringManipulation.cleanDecimalToHyphen(StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.bplLength));
                thisHCSeatInfo.bplOutsideDepth = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth);
                thisHCSeatInfo.bplInsideDepth = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth);

                if (tcVLeg <= 3.25)
                {
                    thisHCSeatInfo.stiffPlateLength = StringManipulation.cleanDecimalToHyphen(StringManipulation.NearestQuarterInch(thisJoistSeatInfo.bplOutsideDepth - 1.0 / 12.0));
                }
                else
                {
                    thisHCSeatInfo.stiffPlateLength = StringManipulation.cleanDecimalToHyphen(StringManipulation.NearestHalfInch(thisJoistSeatInfo.bplOutsideDepth - 2.0 / 12.0));
                }
                thisHCSeatInfo.paMat = "P0604";

                
                double slopeFactor = Math.Sin((Math.PI*90.0/180.0)-Math.Atan((Math.Abs((thisJoistSeatInfo.bplInsideDepth-thisJoistSeatInfo.bplOutsideDepth)/StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.bplLength)))));
                slopeFactor=slopeFactor+(Math.Sqrt(1.0-Math.Pow(slopeFactor,2.0))*(Math.Abs((thisJoistSeatInfo.bplInsideDepth-thisJoistSeatInfo.bplOutsideDepth)/StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.bplLength))));

                
                if (thisJoistSeatInfo.slotGauge >= 4.0 / 12.0)
                {
                    thisHCSeatInfo.paWidth = "3";
                }
                else { thisHCSeatInfo.paWidth = "2 1/2"; }

                if (thisJoistSeatInfo.slotSetback != null)
                {
                    thisHCSeatInfo.slotSetback = StringManipulation.cleanDecimalToHyphen(StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.slotSetback));
                    thisHCSeatInfo.slotSize = thisJoistSeatInfo.slotSize;
                }

                int diffInHCandTC = 0;

   

                bool gaIsGreater3pt5 = false;
                if (thisJoistSeatInfo.slotSetback != null)
                {
                    if (thisJoistSeatInfo.slotGauge*12.0 > 3.50)
                    {
                        gaIsGreater3pt5 = true;
                    }
                }
                if (thisJoistSeatInfo.TC.Contains("A") == true && gaIsGreater3pt5 == false)
                {

                    diffInHCandTC = 0;
                    if (tcThickness < 0.1495)
                    {
                        thisHCSeatInfo.HCMaterial = "1714";
                        if (thisJoistSeatInfo.mark == "A12B")
                        {
                            diffInHCandTC = -1;
                        }

                    }
                    else if (tcThickness < 0.172) { thisHCSeatInfo.HCMaterial = "2015"; }
                    else if (tcThickness < 0.20) { thisHCSeatInfo.HCMaterial = "2018"; }
                    else { thisHCSeatInfo.HCMaterial = "2024"; }

                    double seatVLeg = QueryAngleData.DblVleg(anglesFromSql, thisHCSeatInfo.HCMaterial);
                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * tcVLeg)+ seatVLeg && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * tcVLeg) + seatVLeg)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * tcVLeg) / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * tcVLeg) / 12.0);
                    }
              /*      else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * tcVLeg) >= seatVLeg - 0.1251 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * tcVLeg) >= seatVLeg - 0.1251)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP04 = true;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * tcVLeg) / 12.0 - 0.25/12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * tcVLeg) / 12.0 - 0.25/12.0);

                    }
                }
                if (thisJoistSeatInfo.TC.Contains("A") == true && gaIsGreater3pt5 == true)
                {
                    thisHCSeatInfo.HCMaterial = "3025";
                    if (tcThickness < 0.15625) { diffInHCandTC = 2; }
                    else if (tcThickness < 0.21875) { diffInHCandTC = 1; }
                    else { diffInHCandTC = 0; }

                    double seatVLeg = QueryAngleData.DblVleg(anglesFromSql, thisHCSeatInfo.HCMaterial);
                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor*tcVLeg) + seatVLeg && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor*tcVLeg) + seatVLeg)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * tcVLeg) / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * tcVLeg) / 12.0);
                    }
                 /*   else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * tcVLeg) >= seatVLeg - 0.1251 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * tcVLeg) >= seatVLeg - 0.1251)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP04 = true;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * tcVLeg) / 12.0 - 0.25 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * tcVLeg) / 12.0 - 0.25 / 12.0);

                    }
                }

                if (thisJoistSeatInfo.TC == "3025" || thisJoistSeatInfo.TC == "3028" || thisJoistSeatInfo.TC == "3031")
                {

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 3.0) + 3.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 3.0) + 3.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                        if (thisJoistSeatInfo.TC == "3025")
                        {
                            thisHCSeatInfo.HCMaterial = "3025";
                            diffInHCandTC = 0;
                        }
                        if (thisJoistSeatInfo.TC == "3028")
                        {
                            thisHCSeatInfo.HCMaterial = "3028";
                            diffInHCandTC = 0;
                        }
                        if (thisJoistSeatInfo.TC == "3031")
                        {
                            thisHCSeatInfo.HCMaterial = "3031";
                            diffInHCandTC = 0;
                        }
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 3.0) / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 3.0) / 12.0);
                    }
                    else if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 3.0) + 3.501 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 3.0) + 3.501)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                        thisHCSeatInfo.HCMaterial = "3528";
                        if (thisJoistSeatInfo.TC == "3025") { diffInHCandTC = 1; }
                        if (thisJoistSeatInfo.TC == "3028") { diffInHCandTC = 0; }
                        if (thisJoistSeatInfo.TC == "3031") { diffInHCandTC = -1; }
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 3.0) / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 3.0) / 12.0);
                    }
                    else
                    {
                        thisHCSeatInfo.HCMaterial = thisJoistSeatInfo.TC;
                        if (thisJoistSeatInfo.TC == "3025") { diffInHCandTC = 0; }
                        if (thisJoistSeatInfo.TC == "3028") { diffInHCandTC = 0; }
                        if (thisJoistSeatInfo.TC == "3031") { diffInHCandTC = -1; }
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*3.0) / 12.0 - 0.25 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*3.0) / 12.0 - 0.25 / 12.0);
                        thisHCSeatInfo.plateSeatP04 = true;
                    }

                }

                if (thisJoistSeatInfo.TC == "3528")
                {
                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 3.5) + 3.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 3.5) + 3.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                        thisHCSeatInfo.HCMaterial = "3028";
                        diffInHCandTC = 0;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 3.5) / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 3.5) / 12.0);
                    }
                    else if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 3.5) + 3.501 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 3.5) + 3.501)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                        thisHCSeatInfo.HCMaterial = "3528";
                        diffInHCandTC = 0;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 3.5) / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 3.5) / 12.0);
                    }

                    else
                    {
                        thisHCSeatInfo.HCMaterial = thisJoistSeatInfo.TC;
                        diffInHCandTC = 0;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*3.5) / 12.0 - 0.25 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*3.5) / 12.0 - 0.25 / 12.0);
                        thisHCSeatInfo.plateSeatP04 = true;
                    }

                }

                if (thisJoistSeatInfo.TC == "3531" || thisJoistSeatInfo.TC == "3534")
                {

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 3.5) + 3.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 3.5) + 3.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                        thisHCSeatInfo.HCMaterial = "3031";
                        if (thisJoistSeatInfo.TC == "3531") { diffInHCandTC = 0; }
                        if (thisJoistSeatInfo.TC == "3534") { diffInHCandTC = -1; }
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 3.5) / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 3.5) / 12.0);
                    }
                    else
                    {
                        thisHCSeatInfo.plateSeatP04 = true;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 3.5) / 12.0 - 0.25 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 3.5) / 12.0 - 0.25 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        diffInHCandTC = -1;

                    }



                }

                if (thisJoistSeatInfo.TC == "4037")
                {
                    diffInHCandTC = -1;
                    thisHCSeatInfo.HCMaterial = "3534";

                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*4.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*4.0) / 12.0);

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 4.0) + 3.501 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 4.0) + 3.501)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
              /*      else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 4.0) >= 3.501 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 4.0) >= 3.501 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP04 = true;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 4.0) / 12.0 - 0.25 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 4.0) / 12.0 - 0.25 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        diffInHCandTC = -2;

                    }
                }

                if (thisJoistSeatInfo.TC == "4043")
                {
                    diffInHCandTC = -1;
                    thisHCSeatInfo.HCMaterial = "3534";

                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 4.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 4.0) / 12.0);

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 4.0) + 3.501 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 4.0) + 3.501)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
                    /*      else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 4.0) >= 3.501 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 4.0) >= 3.501 - 0.125)
                          {
                              thisHCSeatInfo.gappedSeat = true;
                          } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 4.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 4.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        diffInHCandTC = 1;
                        thisHCSeatInfo.paWidth = "4";

                    }
                }


                if (thisJoistSeatInfo.TC == "4050")
                {
                    diffInHCandTC = 0;
                    thisHCSeatInfo.HCMaterial = "4050";
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*4.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*4.0) / 12.0);

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 4.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 4.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
                 /*   else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 4.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 4.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 4.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 4.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }
                }
                if (thisJoistSeatInfo.TC == "5043")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 5.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 5.0) / 12.0);

                    thisHCSeatInfo.HCMaterial = "4043";
                    diffInHCandTC = 0;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 5.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 4.0) + 5.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
               /*     else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 5.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 5.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        diffInHCandTC = 1;
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 5.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 5.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }

                }
                if (thisJoistSeatInfo.TC == "5050")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*5.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*5.0) / 12.0);

                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = 0;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 5.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 5.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
                /*    else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 5.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 5.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 5.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 5.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }
                }
                if (thisJoistSeatInfo.TC == "6050")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*6.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*6.0) / 12.0);

                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = 0;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 6.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 6.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
            /*        else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }
                }
                if (thisJoistSeatInfo.TC == "6056")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*6.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*6.0) / 12.0);
                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = -1;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 6.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 6.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
                 /*   else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5/ 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }

                }
                if (thisJoistSeatInfo.TC == "6062")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*6.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*6.0) / 12.0);
                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = -2;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 6.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 6.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
               /*     else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }

                }
                if (thisJoistSeatInfo.TC == "6075")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*6.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*6.0) / 12.0);
                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = -4;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 6.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 6.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
           /*         else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }

                }
                if (thisJoistSeatInfo.TC == "6010")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*6.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*6.0) / 12.0);
                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = -8;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 6.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 6.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
              /*      else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 6.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }

                }
                if (thisJoistSeatInfo.TC == "8062")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*8.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*8.0) / 12.0);
                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = -2;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 8.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 8.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
             /*       else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 8.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 8.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 8.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 8.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }

                }
                if (thisJoistSeatInfo.TC == "8075")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*8.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*8.0) / 12.0);
                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = -4;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 8.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 8.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
           /*         else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 8.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 8.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 8.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 8.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }

                }
                if (thisJoistSeatInfo.TC == "8010")
                {
                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor*8.0) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor*8.0) / 12.0);
                    thisHCSeatInfo.HCMaterial = "4050";
                    diffInHCandTC = -8;

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * 8.0) + 4.001 && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * 8.0) + 4.001)
                    {
                        thisHCSeatInfo.buttedSeat = true;
                    }
           /*         else if (12.0 * thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 8.0) >= 4.001 - 0.125 && 12.0 * thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 8.0) >= 4.001 - 0.125)
                    {
                        thisHCSeatInfo.gappedSeat = true;
                    } */
                    else
                    {
                        thisHCSeatInfo.plateSeatP08 = true;
                        thisHCSeatInfo.paMat = "P0608";
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * 8.0) / 12.0 - 0.5 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * 8.0) / 12.0 - 0.5 / 12.0);
                        if (thisJoistSeatInfo.bplInsideDepth > thisJoistSeatInfo.bplOutsideDepth)
                        {
                            MessageBox.Show(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                        }
                        thisHCSeatInfo.paWidth = "4";

                    }

                }

                if (thisJoistSeatInfo.slotGauge != 0)
                {
                    int intGOL = Convert.ToInt16((thisJoistSeatInfo.slotGauge - (1.0 / 12.0)) / 2) * 12 * 16 + diffInHCandTC;
                    double dblGOL = ((thisJoistSeatInfo.slotGauge - (1.0 / 12.0)) / 2.0) + Convert.ToDouble(diffInHCandTC) / (12 * 16.0);
                    thisHCSeatInfo.GOL = StringManipulation.cleanDecimalToHyphen16ths(dblGOL);
                }
                //////////

                string overiddenSeatType = Convert.ToString(overideDetails.Select(String.Format("Mark='{0}'", thisJoistSeatInfo.mark))[0][1]);

                if (overiddenSeatType== "") { }
                else if (overiddenSeatType == "Butted") { thisHCSeatInfo.buttedSeat = true; thisHCSeatInfo.gappedSeat = false; thisHCSeatInfo.plateSeatP04 = false; thisHCSeatInfo.plateSeatP08 = false; }
                else if (overiddenSeatType == "Gapped") { thisHCSeatInfo.buttedSeat = false; thisHCSeatInfo.gappedSeat = true; thisHCSeatInfo.plateSeatP04 = false; thisHCSeatInfo.plateSeatP08 = false; }
                else if (overiddenSeatType == "1/4\" Plate") { thisHCSeatInfo.buttedSeat = false; thisHCSeatInfo.gappedSeat = false; thisHCSeatInfo.plateSeatP04 = true; thisHCSeatInfo.plateSeatP08 = false; }
                else if (overiddenSeatType == "1/2\" Plate") { thisHCSeatInfo.buttedSeat = false; thisHCSeatInfo.gappedSeat = false; thisHCSeatInfo.plateSeatP04 = false; thisHCSeatInfo.plateSeatP08 = true; thisHCSeatInfo.paMat = "P0608"; }
                else { }
                /////////////
                if (thisHCSeatInfo.gappedSeat == true)
                {
                    gappedHC.Add(thisHCSeatInfo);
                    if (thisJoistSeatInfo.bplSide == "BPL-L")
                    {
                        thisHCSeatInfo.bplLength = StringManipulation.DecimilLengthToHyphen(Math.Max(thisJoistSeatInfo.clearLeft - (thisJoistSeatInfo.tcxL / slopeFactor) - 0.75 / 12, StringManipulation.hyphenLengthToDecimal(thisJoistSeatInfo.bplLength)));
                    }
                    if (thisJoistSeatInfo.bplSide == "BPL-R")
                    {
                        thisHCSeatInfo.bplLength = StringManipulation.DecimilLengthToHyphen(Math.Max(thisJoistSeatInfo.clearRight - (thisJoistSeatInfo.tcxR / slopeFactor) - 0.75 / 12, StringManipulation.hyphenLengthToDecimal(thisJoistSeatInfo.bplLength)));
                    }
                }

                if (thisHCSeatInfo.gappedSeat == true)
                {
                    string newSlopeString = "";
                    if (thisHCSeatInfo.bplOutsideDepth != thisHCSeatInfo.bplInsideDepth)
                    {
                        double seatLength = 12.0 * StringManipulation.hyphenLengthToDecimal(thisHCSeatInfo.bplLength);
                        double insideDepth = 12.0 * StringManipulation.hyphenLengthToDecimal(thisHCSeatInfo.bplInsideDepth);
                        double outsideDepth = 12.0 * StringManipulation.hyphenLengthToDecimal(thisHCSeatInfo.bplOutsideDepth);
                        double addSlopeDepthInches = (seatLength * (insideDepth - outsideDepth)) / 6.0;
                        double newInsideDepth = (outsideDepth + addSlopeDepthInches) / 12.0;
                        thisHCSeatInfo.bplInsideDepth = StringManipulation.cleanDecimalToHyphen(newInsideDepth);

                    }
                }


                if (thisHCSeatInfo.buttedSeat == true) { buttedHC.Add(thisHCSeatInfo); }
                if (thisHCSeatInfo.plateSeatP04 == true) { plattedHCP04.Add(thisHCSeatInfo); }
                if (thisHCSeatInfo.plateSeatP08 == true) { plattedHCP08.Add(thisHCSeatInfo); }

                i++;
            }
            List<List<HCSeatInfo>> listHCSeatInfo = new List<List<HCSeatInfo>>();
            listHCSeatInfo.Add(buttedHC);
            listHCSeatInfo.Add(gappedHC);
            listHCSeatInfo.Add(plattedHCP04);
            listHCSeatInfo.Add(plattedHCP08);
            return listHCSeatInfo;
        }

        public List<List<HCSeatInfo>> listHCSeatInfo2(List<JoistSeatInfo> listJoistSeatInfo, bool showMessages)
        {
            List<HCSeatInfo> buttedHC = new List<HCSeatInfo>();
            List<HCSeatInfo> gappedHC = new List<HCSeatInfo>();
            List<HCSeatInfo> plattedHCP04 = new List<HCSeatInfo>();
            List<HCSeatInfo> plattedHCP08 = new List<HCSeatInfo>();

            var messages = new List<string>();

            int i = 0;
            foreach (JoistSeatInfo thisJoistSeatInfo in listJoistSeatInfo)
            {

                double tcVLeg = QueryAngleData.DblVleg(anglesFromSql, thisJoistSeatInfo.TC);
                double tcThickness = QueryAngleData.DblThickness(anglesFromSql, thisJoistSeatInfo.TC);
                double tcHLeg = QueryAngleData.DblHleg(anglesFromSql, thisJoistSeatInfo.TC);

                HCSeatInfo thisHCSeatInfo = new HCSeatInfo();
                thisHCSeatInfo.mark = thisJoistSeatInfo.mark;
                thisHCSeatInfo.qty = thisJoistSeatInfo.qty * 2;
                thisHCSeatInfo.TC = thisJoistSeatInfo.TC;
                thisHCSeatInfo.seatType = thisJoistSeatInfo.seatType;
                thisHCSeatInfo.bplSide = thisJoistSeatInfo.bplSide;
                thisHCSeatInfo.bplLength = StringManipulation.cleanDecimalToHyphen(StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.bplLength));
                thisHCSeatInfo.bplOutsideDepth = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth);
                thisHCSeatInfo.bplInsideDepth = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth);

                if (tcVLeg <= 3.25)
                {
                    thisHCSeatInfo.stiffPlateLength = StringManipulation.cleanDecimalToHyphen(StringManipulation.NearestQuarterInch(thisJoistSeatInfo.bplOutsideDepth - 1.0 / 12.0));
                }
                else
                {
                    thisHCSeatInfo.stiffPlateLength = StringManipulation.cleanDecimalToHyphen(StringManipulation.NearestHalfInch(thisJoistSeatInfo.bplOutsideDepth - 2.0 / 12.0));
                }
                thisHCSeatInfo.paMat = "P0604";


                double slopeFactor = Math.Sin((Math.PI * 90.0 / 180.0) - Math.Atan((Math.Abs((thisJoistSeatInfo.bplInsideDepth - thisJoistSeatInfo.bplOutsideDepth) / StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.bplLength)))));
                slopeFactor = slopeFactor + (Math.Sqrt(1.0 - Math.Pow(slopeFactor, 2.0)) * (Math.Abs((thisJoistSeatInfo.bplInsideDepth - thisJoistSeatInfo.bplOutsideDepth) / StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.bplLength))));


                if (thisJoistSeatInfo.slotSetback != null)
                {
                    thisHCSeatInfo.slotSetback = StringManipulation.cleanDecimalToHyphen(StringManipulation.ConvertLengthtoDecimal(thisJoistSeatInfo.slotSetback));
                    thisHCSeatInfo.slotSize = thisJoistSeatInfo.slotSize;
                }

                int diffInHCandTC = 0;



                bool gaIsGreater3pt5 = false;
                if (thisJoistSeatInfo.slotSetback != null)
                {
                    if (thisJoistSeatInfo.slotGauge * 12.0 > 3.50)
                    {
                        gaIsGreater3pt5 = true;
                    }
                }

                bool plateDoesntWork = false;
                string secondButtedOption = "";
                string plateSeat = "";
                bool throwOverlapMessage = false;

                if (tcThickness < 0.1495 && !gaIsGreater3pt5) { thisHCSeatInfo.HCMaterial = "1714"; }
                if (tcThickness < 0.1495 && gaIsGreater3pt5) { thisHCSeatInfo.HCMaterial = "2521"; }
                if (tcThickness < 0.172 && !gaIsGreater3pt5) { thisHCSeatInfo.HCMaterial = "2015"; }
                if (tcThickness < 0.172 && gaIsGreater3pt5) { thisHCSeatInfo.HCMaterial = "2521"; }
                if (tcThickness < 0.20 && !gaIsGreater3pt5) { thisHCSeatInfo.HCMaterial = "2018"; plateSeat = "P0604"; }
                if (tcThickness < 0.20 && gaIsGreater3pt5) { thisHCSeatInfo.HCMaterial = "3025"; plateSeat = "P0604"; }
                if (tcThickness >= 0.20 && tcThickness < 0.25 && !gaIsGreater3pt5) { thisHCSeatInfo.HCMaterial = "2024"; plateSeat = "P0604"; }
                if (tcThickness >= 0.20 && tcThickness < 0.25 && gaIsGreater3pt5) { thisHCSeatInfo.HCMaterial = "3025"; plateSeat = "P0604"; }
                if (tcThickness >= 0.25 && tcThickness < 0.25 + 1.0/64.0) { thisHCSeatInfo.HCMaterial = "3025"; secondButtedOption = "3528"; plateSeat = "P0604"; }
                if (tcThickness >= 9.0 / 32.0 - 1.0 / 64.0 && tcThickness < 9.0 / 32.0 + 1.0 / 64.0) { thisHCSeatInfo.HCMaterial = "3028"; secondButtedOption = "3528"; plateSeat = "P0604"; }
                if (tcThickness >= 5.0 / 16.0 - 1.0 / 64.0 && tcThickness < 11.0 / 32.0 + 1.0 / 64.0) { thisHCSeatInfo.HCMaterial = "3031"; secondButtedOption = "3528"; plateSeat = "P0604"; }
                if (tcThickness >= 3.0 / 8.0 - 1.0 / 64.0 && tcThickness < 3.0 / 8.0 + 1.0 / 64.0 && tcHLeg < 3.0) { thisHCSeatInfo.HCMaterial = "3031"; plateSeat = "P0604"; }
                if (tcThickness >= 3.0 / 8.0 - 1.0 / 64.0 && tcThickness < 3.0 / 8.0 + 1.0 / 64.0 && tcHLeg >= 3.0) { thisHCSeatInfo.HCMaterial = "3534"; plateSeat = "P0604"; throwOverlapMessage = true; }
                if (tcThickness >= 7.0 / 16.0 - 1.0 / 64.0 && tcThickness < 7.0 / 16.0 + 1.0 / 64.0 && tcHLeg < 5.0) { thisHCSeatInfo.HCMaterial = "3534"; secondButtedOption = "4043"; plateSeat = "P0608"; throwOverlapMessage = true; }
                if (tcThickness >= 7.0 / 16.0 - 1.0 / 64.0 && tcThickness < 7.0 / 16.0 + 1.0 / 64.0 && tcHLeg >= 5.0) { thisHCSeatInfo.HCMaterial = "4043"; plateSeat = "P0608"; throwOverlapMessage = true; }
                if (tcThickness >= 0.50 - 1.0 / 64.0) { thisHCSeatInfo.HCMaterial = "4050"; }

                double seatVLeg = QueryAngleData.DblVleg(anglesFromSql, thisHCSeatInfo.HCMaterial);
                double seatThickness = QueryAngleData.DblThickness(anglesFromSql, thisHCSeatInfo.HCMaterial);

                bool option1Works = false;

                if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * tcVLeg) + seatVLeg && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * tcVLeg) + seatVLeg)
                {
                    option1Works = true;
                    thisHCSeatInfo.buttedSeat = true;

                    thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * tcVLeg) / 12.0);
                    thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * tcVLeg) / 12.0);
                    diffInHCandTC = Convert.ToInt32(Math.Round((seatThickness - tcThickness) * 16.0));
                }

                bool secondButtedOptionWorks = false;

                if (option1Works == false && secondButtedOption != "")
                {
                    thisHCSeatInfo.HCMaterial = secondButtedOption;
                    seatVLeg = QueryAngleData.DblVleg(anglesFromSql, thisHCSeatInfo.HCMaterial);
                    seatThickness = QueryAngleData.DblThickness(anglesFromSql, thisHCSeatInfo.HCMaterial);

                    if (12.0 * thisJoistSeatInfo.bplOutsideDepth <= (slopeFactor * tcVLeg) + seatVLeg && 12.0 * thisJoistSeatInfo.bplInsideDepth <= (slopeFactor * tcVLeg) + seatVLeg)
                    {
                        secondButtedOptionWorks = true;
                        thisHCSeatInfo.buttedSeat = true;

                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * tcVLeg) / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * tcVLeg) / 12.0);
                        diffInHCandTC = Convert.ToInt32(Math.Round((seatThickness - tcThickness) * 16.0));
                    }

                }

                if (option1Works == false && secondButtedOptionWorks == false)
                {
                    if (plateDoesntWork)
                    {

                        messages.Add(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n" + "TC MATERIAL IS TOO THIN TO ACCEPT THE NECESSARY SEAT.\r\nPLEASE INCREASE THE TC TO A SIZE THAT IS AT LEAST 0.20\" THICK");
                    }
                    else
                    {
                        if (plateSeat == "P0604")
                        {
                            thisHCSeatInfo.plateSeatP04 = true;
                            seatThickness = 0.25;
                        }

                        if (plateSeat == "P0608")
                        {
                            thisHCSeatInfo.plateSeatP08 = true;
                            seatThickness = 0.50;
                        }
                        thisHCSeatInfo.HCOutsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplOutsideDepth - (slopeFactor * tcVLeg) / 12.0 - 0.25 / 12.0);
                        thisHCSeatInfo.HCInsideHeight = StringManipulation.cleanDecimalToHyphen(thisJoistSeatInfo.bplInsideDepth - (slopeFactor * tcVLeg) / 12.0 - 0.25 / 12.0);
                        diffInHCandTC = Convert.ToInt32(Math.Round((0.25 - tcThickness) * 16.0));
                            
                    }
                }

                if (throwOverlapMessage == true)
                {
                    messages.Add(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n" + "CONFIRM THAT W2 OVERLAPS BASEPLATE BY A MINIMUM OF 2\".\r\nIF NOT, INCREASE THE LENGTH OF THE PLATE AND ADJUST ALL DEPTHS ACCORDINGLY.");
                }


                if (thisJoistSeatInfo.slotGauge != 0)
                {
                    int intGOL = Convert.ToInt16((thisJoistSeatInfo.slotGauge - (1.0 / 12.0)) / 2) * 12 * 16 + diffInHCandTC;
                    double dblGOL = ((thisJoistSeatInfo.slotGauge - (1.0 / 12.0)) / 2.0) + Convert.ToDouble(diffInHCandTC) / (12 * 16.0);
                    thisHCSeatInfo.GOL = StringManipulation.cleanDecimalToHyphen16ths(dblGOL);
                }
                //////////

                string overiddenSeatType = Convert.ToString(overideDetails.Select(String.Format("Mark='{0}'", thisJoistSeatInfo.mark))[0][1]);

                if (overiddenSeatType == "") { }
                else if (overiddenSeatType == "Butted") { thisHCSeatInfo.buttedSeat = true; thisHCSeatInfo.gappedSeat = false; thisHCSeatInfo.plateSeatP04 = false; thisHCSeatInfo.plateSeatP08 = false; }
                else if (overiddenSeatType == "Gapped") { thisHCSeatInfo.buttedSeat = false; thisHCSeatInfo.gappedSeat = true; thisHCSeatInfo.plateSeatP04 = false; thisHCSeatInfo.plateSeatP08 = false; }
                else if (overiddenSeatType == "1/4\" Plate") { thisHCSeatInfo.buttedSeat = false; thisHCSeatInfo.gappedSeat = false; thisHCSeatInfo.plateSeatP04 = true; thisHCSeatInfo.plateSeatP08 = false; }
                else if (overiddenSeatType == "1/2\" Plate") { thisHCSeatInfo.buttedSeat = false; thisHCSeatInfo.gappedSeat = false; thisHCSeatInfo.plateSeatP04 = false; thisHCSeatInfo.plateSeatP08 = true; thisHCSeatInfo.paMat = "P0608"; }
                else { }
                /////////////
                if (thisHCSeatInfo.gappedSeat == true)
                {
                    gappedHC.Add(thisHCSeatInfo);
                    if (thisJoistSeatInfo.bplSide == "BPL-L")
                    {
                        thisHCSeatInfo.bplLength = StringManipulation.DecimilLengthToHyphen(Math.Max(thisJoistSeatInfo.clearLeft - (thisJoistSeatInfo.tcxL / slopeFactor) - 0.75 / 12, StringManipulation.hyphenLengthToDecimal(thisJoistSeatInfo.bplLength)));
                    }
                    if (thisJoistSeatInfo.bplSide == "BPL-R")
                    {
                        thisHCSeatInfo.bplLength = StringManipulation.DecimilLengthToHyphen(Math.Max(thisJoistSeatInfo.clearRight - (thisJoistSeatInfo.tcxR / slopeFactor) - 0.75 / 12, StringManipulation.hyphenLengthToDecimal(thisJoistSeatInfo.bplLength)));
                    }
                }

                if (thisHCSeatInfo.gappedSeat == true)
                {
                    string newSlopeString = "";
                    if (thisHCSeatInfo.bplOutsideDepth != thisHCSeatInfo.bplInsideDepth)
                    {
                        double seatLength = 12.0 * StringManipulation.hyphenLengthToDecimal(thisHCSeatInfo.bplLength);
                        double insideDepth = 12.0 * StringManipulation.hyphenLengthToDecimal(thisHCSeatInfo.bplInsideDepth);
                        double outsideDepth = 12.0 * StringManipulation.hyphenLengthToDecimal(thisHCSeatInfo.bplOutsideDepth);
                        double addSlopeDepthInches = (seatLength * (insideDepth - outsideDepth)) / 6.0;
                        double newInsideDepth = (outsideDepth + addSlopeDepthInches) / 12.0;
                        thisHCSeatInfo.bplInsideDepth = StringManipulation.cleanDecimalToHyphen(newInsideDepth);

                    }
                }

                var minSeatHorizontalLeg =
                    thisJoistSeatInfo.slotGauge * 12.0 * 0.5 + thisJoistSeatInfo.slotDiameter * 0.5 + 0.25 - 0.5 - tcThickness + seatThickness;

                var paWidth = Math.Max(tcHLeg > 3 ? 3 : 2.5, Math.Ceiling(minSeatHorizontalLeg * 2.0) / 2.0);

                var paWidthString = StringManipulation.decimalInchestoFraction(paWidth/12.0);

                thisHCSeatInfo.paWidth = paWidthString;


                if (thisHCSeatInfo.buttedSeat == true)
                {
                    double seatHLeg = QueryAngleData.DblHleg(anglesFromSql, thisHCSeatInfo.HCMaterial);

                    if (seatHLeg < minSeatHorizontalLeg - 0.00000001)
                    {
                        messages.Add(thisHCSeatInfo.mark + " " + thisHCSeatInfo.bplSide + " :\r\n" + "SELECTED SEAT DOES NOT MAINTAIN 1/4\" OFFSET FROM SLOT; PLEASE ADJUST HC SEAT DESIGN ACCORDINGLY.");
                    }

                }

                if (thisHCSeatInfo.buttedSeat == true) { buttedHC.Add(thisHCSeatInfo); }
                if (thisHCSeatInfo.plateSeatP04 == true) { plattedHCP04.Add(thisHCSeatInfo); }
                if (thisHCSeatInfo.plateSeatP08 == true) { plattedHCP08.Add(thisHCSeatInfo); }

                i++;
            }

            if(showMessages && messages.Any())
            {
                var message = String.Join("\r\n\r\n", messages);
                MessageBox.Show(message);
            }
            List<List<HCSeatInfo>> listHCSeatInfo = new List<List<HCSeatInfo>>();
            listHCSeatInfo.Add(buttedHC);
            listHCSeatInfo.Add(gappedHC);
            listHCSeatInfo.Add(plattedHCP04);
            listHCSeatInfo.Add(plattedHCP08);

            return listHCSeatInfo;
        }

        public List<List<HCSeatInfo>> organizedHCSeatInfoList(List<List<HCSeatInfo>> listHCSeatInfo)
        {
            List<List<HCSeatInfo>> allSeatLists = new List<List<HCSeatInfo>>();
            for (int i = 0; i <= 3; i++)
            {
                List<HCSeatInfo> seatList = listHCSeatInfo[i];

                for (int index = 0; index < seatList.Count(); index++)
                {
                    if (seatList[index].bplSide.Contains("BPL-L") == true)
                    {
                        seatList[index].mark = seatList[index].mark + "\u00A0" + "LE";
                    }
                    if (seatList[index].bplSide.Contains("BPL-R") == true)
                    {
                        seatList[index].mark = seatList[index].mark + "\u00A0" + "RE";
                    }
                }

                for (int index = 1; index < seatList.Count(); index++)
                {
                    if (
                            seatList[index].mark.Substring(0, seatList[index].mark.Length - 3) == seatList[index - 1].mark.Substring(0, seatList[index - 1].mark.Length - 3) &&
                            seatList[index].HCMaterial == seatList[index - 1].HCMaterial &&
                            seatList[index].bplLength == seatList[index - 1].bplLength &&
                            seatList[index].HCInsideHeight == seatList[index - 1].HCInsideHeight &&
                            seatList[index].HCOutsideHeight == seatList[index - 1].HCOutsideHeight &&
                            seatList[index].slotSetback == seatList[index - 1].slotSetback &&
                            seatList[index].GOL == seatList[index - 1].GOL &&
                            seatList[index].slotSize == seatList[index - 1].slotSize &&
                            seatList[index].bplOutsideDepth == seatList[index - 1].bplInsideDepth &&
                            seatList[index].stiffPlateLength == seatList[index - 1].stiffPlateLength &&
                            seatList[index].paWidth == seatList[index - 1].paWidth
                        )
                    {
                        seatList[index - 1].mark = seatList[index - 1].mark.Substring(0, seatList[index - 1].mark.Length - 3) + "\u00A0" + "BE";
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
                            if (
                                seatList[index].HCMaterial == seatList[index2].HCMaterial &&
                                seatList[index].bplLength == seatList[index2].bplLength &&
                                seatList[index].HCInsideHeight == seatList[index2].HCInsideHeight &&
                                seatList[index].HCOutsideHeight == seatList[index2].HCOutsideHeight &&
                                seatList[index].slotSetback == seatList[index2].slotSetback &&
                                seatList[index].GOL == seatList[index2].GOL &&
                                seatList[index].slotSize == seatList[index2].slotSize &&
                                seatList[index].bplOutsideDepth == seatList[index2].bplInsideDepth &&
                                seatList[index].stiffPlateLength == seatList[index2].stiffPlateLength &&
                                seatList[index].paWidth == seatList[index2].paWidth
                                )

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
                allSeatLists.Add(seatList);
            }

            return allSeatLists;
        }

        public void placeHCs(List<List<string>> joistDataByMarks, List<List<HCSeatInfo>> HCSeatInfoList)
        {
            for (int k = 0; k <= 3; k++)
            {
                List<HCSeatInfo> HCSeatInfo = HCSeatInfoList[k];

                for (int i = 0; i < HCSeatInfo.Count; i++)
                {
                    Word.Range range = Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0);

                    range.Find.Execute("COLOR CODE");
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Find.Execute(HCSeatInfo[i].mark + "    ");
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Find.Execute("QTY");
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Move(Word.WdUnits.wdCharacter, -1);
                    int lineStartInt = range.Start;
                    ///GET SECTION CHARACTERS////
                    range.Find.Execute("DESC");
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    // range.MoveStart(Word.WdUnits.wdLine,1)
                    range.Start = lineStartInt;

                    int charsToDesc = range.Text.Length;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Find.Execute("Section");
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    // range.MoveStart(Word.WdUnits.wdLine,1)
                    range.Start = lineStartInt;
                    int charsToSectionStart = range.Text.Length;
                    int charsInDesc = charsToSectionStart - charsToDesc;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Find.Execute("LENGTH");
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Start = lineStartInt;
                    //range.Move(Word.WdUnits.wdLine,0);
                    int charsToLengthStart = range.Text.Length;
                    int charsInSection = charsToLengthStart - charsToSectionStart;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Find.Execute("TYPE");
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Start = lineStartInt;
                    //range.Move(Word.WdUnits.wdLine, 0);
                    int charsToTypeStart = range.Text.Length;
                    int charsInLength = charsToTypeStart - charsToLengthStart;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Find.Execute("DEPTH");
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Start = lineStartInt;
                    //range.Move(Word.WdUnits.wdLine, 0);
                    int charsToDepthStart = range.Text.Length;
                    int charsInType = charsToDepthStart - charsToTypeStart;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Find.Execute("SETBACK");
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Start = lineStartInt;
                    //range.Move(Word.WdUnits.wdLine, 0);
                    int charsToSetbackStart = range.Text.Length;
                    int charsInDepth = charsToSetbackStart - charsToDepthStart;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                    range.Find.Execute(HCSeatInfo[i].bplSide);
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    range.Move(Word.WdUnits.wdCharacter, charsInDesc);
                    range.MoveEnd(Word.WdUnits.wdCharacter, charsInSection);

                    string newSectionType = "";
                    if (HCSeatInfo[i].buttedSeat==true) { newSectionType = "HC-1"; }
                    if (HCSeatInfo[i].gappedSeat==true) { newSectionType = "HC-2"; }
                    if (HCSeatInfo[i].plateSeatP04 == true) { newSectionType = "HC-3"; }
                    if (HCSeatInfo[i].plateSeatP08 == true) { newSectionType = "HC-4"; }

                    while (newSectionType.Length < charsInSection)
                    {
                        newSectionType = newSectionType + " ";
                    }

                    range.Text = newSectionType;
                    range.Font.Bold = 1;
                    range.Underline = Word.WdUnderline.wdUnderlineSingle;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                    if (HCSeatInfo[i].gappedSeat == true)
                    {

                        string newSlopeString = "";
                        if (HCSeatInfo[i].bplOutsideDepth != HCSeatInfo[i].bplInsideDepth)
                        {

                            if (HCSeatInfo[i].bplSide == "BPL-L")
                            {
                                newSlopeString = HCSeatInfo[i].bplOutsideDepth + " | " + HCSeatInfo[i].bplInsideDepth;
                            }
                            if (HCSeatInfo[i].bplSide == "BPL-R")
                            {
                                newSlopeString = HCSeatInfo[i].bplInsideDepth + " | " + HCSeatInfo[i].bplOutsideDepth;
                            }
                        }

                        if (HCSeatInfo[i].bplOutsideDepth != HCSeatInfo[i].bplInsideDepth)
                        {

                            range.Move(Word.WdUnits.wdCharacter, charsInSection);
                            range.MoveEnd(Word.WdUnits.wdCharacter, charsInLength);

                            string lengthString = HCSeatInfo[i].bplLength.ToString();

                            while (lengthString.Length < charsInLength)
                            {
                                lengthString = lengthString + " ";
                            }

                            range.Text = lengthString;
                            range.Font.Bold = 1;
                            range.Underline = Word.WdUnderline.wdUnderlineSingle;
                            range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                            range.Move(Word.WdUnits.wdCharacter, charsInLength + charsInType);
                            range.MoveEnd(Word.WdUnits.wdCharacter, charsInDepth);
                            
                            while(newSlopeString.Length<charsInDepth)
                            {
                                newSlopeString = newSlopeString + " ";
                            }

                            range.Text = newSlopeString;
                            range.Font.Bold = 1;
                            range.Underline = Word.WdUnderline.wdUnderlineSingle;
                            range.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                        }
                        else
                        {
                            range.Move(Word.WdUnits.wdCharacter, charsInSection);
                            range.MoveEnd(Word.WdUnits.wdCharacter, charsInLength);

                            string lengthString = HCSeatInfo[i].bplLength.ToString();

                            while (lengthString.Length < charsInLength)
                            {
                                lengthString = lengthString + " ";
                            }

                            range.Text = lengthString;
                            range.Font.Bold = 1;
                            range.Underline = Word.WdUnderline.wdUnderlineSingle;
                            range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        }
                    }


                }

            }
        }


        public void createHoldClearSKs(List<List<HCSeatInfo>> HCSeatInfoList)
        {
            /*            List<List<string>> joistData = JoistCoverSheet.JoistData();
                        List<List<HCSeatInfo>> HCSeatInfoList = organizedHCSeatInfoList();
             */

            for (int k = 0; k <= 3; k++)
            {
                List<HCSeatInfo> HCSeatInfo = HCSeatInfoList[k];


                if (HCSeatInfo.Count() != 0)
                {


                    Bitmap SK1 = Properties.Resources.buttedHC;

                    if (HCSeatInfoList[k][0].plateSeatP04 == true)
                    {
                        SK1 = Properties.Resources.plateHCP04;
                    }
                    if (HCSeatInfoList[k][0].plateSeatP08 == true)
                    {
                        SK1 = Properties.Resources.plateHCP08;
                    }
                    if (HCSeatInfoList[k][0].gappedSeat == true)
                    {
                        SK1 = Properties.Resources.gappedHC;
                    }


                    Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                    selection.EndKey(Word.WdUnits.wdStory, 1);
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    Word.Section section = selection.Sections.Add();
                    selection.EndKey(Word.WdUnits.wdStory, 1);
                    selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                    selection.PageSetup.LeftMargin = (float)50;
                    selection.PageSetup.RightMargin = (float)50;

                    
                    section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;

                    Word.Range currentLocation = Globals.ThisAddIn.Application.ActiveDocument.Range(0,0);
                    currentLocation = selection.Range;
                    Word.Range headerRange = Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0);
                    headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                    var window = Globals.ThisAddIn.Application.ActiveWindow;
                    // wdNormalView == Draft View, where SeekView can't be used and isn't needed.
                    if (window.View.Type != Word.WdViewType.wdNormalView)
                    {
                        // -1 Not Header/Footer, 0 Even page header, 1 Odd page header, 4 First page header
                        // 2 Even page footer, 3 Odd page footer, 5 First page footer
                        int rangeType = headerRange.Information[Word.WdInformation.wdHeaderFooterType];
                        if (rangeType == 0 || rangeType == 1 || rangeType == 4)
                            window.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;
                        if (rangeType == 2 || rangeType == 3 || rangeType == 5)
                            window.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
                    }
                    headerRange.Select();
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    selection.Text = "";
                    
                    
                    currentLocation.Select();



                    selection.Font.Name = "Times New Roman";
                    selection.Font.Size = 9;

                    Word.Table tableHoldClearCover = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 2, 6);

                    tableHoldClearCover.Cell(1, 1).Range.Text = "JOB NAME: ";
                    tableHoldClearCover.Cell(2, 1).Range.Text = "LOCATION: ";
                    tableHoldClearCover.Cell(1, 3).Range.Text = "JOB #: ";
                    tableHoldClearCover.Cell(2, 3).Range.Text = "LIST:  ";
                    tableHoldClearCover.Cell(1, 5).Range.Text = "SHEET #:  ";
                    if (k == 0) { tableHoldClearCover.Cell(1, 6).Range.Text = "HC-1"; }
                    if (k == 1) { tableHoldClearCover.Cell(1, 6).Range.Text = "HC-2"; }
                    if (k == 2) { tableHoldClearCover.Cell(1, 6).Range.Text = "HC-3"; }
                    if (k == 3) { tableHoldClearCover.Cell(1, 6).Range.Text = "HC-4"; }

                    for (int i = 1; i <= 2; i++)
                    {
                        for (int col = 1; col <= 5; col = col + 2)
                        {
                            tableHoldClearCover.Cell(i, col).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            tableHoldClearCover.Cell(i, col).Range.Underline = Word.WdUnderline.wdUnderlineSingle;
                            tableHoldClearCover.Cell(i, col).Range.Font.Bold = 1;
                        }

                    }


                    tableHoldClearCover.Cell(1, 2).Range.Text = joistData[8][0];
                    tableHoldClearCover.Cell(2, 2).Range.Text = joistData[9][0];
                    tableHoldClearCover.Cell(1, 4).Range.Text = joistData[7][0];
                    tableHoldClearCover.Cell(2, 4).Range.Text = joistData[10][0];

                    for (int i = 1; i <= 2; i++)
                    {
                        for (int col = 2; col <= 4; col = col + 2)
                        {
                            tableHoldClearCover.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        }
                    }




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

                    IDataObject idat = null;
                    Exception threadEx = null;
                    Thread staThread = new Thread(
                        delegate()
                        {
                            try
                            {

                                string sk1FileName = System.IO.Path.GetTempFileName();
                                Byte[] sk1ByteArray = ImageToByte(SK1);
                                System.IO.File.WriteAllBytes(sk1FileName, sk1ByteArray);

                                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("bookmark", selection.Range);
                                var shape = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["bookmark"].Range.InlineShapes.AddPicture(sk1FileName);

                                float dpiX, dpiY;
                                Graphics graphics = this.CreateGraphics();
                                dpiX = graphics.DpiX;
                                dpiY = graphics.DpiY;
                                float initialWidth = shape.Width;
                                float initialHeight = shape.Height;
                                float ratio = initialWidth / initialHeight;

                                shape.Width = Convert.ToInt16(72.0 * 9.5);
                                shape.Height = Convert.ToInt16(72.0 * 9.5) / ratio;

                            }

                            catch (Exception ex)
                            {
                                threadEx = ex;
                            }
                        });
                    staThread.SetApartmentState(ApartmentState.STA);
                    staThread.Start();
                    staThread.Join();




                    selection.EndKey(Word.WdUnits.wdStory, 1);
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    selection.Text = "\r\n";
                    selection.EndKey(Word.WdUnits.wdStory, 1);
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    Word.Table tblHC = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 3, 12);


                    tblHC.Columns[1].Width = 45;
                    tblHC.Columns[2].Width = 160;
                    tblHC.Columns[3].Width = 45;
                    tblHC.Columns[4].Width = 45;
                    tblHC.Columns[5].Width = 45;
                    tblHC.Columns[6].Width = 45;
                    tblHC.Columns[7].Width = 45;
                    tblHC.Columns[8].Width = 45;
                    tblHC.Columns[9].Width = 60;
                    tblHC.Columns[10].Width = 45;
                    tblHC.Columns[11].Width = 45;
                    tblHC.Columns[12].Width = 65;




                    tblHC.Borders.Enable = 1;


                    tblHC.Cell(1, 7).Merge(MergeTo: tblHC.Cell(1, 9));

                    for (int col = 1; col <= tblHC.Columns.Count - 2; col++)
                    {
                        if (col != 7)
                        {
                            tblHC.Cell(1, col).Borders.Enable = 0;
                        }
                    }
                    tblHC.Cell(1, 7).Borders.Enable = 1;


                    for (int col = 1; col <= tblHC.Columns.Count; col++)
                    {
                        tblHC.Cell(2, col).Borders.Enable = 1;
                    }

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
                    tblHC.Cell(2, 12).Range.Text = "Stiff. Length";

                    for (int col = 1; col <= 11; col++)
                    {
                        tblHC.Cell(2, col).Range.Font.Size = 8;
                        tblHC.Cell(2, col).Range.Font.Bold = 1;
                        tblHC.Cell(2, col).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    }

                    int row = 3;

                    foreach (HCSeatInfo hcSeatInfo in HCSeatInfo)
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
                        tblHC.Cell(row, 12).Range.Text = hcSeatInfo.stiffPlateLength;


                        tblHC.Rows.Add();
                        row = row + 1;

                    }


                    for (int row1 = 2; row1 <= tblHC.Rows.Count; row1++)
                    {
                        for (int col1 = 1; col1 <= tblHC.Columns.Count; col1++)
                        {
                            if (col1 != 2)
                            {
                                tblHC.Cell(row1, col1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            }
                        }
                    }
                    if (k == 1)
                    {
                        tblHC.Cell(2, 5).Select();
                        selection.Columns.Delete();
                        tblHC.Cell(2, 5).Select();
                        selection.Columns.Delete();
                    }
                    if (k == 2)
                    {
   //                     for (int row2 = 3; row2 <= tblHC.Rows.Count - 1; row2++)
   //                     {
    //                        tblHC.Cell(row2, 3).Range.Text = "P0604";
    //                    }
                        tblHC.Cell(2, 6).Select();
                        selection.InsertColumnsRight();
                        tblHC.Cell(2, 7).Range.Text = "C";
                        tblHC.Cell(1, 7).Width = 45;
                        tblHC.Cell(1, 7).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        tblHC.Cell(1, 7).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;

                        for (int row2 = 1; row2 <= tblHC.Rows.Count; row2++)
                        {
                            tblHC.Cell(row2, 1).Width = 25;
                            tblHC.Cell(row2, 2).Width = 135;
                        }
                        int row4 = 3;

                        foreach (HCSeatInfo hcSeatInfo in HCSeatInfo)
                        {
                            tblHC.Cell(row4, 7).Range.Text = hcSeatInfo.paWidth;
                            tblHC.Cell(row4, 3).Range.Text = hcSeatInfo.paMat;
                            row4++;
                        }
                    }

                    if (k == 3)
                    {
                        //                     for (int row2 = 3; row2 <= tblHC.Rows.Count - 1; row2++)
                        //                     {
                        //                        tblHC.Cell(row2, 3).Range.Text = "P0604";
                        //                    }
                        tblHC.Cell(2, 6).Select();
                        selection.InsertColumnsRight();
                        tblHC.Cell(2, 7).Range.Text = "C";
                        tblHC.Cell(1, 7).Width = 45;
                        tblHC.Cell(1, 7).Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                        tblHC.Cell(1, 7).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;

                        for (int row2 = 1; row2 <= tblHC.Rows.Count; row2++)
                        {
                            tblHC.Cell(row2, 1).Width = 25;
                            tblHC.Cell(row2, 2).Width = 135;
                        }
                        int row4 = 3;

                        foreach (HCSeatInfo hcSeatInfo in HCSeatInfo)
                        {
                            tblHC.Cell(row4, 7).Range.Text = hcSeatInfo.paWidth;
                            tblHC.Cell(row4, 3).Range.Text = hcSeatInfo.paMat;
                            row4++;
                        }
                    }

                    tblHC.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    selection.EndKey(Word.WdUnits.wdStory, 1);
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
            }

            this.Close();

        }




        private void btnCreateHoldClears_Click(object sender, EventArgs e)
        {
            string clipboard = Clipboard.GetText();

            List<List<string>> joistData = JoistCoverSheet.JoistData();
            List<List<string>> joistDataByMarks = joistDataByMark();
            List<Tuple<string, bool, bool>> marksWithHoldClears = whichNeedHoldClears(joistDataByMarks);
            AllSeatInfo allseatInfo = getAllSeatInfo(joistDataByMarks, marksWithHoldClears);
            List<List<HCSeatInfo>> HCSeatInfoList1 = listHCSeatInfo2(allseatInfo.HoldClear, true);
            List<List<HCSeatInfo>> HCSeatInfoList2 = listHCSeatInfo2(allseatInfo.HoldClear, false);
            createHoldClearSKs(organizedHCSeatInfoList(HCSeatInfoList1));
            placeHCs(joistDataByMarks, HCSeatInfoList2);
            if (noOverides == false)
            {
                MessageBox.Show("PLEASE CONFIRM THAT ALL OVERRIDDEN DETAILS ARE ACCURATE");
            }

            try { Clipboard.SetText(clipboard); }
            catch { }

        }




        public Image textOnImage(Image image, List<Tuple<string, int, int, int>> listOfText)
        {
            Bitmap bitMapImage = new
                System.Drawing.Bitmap(image);
            Graphics graphicImage = Graphics.FromImage(bitMapImage);

            graphicImage.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            foreach (Tuple<string, int, int, int> thisText in listOfText)
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

        public static byte[] ImageToByte(Image img)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }



    }



    public class JoistSeatInfo
    {
        private StringManipulation stringManipulation = new StringManipulation();
        public string mark;
        public int qty;
        public string TC;
        public string bplSide;
        public string bplLength;
        public double bplOutsideDepth;
        public double bplInsideDepth;
        public string slotSetback;
        public string slotSize;
        public double slotDiameter
        {
            get
            {
                if (slotSize != null && slotSize.Trim() != "")
                {
                    var slotSizeArray = slotSize.Split(new char[] { 'x' });
                    var slotDiameterString = slotSizeArray[0];
                    var slotDiameter = stringManipulation.hyphenLengthToDecimal("0-0 " + slotDiameterString) * 12.0;
                    return slotDiameter;

                }
                else
                {
                    return 0.0;
                }
            }
        }

        public double slotLength
        {
            get
            {
                if (slotSize != null && slotSize.Trim() != "")
                {
                    var slotSizeArray = slotSize.Split(new char[] { 'x' });
                    var slotLengthString = slotSizeArray[1];
                    var slotLength = stringManipulation.hyphenLengthToDecimal("0-" + slotLengthString) * 12.0;
                    return slotLength;

                }
                else
                {
                    return 0.0;
                }
            }
        }
        public double slotGauge;
        public double clearLeft;
        public double clearRight;
        public string seatType;
        public double tcxL;
        public double tcxR;
        public double ota;
        public double oal;
    }
    public class HCSeatInfo
    {
        private StringManipulation stringManipulation = new StringManipulation();
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
        public double slotDiameter
        {
            get
            {
                if (slotSize != null && slotSize.Trim() != "")
                {
                    var slotSizeArray = slotSize.Split(new char[] { 'x' });
                    var slotDiameterString = slotSizeArray[0];
                    var slotDiameter = stringManipulation.hyphenLengthToDecimal("0-0 " + slotDiameterString) * 12.0;
                    return slotDiameter;

                }
                else
                {
                    return 0.0;
                }
            } 
        }

        public double slotLength
        {
            get
            {
                if (slotSize != null && slotSize.Trim() != "")
                {
                    var slotSizeArray = slotSize.Split(new char[] { 'x' });
                    var slotLengthString = slotSizeArray[1];
                    var slotLength = stringManipulation.hyphenLengthToDecimal("0-" + slotLengthString) * 12.0;
                    return slotLength;

                }
                else
                {
                    return 0.0;
                }
            }
        }
        public string GOL;
        public bool gappedSeat = false;
        public bool plateSeatP04 = false;
        public bool plateSeatP08 = false;
        public bool buttedSeat = false;
        public string stiffPlateLength;
        public string paWidth;
        public string paMat;
        public string seatType;
    }

    public class AllSeatInfo
    {
        public List<JoistSeatInfo> HoldClear;
        public List<JoistSeatInfo> Tplate;
        public List<JoistSeatInfo> BCG;
        public List<JoistSeatInfo> Standard;

    }




}
