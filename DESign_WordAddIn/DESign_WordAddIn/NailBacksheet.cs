﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using DESign_WordAddIn;
using System.Globalization;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using DESign_BASE;
using DESign_WordAddIn;
using System.Xml.Serialization;
using Newtonsoft.Json;



namespace DESign_WordAddIn
{
    public partial class FormNailBacksheet : Form
    {
        

        StringManipulation StringManipulation = new StringManipulation();

        ExcelDataExtraction excelDataExtraction = new ExcelDataExtraction();

        DESign_BASE.QueryAngleData QueryAngleData = new DESign_BASE.QueryAngleData();
        List<DESign_BASE.Angle> anglesFromSql = QueryAngleData.AnglesFromSql("Fallon");


        public FormNailBacksheet()
        {
            InitializeComponent();
            tBoxScrewSpacing.Text = "24";
            cbWoodThickness.Text = "3x";
        }

        List<TextBox> tBoxAList = new List<TextBox>();
        List<TextBox> tBoxBList = new List<TextBox>();
        List<string> stringListLengthA = new List<string>();
        List<string> stringListLengthB = new List<string>();
        TextBox tboxAllAs = new TextBox();
        TextBox tboxAllBs = new TextBox();

        ComboBox comboBoxNailPlacement = new ComboBox();

        CoverSheetInfo joistCoverSheetInfo = CoverSheetInfo.FromV1So();


        private void FormNailBacksheet_Load(object sender, EventArgs e)
        {

            string clipboard = Clipboard.GetText();

            

            var labelMarkTitle = new Label();
            var labelATitle = new Label();
            var labelBTitle = new Label();


            labelMarkTitle.Size = new System.Drawing.Size(60, 30);
            labelATitle.Size = new System.Drawing.Size(50, 30);
            labelBTitle.Size = new System.Drawing.Size(50, 30);

            labelMarkTitle.Location = new Point(235, 0);
            labelATitle.Location = new Point(315, 0);
            labelBTitle.Location = new Point(395, 0);

            labelMarkTitle.Text = "MARK";
            labelATitle.Text = "A";
            labelBTitle.Text = "B";

            labelMarkTitle.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelATitle.Font = new Font("Times New Roman", 9, FontStyle.Bold);
            labelBTitle.Font = new Font("Times New Roman", 9, FontStyle.Bold);

            labelMarkTitle.TextAlign = ContentAlignment.MiddleLeft;
            labelATitle.TextAlign = ContentAlignment.MiddleCenter;
            labelBTitle.TextAlign = ContentAlignment.MiddleCenter;

            var labelAllMarks = new Label();

            labelAllMarks.Size = labelMarkTitle.Size;
            tboxAllAs.Size = labelATitle.Size;
            tboxAllBs.Size = labelBTitle.Size;

            labelAllMarks.Location = new Point(235, 30);
            tboxAllAs.Location = new Point(315, 30);
            tboxAllBs.Location = new Point(395, 30);

            labelAllMarks.Text = "ALL";

            labelAllMarks.TextAlign = ContentAlignment.MiddleLeft;
            tboxAllAs.Text = "";
            tboxAllBs.Text = "";

            this.Controls.Add(labelMarkTitle);
            this.Controls.Add(labelATitle);
            this.Controls.Add(labelBTitle);

            this.Controls.Add(labelAllMarks);
            this.Controls.Add(tboxAllAs);
            this.Controls.Add(tboxAllBs);


            List<string> joistMarks = joistCoverSheetInfo.SoSummary.Select(s => s.Mark).ToList();

            int joistDataLength = joistMarks.Count();



            var labelMark = new Label[joistDataLength];
            var tboxA = new TextBox[joistDataLength];
            var tboxB = new TextBox[joistDataLength];


            for (var i = 0; i < joistDataLength; i++)
            {
                var labelMarks = new Label();
                var tboxAs = new TextBox();
                var tboxBs = new TextBox();

                int Y = 70 + (i * 25);

                labelMarks.Text = joistMarks[i];
                labelMarks.Location = new Point(235, Y);

                labelMarks.Size = new System.Drawing.Size(50, 25);

                tboxAs.Location = new Point(315, Y);
                tboxAs.Size = new System.Drawing.Size(50, 10);

                tboxBs.Location = new Point(395, Y);
                tboxBs.Size = new System.Drawing.Size(50, 10);

                this.Controls.Add(tboxAs);
                this.Controls.Add(labelMarks);
                this.Controls.Add(tboxBs);

                tBoxAList.Add(tboxAs);
                tBoxBList.Add(tboxBs);


                tboxA[i] = tboxAs;
                labelMark[i] = labelMarks;
                tboxB[i] = tboxBs;
            }

            tboxTolerance.Text = "";

            comboBoxNailPlacement.DrawMode = System.Windows.Forms.DrawMode.Normal;

            string[] placementTypes = new string[] { "Staggered", "Non-Staggered" };

            comboBoxNailPlacement.DataSource = placementTypes;

            comboBoxNailPlacement.Location = new Point(10, 92);
            comboBoxNailPlacement.Size = new System.Drawing.Size(164, 24);
            comboBoxNailPlacement.TabIndex = 0;
            comboBoxNailPlacement.DropDownWidth = scaleWidth(164);

            comboBoxNailPlacement.DropDownStyle = ComboBoxStyle.DropDownList;

            comboBoxNailPlacement.Enabled = true;

            try { Clipboard.SetText(clipboard); }
            catch { }

            this.Controls.Add(comboBoxNailPlacement);


        }


        internal ExcelDataExtraction.NailerInformation GetNailerInformation()
        {
            ExcelDataExtraction.NailerInformation nailerInformation = new ExcelDataExtraction.NailerInformation();
            string clipboard = Clipboard.GetText();

            List<string> shopOrderjoistMarks = joistCoverSheetInfo.SoSummary.Select(s => s.Mark).ToList();
            int shopOrdernumberOfMarks = joistCoverSheetInfo.SoSummary.Count();

            List<string> excelJoistMarks = null;
            List<string> excelAs = null;
            List<string> excelBs = null;
            List<string> excelSpacing = null;
            string excelInitials = null;
            string excelPattern = null;
            string excelWoodThickness = null;
            bool isPanelized = cbIsPanelized.Checked;
            bool crimpedWebs = cbCrimpedWebs.Checked;

            if (checkBoxExcelData.Checked)
            {
                try
                {

                    string excelFilePath = Globals.ThisAddIn.Application.ActiveDocument.Path.ToString() + "\\Note Info.xlsx";


                    ExcelDataExtraction.NailerInformation excelJoistData = excelDataExtraction.exlNailerValues(excelFilePath);
                    excelJoistMarks = excelJoistData.Marks;
                    excelAs = excelJoistData.As;
                    excelBs = excelJoistData.Bs;
                    excelSpacing = excelJoistData.Spacing;
                    excelInitials = excelJoistData.Initials;
                    excelPattern = excelJoistData.Pattern;
                    excelWoodThickness = excelJoistData.WoodThickness;
                    isPanelized = excelJoistData.IsPanelized;
                    crimpedWebs = excelJoistData.CrimpedWebs;
                }
                catch
                {
                    MessageBox.Show("ISSUE WITH EXCEL SHEET");
                }
            }





            // NEW METHOD FOR LISTLENGTHS
            List<double> doubleListLengthA = new List<double>();

            for (int i = 0; i < shopOrdernumberOfMarks; i++)
            {
                string stringLengthA = null;
                string sOjoistMark = shopOrderjoistMarks[i];
                if (checkBoxExcelData.Checked)
                {
                    try
                    {
                        int SOJoistMarkIndex = Array.FindIndex(excelJoistMarks.ToArray(), t => t.Equals(sOjoistMark, StringComparison.InvariantCultureIgnoreCase));
                        stringLengthA = excelAs[SOJoistMarkIndex];
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("ISSUE WITH EXCEL SHEET AT MARK " + sOjoistMark + ";" + Environment.NewLine + "PLEASE UPDATE OR MANUALLY INPUT A & B");
                        

                    }

                    //     int SOJoistMarkIndex = excelJoistMarks.IndexOf(shopOrderjoistMarks[i]);



                }
                else
                {
                    if (tboxAllAs.Text != "")
                    {
                        stringLengthA = StringManipulation.convertLengthStringtoHyphenLength(tboxAllAs.Text);
                    }
                    else
                    {
                        stringLengthA = StringManipulation.convertLengthStringtoHyphenLength(tBoxAList[i].Text);
                    }
                }
                double doubleLengthA = StringManipulation.hyphenLengthToDecimal(stringLengthA);

                stringLengthA = StringManipulation.DecimilLengthToHyphen(doubleLengthA);

                doubleListLengthA.Add(doubleLengthA);
                stringListLengthA.Add(stringLengthA);
            }

            List<string> spacingList = new List<string>();
            for (int i = 0; i < shopOrdernumberOfMarks; i ++)
            {
                string sOjoistMark = shopOrderjoistMarks[i];
                string thisSpace = null;
                if (checkBoxExcelData.Checked)
                {
                    try
                    {
                        int SOJoistMarkIndex = Array.FindIndex(excelJoistMarks.ToArray(), t => t.Equals(sOjoistMark, StringComparison.InvariantCultureIgnoreCase));
                        thisSpace = excelSpacing[SOJoistMarkIndex];
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("ISSUE WITH EXCEL SHEET AT MARK: " + sOjoistMark + ";" + Environment.NewLine + "PLEASE UPDATE");

                    }

                }
                else
                {
                    thisSpace = tBoxScrewSpacing.Text;
                }
                spacingList.Add(thisSpace);
            }



            List<double> doubleListLengthB = new List<double>();

            for (int i = 0; i < shopOrdernumberOfMarks; i++)
            {
                string stringLengthB = null;
                string sOjoistMark = shopOrderjoistMarks[i];


                if (checkBoxExcelData.Checked)
                {
                    try
                    {
                        int SOJoistMarkIndex = Array.FindIndex(excelJoistMarks.ToArray(), t => t.Equals(sOjoistMark, StringComparison.InvariantCultureIgnoreCase));
                        stringLengthB = excelBs[SOJoistMarkIndex];
                    }
                    catch(Exception e)
                    {
                        MessageBox.Show("ISSUE WITH EXCEL SHEET AT MARK: " + sOjoistMark + ";" + Environment.NewLine + "PLEASE UPDATE");
                        
                    }        

                }
                else
                {

                    if (tboxAllBs.Text != "")
                    {
                        stringLengthB = StringManipulation.convertLengthStringtoHyphenLength(tboxAllBs.Text);
                    }
                    else
                    {
                        stringLengthB = StringManipulation.convertLengthStringtoHyphenLength(tBoxBList[i].Text);
                    }

                }
                double doubleLengthB = StringManipulation.hyphenLengthToDecimal(stringLengthB);
                stringLengthB = StringManipulation.DecimilLengthToHyphen(doubleLengthB);

                doubleListLengthB.Add(doubleLengthB);
                stringListLengthB.Add(stringLengthB);

            }

            //END NEW METHOD FOR LIST LENGTHS
            List<double> listLengthJoist = joistCoverSheetInfo.SoSummary.Select(s => s.Length).ToList();

            List<double> woodLength = new List<double>();
            double woodLengthi = 0;
            for (int i = 0; i < shopOrdernumberOfMarks; i++)
            {
                woodLengthi = listLengthJoist[i] - doubleListLengthA[i] - doubleListLengthB[i];

                woodLength.Add(woodLengthi);
            }

            // SORTING OF woodLength:


            bool hasHyphen = tboxTolerance.Text.Contains("-");
            bool hasBackslash = tboxTolerance.Text.Contains("/");
            bool hasSpace = tboxTolerance.Text.Contains(" ");

            double tolerance = 0;
            if (hasHyphen == true) { tolerance = StringManipulation.hyphenLengthToDecimal(tboxTolerance.Text); }
            else if (hasHyphen == false && hasSpace == true && hasBackslash == true) { tolerance = StringManipulation.hyphenLengthToDecimal("0-" + tboxTolerance.Text); }
            else if (hasHyphen == false && hasSpace == false && hasBackslash == true) { tolerance = StringManipulation.hyphenLengthToDecimal("0-0 " + tboxTolerance.Text); }
            else { }

            woodLength = StringManipulation.doubleListwithTolerance(woodLength, tolerance);


            //

            List<string> listStringWoodLength = new List<string>();
            string stringWoodLengthi = "a";
            for (int i = 0; i < shopOrdernumberOfMarks; i++)
            {
                stringWoodLengthi = StringManipulation.DecimilLengthToHyphen(woodLength[i]);
                listStringWoodLength.Add(stringWoodLengthi);
            }


            try { Clipboard.SetText(clipboard); }
            catch { }


            string pattern = "Staggered";
            string initials = "";
            string spacing = "";
            string woodThickness = "3x";
            if (checkBoxExcelData.Checked)
            {
                pattern = excelPattern;
                initials = excelInitials;
                woodThickness = excelWoodThickness;

            }
            else
            {
                pattern = comboBoxNailPlacement.Text;
                initials = tBoxDWGBY.Text;
                spacing = tBoxScrewSpacing.Text;
                woodThickness = cbWoodThickness.Text;
            }

            nailerInformation.WoodLengths = listStringWoodLength;
            nailerInformation.Spacing = spacingList;
            nailerInformation.Pattern = pattern;
            nailerInformation.Initials = initials;
            nailerInformation.WoodThickness = woodThickness;
            nailerInformation.IsPanelized = isPanelized;
            nailerInformation.CrimpedWebs = crimpedWebs;

            return nailerInformation;


        }
        private void btnCreateTable_Click(object sender, EventArgs e)
        {
            string clipboard = Clipboard.GetText();

            int numberOfMarks = joistCoverSheetInfo.SoSummary.Count();

            ExcelDataExtraction.NailerInformation nailerInfo = GetNailerInformation();
            List<string> woodLengths1 = nailerInfo.WoodLengths;
            List<string> woodWidths = new List<string>(numberOfMarks);
            List<double> horizontalLegs = new List<double>();

            
            for (int i = 0; i < numberOfMarks; i++)
            {
                var soSummaryLine = joistCoverSheetInfo.SoSummary[i];
             
                if (nailerInfo.CrimpedWebs)
                {
                    woodWidths.Add(QueryAngleData.WNtcWidth(anglesFromSql, soSummaryLine.Tc) + " 1/8" );
                }
                else
                {
                    woodWidths.Add(QueryAngleData.WNtcWidth(anglesFromSql, soSummaryLine.Tc) + " 1/4" );
                }
            }

            var TCs = joistCoverSheetInfo.SoSummary.Select(s => s.Tc).ToList();

            if ((TCs.Contains("A50A28") || TCs.Contains("A48A28") || TCs.Contains("A48A29")) &&
                (TCs.Contains("A42A28") || TCs.Contains("A44A") || TCs.Contains("A46A28")))
            {
                MessageBox.Show("Shop Order must be split due to TC sizes");
                throw new System.Exception();
            }

            if (
                TCs
                .Select (tc => QueryAngleData.WNtcWidth(anglesFromSql, tc))
                .Distinct()
                .Count() > 1)
            {
                MessageBox.Show("Shop Order must be split due to TC sizes");
                throw new System.Exception();
            }
                        


            Word.Selection selection = Globals.ThisAddIn.Application.Selection;

            selection.HomeKey(Word.WdUnits.wdStory);

            var requres1InchGap =
                TCs
                .Select (tc => QueryAngleData.Requres1InchGap(anglesFromSql, tc))
                .First();




            var coverSheetNotes = new List<string>();

            if (nailerInfo.CrimpedWebs == false)
            {
                requres1InchGap = false;
                var tcWidth = woodWidths[0].Replace(" 1/4", "");
                coverSheetNotes.Add("SPECIAL WOOD NAILER; " + tcWidth + " 1/8\"" + " STEEL O-T-O, " + tcWidth + " 1/4\" WOOD WIDTH – SEE N1");
            }
            else
            {
                var tcWidth = woodWidths[0].Replace(" 1/8", "");
                coverSheetNotes.Add("WOOD NAILER; " + tcWidth + "\"" + " STEEL WIDTH, " + tcWidth + " 1/8\" WOOD WIDTH – SEE N1");
            }

            if (requres1InchGap)
            {
                coverSheetNotes.Add("PROVIDE 1\" GAP BETWEEN TOP CHORD ANGLES, END CRIMP WEBS");
            }


            if (nailerInfo.WoodThickness == "2\"")
            {
                coverSheetNotes.Add("SPECIAL WOODNAILER: USE 2\" WOOD - SEE N1");
            }

            Predicate<string> isComposite = delegate (string tc) { return tc.Contains("W"); };
            if (TCs.Find(isComposite) != null)
            {
                coverSheetNotes.Add("COMPOSITE WOODNAILER; THIS JOIST MUST BE BUILT IN BLDG 2");
                coverSheetNotes.Add("WOOD MUST BE CONTINUOUS, FINGER JOINTS ONLY, NO BUTTED JOINTS");
            }

            if (nailerInfo.IsPanelized)
            {
                coverSheetNotes.Add("PANELIZED INC.; NO MID-SPAN JOINTS");
            }

            var coverSheetNote = String.Join("\r\n", coverSheetNotes);

            addTextBox(selection, 130, 700, 350, 20, coverSheetNote);


            if (requres1InchGap)
            {

                var docWords = Globals.ThisAddIn.Application.ActiveDocument.Words;

                foreach (Word.Range word in docWords)
                {
                    if (word.Text.Contains("R112"))
                    {
                        MessageBox.Show("Shop-Order Contains \"R112\";\r\n Please confirm that this is OK since the gap will be 1\".");
                    }
                }

               
                

                selection.Find.Execute("WEB CUT SHEET");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("WebCutStart", selection.Range);
                selection.Find.Execute("\f");
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("WebCutEnd", selection.Range);

                int webCutStart = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["WebCutStart"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
                int webCutEnd = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["WebCutEnd"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];

                for (int page = webCutStart; page <= webCutEnd; page++)
                {
                    selection.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, page);
                    addTextBox(selection, 460, 135, 120, 70, "PROVIDE 1\" GAP BETWEEN TOP CHORD ANGLES, END CRIMP WEBS");
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }
                
            }

            

            foreach (string mark in joistCoverSheetInfo.SoSummary.Select(s => s.Mark))
            {
                Word.Range range = Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0);

                range.Find.Execute("Color Code");
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                range.Find.Execute(mark + "    ");
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range.Find.Execute("W2L");
                range.MoveStart(Word.WdUnits.wdWord, -1);

                string[] w2LText = range.Text.Split(new string[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
                string w2LMaterial = w2LText[0].Trim();
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                range.Find.Execute("W2R");
                range.MoveStart(Word.WdUnits.wdWord, -1);

                string[] w2Rtext = range.Text.Split(new string[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
                string w2RMaterial = w2Rtext[0].Trim();

                string[] unacceptableW2s = new string[] { "C36BA", "C38BA", "C40BA", "CW40BA" };

                bool badW2L = unacceptableW2s.Contains(w2LMaterial);
                bool badW2R = unacceptableW2s.Contains(w2RMaterial);

                string message = mark + ":\r\n";
                if (badW2L) { message = message + "Check W2L Material\r\n"; }
                if (badW2R) { message = message + "Check W2R Material\r\n"; }


                if (badW2L || badW2R)
                {
                    MessageBox.Show(message);
                }

            }


            selection.Find.Execute("CHORD CUT SHEET");

            if (nailerInfo.WoodThickness == "2\"")
            {
                addTextBox(selection, 420, 135, 120, 50, "WOODNAILER; SEE N1\r\n*** 2\" WOOD ***");
            }
            else
            {
                addTextBox(selection, 420, 135, 120, 20, "WOODNAILER; SEE N1");
            }




            selection.EndKey(Word.WdUnits.wdStory, 1);

            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            Word.Section section = selection.Sections.Add();

            selection.EndKey(Word.WdUnits.wdStory, 1);

            selection.EndKey(Word.WdUnits.wdStory, 1);


            //ADDING JOBINFORMATION & LIST NUMBER


            Word.Table tableNailBackSheetTitle = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 4, 5);

            tableNailBackSheetTitle.Cell(1, 1).Range.Text = "JOB NAME: ";
            tableNailBackSheetTitle.Cell(2, 1).Range.Text = "LOCATION: ";
            tableNailBackSheetTitle.Cell(3, 1).Range.Text = "JOB #:     ";
            tableNailBackSheetTitle.Cell(4, 1).Range.Text = "LIST:       ";

            for (int i = 1; i <= 4; i++)
            {
                tableNailBackSheetTitle.Cell(i, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                tableNailBackSheetTitle.Cell(i, 1).Range.Underline = Word.WdUnderline.wdUnderlineSingle;
                tableNailBackSheetTitle.Cell(i, 1).Range.Font.Bold = 1;

            }

            tableNailBackSheetTitle.Cell(1, 2).Range.Text = joistCoverSheetInfo.JobLocation;
            tableNailBackSheetTitle.Cell(2, 2).Range.Text = joistCoverSheetInfo.JobName;
            tableNailBackSheetTitle.Cell(3, 2).Range.Text = joistCoverSheetInfo.JobNumber;
            tableNailBackSheetTitle.Cell(4, 2).Range.Text = joistCoverSheetInfo.ListNumber;

            for (int i = 1; i <= 4; i++)
            {
                tableNailBackSheetTitle.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }


            tableNailBackSheetTitle.Cell(1, 3).Range.Text = "DWG. BY:";
            tableNailBackSheetTitle.Cell(2, 3).Range.Text = nailerInfo.Initials;

            tableNailBackSheetTitle.Cell(1, 4).Range.Text = "CHK'D BY:";

            tableNailBackSheetTitle.Cell(1, 5).Range.Text = "SHEET #:";
            tableNailBackSheetTitle.Cell(2, 5).Range.Text = "N1 of 1";

            tableNailBackSheetTitle.Cell(4, 5).Range.Text = "TO SHOP";

            tableNailBackSheetTitle.Cell(1, 3).Borders.Enable = 1;
            tableNailBackSheetTitle.Cell(2, 3).Borders.Enable = 1;
            tableNailBackSheetTitle.Cell(1, 4).Borders.Enable = 1;
            tableNailBackSheetTitle.Cell(2, 4).Borders.Enable = 1;
            tableNailBackSheetTitle.Cell(1, 5).Borders.Enable = 1;
            tableNailBackSheetTitle.Cell(2, 5).Borders.Enable = 1;
            tableNailBackSheetTitle.Cell(4, 5).Borders.Enable = 1;

            tableNailBackSheetTitle.Cell(1, 3).Range.Font.Bold = 1;
            tableNailBackSheetTitle.Cell(1, 4).Range.Font.Bold = 1;
            tableNailBackSheetTitle.Cell(1, 5).Range.Font.Bold = 1;
            tableNailBackSheetTitle.Cell(4, 5).Range.Font.Bold = 1;




            for (int row = 1; row <= 4; row++)
                for (int column = 3; column <= 5; column++)
                {
                    tableNailBackSheetTitle.Cell(row, column).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            selection.EndKey(Word.WdUnits.wdStory, 1);

            for (int row = 1; row <= 4; row++)
                for (int column = 1; column <= 5; column++)
                {
                    tableNailBackSheetTitle.Cell(row, column).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }
            selection.EndKey(Word.WdUnits.wdStory, 1);

            tableNailBackSheetTitle.Columns[1].Width = 65;
            tableNailBackSheetTitle.Columns[2].Width = 250;
            tableNailBackSheetTitle.Columns[3].Width = 65;
            tableNailBackSheetTitle.Columns[4].Width = 65;
            tableNailBackSheetTitle.Columns[5].Width = 65;

            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.EndKey(Word.WdUnits.wdStory);
            selection.Text = "\r\n" + "\r\n";
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);


            // END ADDING JOB INFORMATION & LIST NUMBER


            Image NailPlacement = Properties.Resources.Staggered;


            if (nailerInfo.Pattern == "Staggered") { NailPlacement = Properties.Resources.Staggered; }
            else if (nailerInfo.Pattern == "Non-Staggered") { NailPlacement = Properties.Resources.NonStaggered; }
            else { }

            string nailSpacing = null;
            string screwSpacing = null;
            string woodThickness = null;
            if (checkBoxExcelData.Checked)
            {
                var spacings = nailerInfo.Spacing.Distinct();
                if (spacings.Count() > 1)
                {
                    MessageBox.Show("Shop Order must be split due to varying screw spacings");
                    throw new System.Exception();
                }
                else
                {
                    nailSpacing = spacings.First() + "\" MAX";
                    screwSpacing = spacings.First();
                }
                woodThickness = nailerInfo.WoodThickness == "3x" ? "2 1/2\"" : "2\"";
            }
            else
            {
                nailSpacing = tBoxScrewSpacing.Text + "\" MAX";
                screwSpacing = tBoxScrewSpacing.Text;                
                woodThickness = cbWoodThickness.Text == "3x" ? "2 1/2\"" : "2\"";
            }



            string hyphenScrewSpacing = null;
            double dblScrewSpacing = 0.0;
            double dblHalfScrewSpace = 0.0;
            string hyphenHalfScrewSpace = null;
            string halfScrewSpace_Inch = null;


            hyphenScrewSpacing = StringManipulation.convertLengthStringtoHyphenLength(screwSpacing);
            dblScrewSpacing = StringManipulation.ConvertLengthtoDecimal(hyphenScrewSpacing);
            dblHalfScrewSpace = dblScrewSpacing / 2.0;
            hyphenHalfScrewSpace = StringManipulation.decimalInchestoFraction(dblHalfScrewSpace);
            halfScrewSpace_Inch = hyphenHalfScrewSpace.Split('-')[1];


            List<Tuple<string, int, int, int>> listOfText = new List<Tuple<string, int, int, int>>();

            if (nailerInfo.Pattern == "Non-Staggered")
            {

                var text1 = new Tuple<string, int, int, int>(nailSpacing, 22, 252, 150);
                var text2 = new Tuple<string, int, int, int>(nailSpacing, 22, 1073, 150);
                var text3 = new Tuple<string, int, int, int>(nailSpacing, 22, 252, 325);
                var text4 = new Tuple<string, int, int, int>(nailSpacing, 22, 1073, 325);
                var text5 = new Tuple<string, int, int, int>(woodThickness, 22, 600, 565);
                listOfText.Add(text1);
                listOfText.Add(text2);
                listOfText.Add(text3);
                listOfText.Add(text4);
                listOfText.Add(text5);
                NailPlacement = textOnImage(NailPlacement, listOfText);

            }
            else if (nailerInfo.Pattern == "Staggered")
            {
                var text1 = new Tuple<string, int, int, int>(nailSpacing, 22, 253, 140);
                var text2 = new Tuple<string, int, int, int>(nailSpacing, 22, 1073, 140);
                var text3 = new Tuple<string, int, int, int>(nailSpacing, 22, 420, 338);
                var text4 = new Tuple<string, int, int, int>(nailSpacing, 22, 890, 338);
                var text5 = new Tuple<string, int, int, int>(halfScrewSpace_Inch + "\"\r\nMAX", 20, 182, 328);
                var text6 = new Tuple<string, int, int, int>(halfScrewSpace_Inch + "\"\r\nMAX", 20, 1183, 328);
                var text7 = new Tuple<string, int, int, int>(woodThickness, 22, 600, 555);
                listOfText.Add(text1);
                listOfText.Add(text2);
                listOfText.Add(text3);
                listOfText.Add(text4);
                listOfText.Add(text5);
                listOfText.Add(text6);
                listOfText.Add(text7);

                NailPlacement = textOnImage(NailPlacement, listOfText);

            }
            else { }
            Clipboard.SetImage(NailPlacement);



            selection.Paste();
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.Text = "\r\n";
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);


            bool areAsEqual = StringManipulation.areStringElementsEqual(stringListLengthA); //tboxListA
            bool areBsEqual = StringManipulation.areStringElementsEqual(stringListLengthB); //tboxListB
            bool areWoodLengthsEqual = StringManipulation.areStringElementsEqual(woodLengths1);
            bool areWoodWithsEqual = StringManipulation.areStringElementsEqual(woodWidths);


            if (areAsEqual && areBsEqual && areWoodLengthsEqual && areWoodWithsEqual == true)
            {
                Word.Table tableNailBacksheetALL = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 2, 7);

                tableNailBacksheetALL.Cell(1, 1).Range.Text = "MARK(S)";
                tableNailBacksheetALL.Cell(1, 2).Range.Text = "QTY.";
                tableNailBacksheetALL.Cell(1, 3).Range.Text = "A";
                tableNailBacksheetALL.Cell(1, 4).Range.Text = "B";
                tableNailBacksheetALL.Cell(1, 5).Range.Text = "WOOD SIZE";
                tableNailBacksheetALL.Cell(1, 6).Range.Text = "WOOD OAL";
                tableNailBacksheetALL.Cell(1, 7).Range.Text = "REMARKS";

                tableNailBacksheetALL.Cell(2, 1).Range.Text = "ALL";
                tableNailBacksheetALL.Cell(2, 2).Range.Text = joistCoverSheetInfo.TotalQuantity.ToString();
                tableNailBacksheetALL.Cell(2, 3).Range.Text = stringListLengthA[0]; //tboxListA
                tableNailBacksheetALL.Cell(2, 4).Range.Text = stringListLengthB[0]; //tboxListB
                var thickness = (nailerInfo.WoodThickness == "3x") ? "2.5" : "2";
                tableNailBacksheetALL.Cell(2, 5).Range.Text = woodWidths[0] + " x " + thickness;
                tableNailBacksheetALL.Cell(2, 6).Range.Text = woodLengths1[0];
                if (nailerInfo.WoodThickness == "2\"")
                {
                    tableNailBacksheetALL.Cell(2, 7).Range.Text = "** 2\" WOOD";
                }

                for (int i = 1; i <= 7; i++)
                {
                    tableNailBacksheetALL.Cell(1, i).Range.Font.Bold = 1;
                }

                for (int row = 1; row <= numberOfMarks + 1; row++)
                    for (int column = 1; column <= 7; column++)
                    {
                        tableNailBacksheetALL.Cell(row, column).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                selection.EndKey(Word.WdUnits.wdStory, 1);

                tableNailBacksheetALL.Borders.Enable = 1;

            }
            else
            {
                Word.Table tableNailBacksheet = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, numberOfMarks + 1, 7);

                tableNailBacksheet.Cell(1, 1).Range.Text = "MARK(S)";
                tableNailBacksheet.Cell(1, 2).Range.Text = "QTY.";
                tableNailBacksheet.Cell(1, 3).Range.Text = "A";
                tableNailBacksheet.Cell(1, 4).Range.Text = "B";
                tableNailBacksheet.Cell(1, 5).Range.Text = "WOOD WIDTH";
                tableNailBacksheet.Cell(1, 6).Range.Text = "WOOD OAL";
                tableNailBacksheet.Cell(1, 7).Range.Text = "REMARKS";



                for (int i = 1; i <= 7; i++)
                {
                    tableNailBacksheet.Cell(1, i).Range.Font.Bold = 1;
                }


                for (int i = 0; i < numberOfMarks ; i++)
                {
                    var soSummaryLine = joistCoverSheetInfo.SoSummary[i];
                    tableNailBacksheet.Cell(i + 2, 1).Range.Text = soSummaryLine.Mark;
                    tableNailBacksheet.Cell(i + 2, 2).Range.Text = soSummaryLine.Quantity.ToString();
                    tableNailBacksheet.Cell(i + 2, 6).Range.Text = woodLengths1[i];
                    tableNailBacksheet.Cell(i + 2, 5).Range.Text = woodWidths[i];
                    tableNailBacksheet.Cell(i + 2, 3).Range.Text = stringListLengthA[i]; //tboxlistA
                    tableNailBacksheet.Cell(i + 2, 4).Range.Text = stringListLengthB[i]; //tboxListB
                    if (nailerInfo.WoodThickness == "2\"")
                    {
                    tableNailBacksheet.Cell(i + 2, 7).Range.Text = "** 2\" WOOD";
                    }

                }

                for (int row = 1; row <= numberOfMarks + 1; row++)
                    for (int column = 1; column <= 7; column++)
                    {
                        tableNailBacksheet.Cell(row, column).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                selection.EndKey(Word.WdUnits.wdStory, 1);

                tableNailBacksheet.Borders.Enable = 1;

                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                selection.MoveUp(Word.WdUnits.wdLine, 1);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selection.Text = "\r\n" + "\r\n";
                selection.HomeKey(Word.WdUnits.wdLine);
                selection.MoveUp(Word.WdUnits.wdLine, 6);

            }

            var nailerSheetNotes = new List<string>();

            if (nailerInfo.CrimpedWebs == false)
            {
                var tcWidth = woodWidths[0].Replace(" 1/4", "");
                nailerSheetNotes.Add("SPECIAL WOOD NAILER; " + tcWidth + " 1/4\" WOOD WIDTH");
            }

            if (nailerInfo.WoodThickness == "2\"")
            {
                nailerSheetNotes.Add("SPECIAL WOODNAILER: USE 2\" WOOD");
            }

            if (TCs.Find(isComposite) != null)
            {
                nailerSheetNotes.Add("COMPOSITE WOODNAILER; THIS JOIST MUST BE BUILT IN BLDG 2");
                nailerSheetNotes.Add("WOOD MUST BE CONTINUOUS, FINGER JOINTS ONLY, NO BUTTED JOINTS");
            }

            if (nailerInfo.IsPanelized)
            {
                nailerSheetNotes.Add("PANELIZED INC.; NO MID-SPAN JOINTS");
            }

            var nailerSheetNote = String.Join("\r\n", nailerSheetNotes);

            addTextBox(selection, 130, 700, 350, 20, nailerSheetNote);


            try { Clipboard.SetText(clipboard); }
            catch { }


            this.Close();

        }

        private void comboBoxNailPlacement_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tBoxDWGBY_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBoxExcelData_CheckedChanged(object sender, EventArgs e)
        {

        }

        public Image textOnImage(Image image, string text, int fontSize, int x, int y)
        {
            Bitmap bitMapImage = new
                System.Drawing.Bitmap(image);
            Graphics graphicImage = Graphics.FromImage(bitMapImage);

            graphicImage.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            Font font1 = new Font("Calibri", fontSize, FontStyle.Regular);
            Point point1 = new Point(x, y);
            graphicImage.DrawString(text, font1, SystemBrushes.WindowText, point1);

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

        public int scaleHeight(int y)
        {
            double heightScale = SystemInformation.PrimaryMonitorSize.Height / 1080.0;

            int scaleHeight = Convert.ToInt32(heightScale * y);

            return scaleHeight;
        }

        public int scaleWidth(int x)
        {
            double widthscale = SystemInformation.PrimaryMonitorSize.Width / 1920.0;

            int scaleWidth = Convert.ToInt32(widthscale * x);

            return scaleWidth;
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

        private void addTextBox(Word.Selection selection, float left, float top, float width, float height, string text)
        {
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            Word.Shape wdNailerTextBox = Globals.ThisAddIn.Application.ActiveDocument.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            wdNailerTextBox.TextFrame.TextRange.Bold = 1;
            wdNailerTextBox.TextFrame.TextRange.Font.Size = 8;
            wdNailerTextBox.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wdNailerTextBox.TextFrame.MarginTop = 1;
            wdNailerTextBox.TextFrame.MarginLeft = 1;
            wdNailerTextBox.TextFrame.MarginBottom = 1;
            wdNailerTextBox.TextFrame.MarginRight = 1;
            wdNailerTextBox.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0;
            wdNailerTextBox.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0;
            wdNailerTextBox.TextFrame.TextRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            wdNailerTextBox.TextFrame.ContainingRange.Text = text;
            wdNailerTextBox.TextFrame.AutoSize = -1;


            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.HomeKey(Word.WdUnits.wdStory);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
        }

        private void tboxTolerance_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
