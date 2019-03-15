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
using Word = Microsoft.Office.Interop.Word;
using System.Drawing.Text;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Collections;
using DESign_BOT;
using Ookii.Dialogs.WinForms;
using System.Text.RegularExpressions;






namespace DESign_BOT
{
    class FolderOperations
    {
        StringManipulation stringManipulation = new StringManipulation();
        public List<string> GetShopOrderFiles(string[] listOfStringsThatFileCantContain)
        {
            List<string> filePaths = new List<string>();

            var fsd = new VistaFolderBrowserDialog();
            fsd.Description = "Select Directory To Shop Orders";
            fsd.UseDescriptionForTitle = true;
            fsd.SelectedPath = @"\\nmbsfaln-fs\engr\_MANUALSHOPORDERS\";


            string folder = null;

            if (fsd.ShowDialog() == DialogResult.OK)
            {
                folder = fsd.SelectedPath;
            }

            string[] files = System.IO.Directory.GetFiles(folder);
            string extension;
            string path;
            string fileName;

            foreach (string s in files)
            {

                path = System.IO.Path.GetFullPath(s);
                extension = System.IO.Path.GetExtension(path);
                fileName = System.IO.Path.GetFileName(s);

                bool fileIsAllowed = true;
                foreach (var str in listOfStringsThatFileCantContain)
                {
                    if (fileName.Contains(str))
                    {
                        fileIsAllowed = false;
                    }
                }


                if (fileName.Contains("~$") == false && fileIsAllowed)
                {
                    if (extension == ".docx" || extension == ".doc" || extension == ".rtf")
                    {
                        filePaths.Add(path);
                    }
                }

            }

            return filePaths;



        }

        public List<List<List<string>>> AllJoistData() 
        {
            

            List<List<List<string>>> AllJoistData = new List<List<List<string>>>();
            List<string> filePaths = GetShopOrderFiles(new string[] { "AC", "G" });

            string totalFiles = Convert.ToString(filePaths.Count());
            int fileCount = 0;

            Form fileOpenCounter = new Form();
               fileOpenCounter.Size = new Size(235, 100);
               fileOpenCounter.Text = "FILE COUNTER";
            Label fileCounterLabel = new Label();
               fileCounterLabel.Text = "OPENED FILES:";
               fileCounterLabel.AutoSize = true;
               fileCounterLabel.Location = new Point(12, 20);
            TextBox status = new TextBox();
               status.Size = new Size(100, 20);
               status.Location = new Point(110, 16);
               status.ReadOnly = true;

            fileOpenCounter.Controls.Add(fileCounterLabel);
            fileOpenCounter.Controls.Add(status);

            fileOpenCounter.Show();

           
            


            foreach (string file in filePaths)
            {
                fileCount++;
                status.Text = String.Format("{0}/{1}", Convert.ToString(fileCount), totalFiles);


                string copiedFile = System.IO.Path.GetTempFileName();
                Byte[] sInByteArray = System.IO.File.ReadAllBytes(file);
                System.IO.File.WriteAllBytes(copiedFile, sInByteArray);

                Word.Application wordApp;
                Word.Document wordDoc;

                wordApp = new Word.Application();
                wordApp.Visible = false;

                wordDoc = wordApp.Documents.Open(copiedFile);

                Word.Selection selection = wordApp.Selection;

                List<string> joistMarks = new List<string>();
                List<string> joistQuantities = new List<string>();
                List<string> joistLengths = new List<string>();
                List<string> joistDepth = new List<string>();
                List<string> joistTC = new List<string>();
                List<string> joistBC = new List<string>();
                List<string> joistDescription = new List<string>();
                List<string> joistBaseLength = new List<string>();
                List<string> joistTCWidth = new List<string>();



                selection.HomeKey(Word.WdUnits.wdStory, 0);
                selection.Find.Execute("MARK");

                selection.MoveDown(Word.WdUnits.wdLine, 2);
                string baseLength;
                string[] joistDataArray = null;
                do
                {
                    selection.HomeKey(Word.WdUnits.wdLine);
                    selection.EndKey(Extend: Word.WdUnits.wdLine);
                    selection.Copy();
                    string joistData = selection.Text;
                    Char[] delimChars = { ' ', '\u00A0' };
                    joistDataArray = joistData.Split(delimChars, StringSplitOptions.RemoveEmptyEntries);

                    if (joistDataArray.Length > 7)
                    {
                        joistMarks.Add(joistDataArray[0]);
                        joistQuantities.Add(joistDataArray[1]);
                        joistDescription.Add(joistDataArray[2]);

                        bool lengthWithFraction = joistDataArray[4].Contains('/');
                        int subtractIndex = 0;
                        if (lengthWithFraction == true)
                        {
                            string joistLength = joistDataArray[3] + " " + joistDataArray[4];
                            joistLengths.Add(joistLength);
                        }
                        else if (lengthWithFraction == false)
                        {
                            joistLengths.Add(joistDataArray[3]);
                            subtractIndex = -1;
                        }

                        joistDepth.Add(joistDataArray[8 + subtractIndex]);

                        string TCBC = joistDataArray[10 + subtractIndex];
                        Char[] TCBCdelimChars = { '/' };
                        string[] arrayTCBC = TCBC.Split(TCBCdelimChars, StringSplitOptions.RemoveEmptyEntries);
                        string TC = arrayTCBC[0];

                        joistTC.Add(arrayTCBC[0]);
                        joistBC.Add(arrayTCBC[1]);

                        /////
                        Word.Range wordRange = wordDoc.Range();
                        wordRange.SetRange(selection.Range.Start, selection.Range.End);

                        selection.EndKey(Word.WdUnits.wdLine, false);
                        selection.Find.Execute(joistDataArray[0] + "  ");
                        selection.EndKey(Word.WdUnits.wdLine, false);
                        selection.Find.Execute("TOTAL LENGTH");
                        selection.HomeKey(Word.WdUnits.wdLine, false);
                        selection.MoveDown(Word.WdUnits.wdLine, 1);
                        selection.EndKey(Word.WdUnits.wdLine, true);

                        selection.Copy();
                        string selectionText = selection.Text;
                        string[] selectionTextArray = selectionText.Split(new string[] { "  ", "\u00A0", "\v" }, StringSplitOptions.RemoveEmptyEntries);

                        string totalLength = selectionTextArray[0];
                        string TCXL = selectionTextArray[1];
                        string TCXR = selectionTextArray[3];

                        baseLength =
                            stringManipulation.DecimilLengthToHyphen(
                                stringManipulation.hyphenLengthToDecimal(totalLength)
                                - stringManipulation.hyphenLengthToDecimal(TCXL)
                                - stringManipulation.hyphenLengthToDecimal(TCXR)
                            );
                        joistBaseLength.Add(baseLength);

                        selection.SetRange(wordRange.Start, wordRange.End);

                        string woodWidth = null;

                        string firstTwoChar = TC.Substring(0, 2);

                        if (TC.Contains("18") == true | TC.Contains("A30") == true) { woodWidth = "5\""; }
                        else if (TC.Contains("29") == true) { woodWidth = "7 1/8\""; }
                        else if (TC.Contains("A44") == true) { woodWidth="7\""; }
                        else if (TC != "A28" && TC.Contains("A28") == true) { woodWidth="7\""; }
                        else if (firstTwoChar.Contains("30")) { woodWidth = "7\""; }
                        else if (firstTwoChar.Contains("35")) { woodWidth = "8\""; }
                        else if (firstTwoChar.Contains("40")) { woodWidth = "9\""; }
                        else if (firstTwoChar.Contains("50")) { woodWidth = "11\""; }
                        else if (firstTwoChar.Contains("60")) { woodWidth = "13\""; }
                        else { woodWidth = " "; }



                        joistTCWidth.Add(woodWidth);

                    }
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    selection.Find.Execute("\u00A0");





                } while (joistDataArray.Length >= 7);

                //GATHERING OTHER INFORMATION
                selection.HomeKey(Word.WdUnits.wdStory, 0);
                selection.MoveDown(Word.WdUnits.wdLine, 2);
                selection.HomeKey(Word.WdUnits.wdLine, 1);
                selection.EndKey(Extend: Word.WdUnits.wdLine);
                selection.Copy();
                string jobTitle = selection.Text;
                string[] jobTitleArray = jobTitle.Split(new string[] { "  ", "\u00A0" }, StringSplitOptions.RemoveEmptyEntries);


                List<string> jobNumber = new List<string>();
                jobNumber.Add(jobTitleArray[0]);
                List<string> jobName = new List<string>();
                jobName.Add(jobTitleArray[1]);
                List<string> jobLocation = new List<string>();
                jobLocation.Add(jobTitleArray[2]);


                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.HomeKey(Word.WdUnits.wdStory, 0);





                //END OF GATHERING OTHER INFORMATION



                List<List<string>> JoistData = new List<List<string>>();

                JoistData.Add(joistMarks);
                JoistData.Add(joistQuantities);
                JoistData.Add(joistLengths);
                JoistData.Add(joistDepth);
                JoistData.Add(joistTC);
                JoistData.Add(joistBC);
                JoistData.Add(joistDescription);
                JoistData.Add(joistBaseLength);
                JoistData.Add(joistTCWidth);
                JoistData.Add(jobName);
                JoistData.Add(jobNumber);
                JoistData.Add(jobLocation);

 /*             JoistData.Add(ListTotalJoistQuantity);
                JoistData.Add(jobNumber);
                JoistData.Add(jobName);
                JoistData.Add(jobLocation);
                JoistData.Add(listNumber);
*/
                AllJoistData.Add(JoistData); 
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.HomeKey(Word.WdUnits.wdStory, 0);

                wordDoc.Close();
                wordApp.Quit();
            }

            fileOpenCounter.Close();
            return AllJoistData;
        }
        public class JoistSummary
        {
            public string Mark { get; }
            public string Sequence { get; }
            public int Quantity { get; }

            public JoistSummary(string mark, int quantity, string sequence)
            {
                this.Mark = mark;
                this.Sequence = sequence;
                this.Quantity = quantity;
            }
        }

        public (string, List<JoistSummary>) GetJoistSummaries()
        {


            List<JoistSummary> joistSummaries = new List<JoistSummary>();
            List<string> filePaths = GetShopOrderFiles(new string[] { "AC", "G" });

            string totalFiles = Convert.ToString(filePaths.Count());
            int fileCount = 0;

            Form fileOpenCounter = new Form();
            fileOpenCounter.Size = new Size(235, 100);
            fileOpenCounter.Text = "FILE COUNTER";
            Label fileCounterLabel = new Label();
            fileCounterLabel.Text = "OPENED FILES:";
            fileCounterLabel.AutoSize = true;
            fileCounterLabel.Location = new Point(12, 20);
            TextBox status = new TextBox();
            status.Size = new Size(100, 20);
            status.Location = new Point(110, 16);
            status.ReadOnly = true;

            fileOpenCounter.Controls.Add(fileCounterLabel);
            fileOpenCounter.Controls.Add(status);

            fileOpenCounter.Show();

            Word.Application wordApp;
            wordApp = new Word.Application();
            wordApp.Visible = false;
            Word.Document wordDoc = null;

            var jobNumber = "";

            foreach (string file in filePaths)
            {
                fileCount++;
                status.Text = String.Format("{0}/{1}", Convert.ToString(fileCount), totalFiles);

                wordDoc = wordApp.Documents.Open(file, ReadOnly: true);

                var docText = wordDoc.Range().Text;
                var summaryText = docText.Substring(0, docText.IndexOf("MATERIAL"));

                Regex sequenceRegex = new Regex(@"([a-zA-Z\d]+-[a-zA-Z\d]+)-([a-zA-Z\d]+) +");
                var sequenceMatch = sequenceRegex.Match(summaryText);
                jobNumber = sequenceMatch.Success ? sequenceMatch.Groups[1].Value : "";
                var sequence = sequenceMatch.Success ? Regex.Split(sequenceMatch.Groups[2].Value, @"\D+")[0] : "";

                Regex joistSummaryRegex = new Regex(@"([a-zA-Z\d]+) +(\d+) +([\d|\.]+[LH|K|DLH|G|BG|VG|KA|]+[\d|\/]+) +(\d+-\d+ ?\d?\/?\d?) +\d+ +");
                var joistSummaryMatches = joistSummaryRegex.Matches(summaryText);
                for (int i = 0; i < joistSummaryMatches.Count; i++)
                {
                    var mark = joistSummaryMatches[i].Groups[1].Value;
                    var quantity = short.Parse(joistSummaryMatches[i].Groups[2].Value);
                    joistSummaries.Add(new JoistSummary(mark, quantity, sequence));
                }

            }
            wordDoc.Close();
            wordApp.Quit();
            fileOpenCounter.Close();
            var jobInfo = (JobNumber: jobNumber, JoistSummaries: joistSummaries);
            return jobInfo;
        }


        public string WoodRequirements()
        {

            List<List<List<string>>> allJoistData = AllJoistData();


            double fiveInch = 0;
            double sevenInch = 0;
            double eightInch = 0;
            double nineInch = 0;
            double elevenInch = 0;
            double thirteenInch = 0;

            for (int i = 0; i < allJoistData.Count(); i++)
            {

                List<List<string>> shopOrderJoistData = allJoistData[i];

                List<string> shopOrderjoistMarks = shopOrderJoistData[0];
                List<string> shopOrderJoistQuantities = shopOrderJoistData[1];
                int shopOrdernumberOfMarks = shopOrderJoistData[0].Count;

                List<string> joistLengths = shopOrderJoistData[2];
                List<string> TCs = shopOrderJoistData[4];


                double joistLength = 0;
                double joistQuantity = 0;
                double totalLength = 0;
                for (int k = 0; k < shopOrdernumberOfMarks; k++)
                {
                    joistQuantity = Convert.ToDouble(shopOrderJoistQuantities[k]);
                    joistLength = stringManipulation.hyphenLengthToDecimal(joistLengths[k]);
                    totalLength = joistQuantity * joistLength;
                    string woodWidth=shopOrderJoistData[8][k];
/*
                    string firstTwoChar = TCs[k].Substring(0, 2);
                    bool contains18 = TCs[k].Contains("1.8");

                    if (TCs[k].Contains("18") == true | TCs[k].Contains("A30") == true) { woodWith = "5\""; }
                    else if (TCs[k].Contains("29") == true) { woodWith = "7 1/8\""; }
                    else if (TCs[k].Contains("A44") == true) { woodWith = "7\""; }
                    else if (TCs[k] != "A28" && TCs[i].Contains("A28") == true) { woodWith="7\""; }
                    else if (firstTwoChar.Contains("30")) { woodWith ="7\""; }
                    else if (firstTwoChar.Contains("35")) { woodWith ="8\""; }
                    else if (firstTwoChar.Contains("40")) { woodWith ="9\""; }
                    else if (firstTwoChar.Contains("50")) { woodWith ="10\""; }
                    else if (firstTwoChar.Contains("60")) { woodWith ="13\""; }
                    else { woodWith =" "; }
*/
                    if (woodWidth == "5\"") { fiveInch = fiveInch + totalLength; }
                    else if (woodWidth == "7 1/8\"") { sevenInch = sevenInch + totalLength; }
                    else if (woodWidth == "7\"") { sevenInch = sevenInch + totalLength; }
                    else if (woodWidth == "8\"") { eightInch = eightInch + totalLength; }
                    else if (woodWidth == "9\"") { nineInch = nineInch + totalLength; }
                    else if (woodWidth == "11\"") { elevenInch = elevenInch + totalLength; }
                    else if (woodWidth == "13\"") { thirteenInch = thirteenInch + totalLength; }
                 

                    else
                    {
                        MessageBox.Show("Can't determine width of joist " + shopOrderjoistMarks[k] + " ; its length is not included");
                    }


                }



               
            }
            string stringFiveInch, stringSevenInch, stringEightInch, stringNineInch, stringElevenInch, stringThirteenInch;
            stringFiveInch = stringSevenInch = stringEightInch = stringNineInch = stringElevenInch = stringThirteenInch = String.Empty;
            if (fiveInch !=0)
            {
                stringFiveInch = "5\" = " + Convert.ToString(Convert.ToInt32(fiveInch)) + "  lf \r\n";
            }
            if (sevenInch != 0)
            {
                stringSevenInch = "7\" = " + Convert.ToString(Convert.ToInt32(sevenInch)) + "  lf \r\n";
            }
            if (eightInch != 0)
            {
                stringEightInch = "8\" = " + Convert.ToString(Convert.ToInt32(eightInch)) + "  lf \r\n";
            }
            if (nineInch != 0)
            {
                stringNineInch = "9\" = " + Convert.ToString(Convert.ToInt32(nineInch)) + "  lf \r\n";
            }
            if (elevenInch != 0)
            {
                stringElevenInch = "11\" = " + Convert.ToString(Convert.ToInt32(elevenInch)) + "  lf \r\n";
            }
            if (thirteenInch != 0)
            {
                stringThirteenInch = "13\" = " + Convert.ToString(Convert.ToInt32(thirteenInch)) + "  lf \r\n";
            }
            /*
            string woodRequirements =

                "5\" = " + Convert.ToString(Convert.ToInt16(fiveInch)) + "  lf \r\n" +
                "7 1/8\" = " + Convert.ToString(Convert.ToInt16(sevenInch)) + "  lf \r\n" +
                "8 1/8\" = " + Convert.ToString(Convert.ToInt16(eightInch)) + "  lf \r\n" +
                "9 1/8\" = " + Convert.ToString(Convert.ToInt16(nineInch)) + "  lf \r\n" +
                "10 1/8\" = " + Convert.ToString(Convert.ToInt16(tenInch)) + "  lf \r\n" +
                "11 1/8\" = " + Convert.ToString(Convert.ToInt16(elevenInch)) + "  lf \r\n";
            */
            string woodRequirements =

                stringFiveInch + stringSevenInch + stringEightInch + stringNineInch + stringElevenInch + stringThirteenInch;
                

            return woodRequirements;
        }

        public List<List<string>> JoistDataByMark()
        {
            
            List<List<List<string>>> allJoistData = AllJoistData();

            List<List<string>> shopOrderData = new List<List<string>>();

            List<List<string>> joistDataByMark = new List<List<string>>();

            for (int i = 0; i < allJoistData.Count(); i++)
            {


                shopOrderData = allJoistData[i];

                for (int k=0; k<shopOrderData[0].Count(); k++)
                {
                    List<string> markData = new List<string>();
                    markData.Add(shopOrderData[0][k]);
                    markData.Add(shopOrderData[1][k]);
                    markData.Add(shopOrderData[6][k]);
                    markData.Add(shopOrderData[7][k]);
                    markData.Add(shopOrderData[8][k]);
                    markData.Add(shopOrderData[9][0]);
                    markData.Add(shopOrderData[10][0]);
                    markData.Add(shopOrderData[11][0]);
                    joistDataByMark.Add(markData);
                }



            }
            return joistDataByMark;

        }

        public List<List<string>> arrangedJoistDataByMark()
        {

            List<List<List<string>>> allJoistData = AllJoistData();

            List<List<string>> shopOrderData = new List<List<string>>();

            List<List<string>> joistDataByMark = new List<List<string>>();

            int numberOfMarks=0;

            for (int j=0; j<allJoistData.Count(); j++)
            {
                shopOrderData = allJoistData[j];
                numberOfMarks = numberOfMarks + shopOrderData[0].Count();
            }


            string joistDataString = null;

            string[] arrayJoistDataStrings = new string[numberOfMarks];
            int counter = 0;

            for (int i = 0; i < allJoistData.Count(); i++)
            {
               

                shopOrderData = allJoistData[i];

                for (int k = 0; k < shopOrderData[0].Count(); k++)
                {
                    
                    joistDataString = string.Format("{0}   {1}   {2}   {3}   {4}   {5}   {6}   {7}", 
                             shopOrderData[0][k], shopOrderData[1][k], 
                             shopOrderData[6][k], shopOrderData[7][k], 
                             shopOrderData[8][k], shopOrderData[9][0],
                             shopOrderData[10][0],shopOrderData[11][0]);

                    arrayJoistDataStrings[counter] = joistDataString;
                    counter = counter + 1;
                }



            }

            Array.Sort(arrayJoistDataStrings, new AlphanumComparatorFast());

            for (int i = 0; i < arrayJoistDataStrings.Count(); i++)
            {
                    List<string> markData = new List<string>();

                string[] markDataArray = arrayJoistDataStrings[i].Split(new string[] {"   "}, StringSplitOptions.RemoveEmptyEntries);
                if (markDataArray.Count() == 8)
                {
                    markData.Add(markDataArray[0]);
                    markData.Add(markDataArray[1]);
                    markData.Add(markDataArray[2]);
                    markData.Add(markDataArray[3]);
                    markData.Add(markDataArray[4]);
                    markData.Add(markDataArray[5]);
                    markData.Add(markDataArray[6]);
                    markData.Add(markDataArray[7]);
                    joistDataByMark.Add(markData);
                }
                else
                {
                    markData.Add(markDataArray[0]);
                    markData.Add(markDataArray[1]);
                    markData.Add(markDataArray[2]);
                    markData.Add(markDataArray[3]);
                    markData.Add("");
                    markData.Add(markDataArray[4]);
                    markData.Add(markDataArray[5]);
                    markData.Add(markDataArray[6]);
                    joistDataByMark.Add(markData);
                }
                



            }
            return joistDataByMark;

        }

        public void showJoistDataByMark()
        {
            List<List<string>> joistDataByMark = JoistDataByMark();

            for (int i = 0; i< joistDataByMark.Count(); i++)
            {
                List<string> markData = joistDataByMark[i];
                MessageBox.Show
                    (
                        string.Format("{0}  {1}  {2}  {3}  {4}", markData[0], markData[1], markData[2], markData[3], markData[4])
                    );
            }
        }

        public void createTCWidthDocument ()
        {

            List<List<string>> joistDataByMark = arrangedJoistDataByMark();

            Word.Application wordApp = new Word.Application();
            wordApp.Visible = false;

            Word.Document wordDoc = new Word.Document();

            wordDoc.Activate();

            wordDoc.Paragraphs.SpaceAfter = 0;
           ///
            Word.Selection selection = wordApp.Application.Selection;


            selection.Font.Name = "Calibri";
            selection.Font.Size = 11;
            selection.Font.Bold = 1;

            Word.Section section = selection.Sections.Add();

            Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

            string jobNumber = joistDataByMark[0][6];

            string[] jobNumberArray = jobNumber.Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);

            jobNumber = jobNumberArray[0] + "-" + jobNumberArray[1];
            

            headerRange.Text = "TOP CHORD WIDTHS: "+ joistDataByMark[0][5]  + " [NMBS: " + jobNumber + "]" + "     " + joistDataByMark[0][7] +
                "\r\n\r\n" + "  MARK         QUANTITY      DESCRIPTION              BASE LENGTH        TC WIDTH";

            Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
            footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            selection.HomeKey(Word.WdUnits.wdStory, 0);
            //selection.Text = "TOP CHORD WIDTHS: \r\n";

           // selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
           // selection.MoveDown(Word.WdUnits.wdLine, 1);

            selection.Font.Bold = 0;

            Word.Table tableTCWidths = wordApp.ActiveDocument.Tables.Add(selection.Range,joistDataByMark.Count()+1,5);


            tableTCWidths.Cell(1, 1).Range.Text = "";
            tableTCWidths.Cell(1, 2).Range.Text = "";
            tableTCWidths.Cell(1, 3).Range.Text = "";
            tableTCWidths.Cell(1, 4).Range.Text = "";
            tableTCWidths.Cell(1, 5).Range.Text = "";


            for (int i = 1; i <= 5; i++)
            {
                tableTCWidths.Cell(1, i).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                //tableTCWidths.Cell(1, i).Range.Underline = Word.WdUnderline.wdUnderlineSingle;
                tableTCWidths.Cell(1, i).Range.Font.Bold = 1;
                //    tableNailBackSheetTitle.Cell(i,1).Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            }

            int rowcounter = 0;
            int numRows = joistDataByMark.Count()+1;
            for (int row = 2; row <= numRows; row++)
            {
                for (int col = 1; col<=5; col++)
                {

                    tableTCWidths.Cell(row, col).Range.Text = joistDataByMark[row - 2][col - 1];
         
                }
            }

            

            for (int i = 1; i <= 4; i++)
            {
                    tableTCWidths.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }

            tableTCWidths.Columns[1].Width = 50;
            tableTCWidths.Columns[2].Width = 60;
            tableTCWidths.Columns[3].Width = 95;
            tableTCWidths.Columns[4].Width = 80;
            tableTCWidths.Columns[5].Width = 60;


            for (int row =1; row<=joistDataByMark.Count()+1; row++)
            {
                tableTCWidths.Rows[row].Height = 7;
            }

            for (int row = 1; row <= joistDataByMark.Count() + 1; row++)
            {
                for (int col = 1; col <= 5; col++)
                {
                    tableTCWidths.Cell(row, col).Range.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tableTCWidths.Cell(row, col).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                }
            }
            wordApp.Visible = true;
        }
        

    }
}

