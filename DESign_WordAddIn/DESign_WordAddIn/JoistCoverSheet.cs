using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using DESign_WordAddIn;
using DESign_BASE;

namespace DESign_WordAddIn
{
    public class SoSummaryLine
    {
        public string Mark { get; set; }
        public int Quantity { get; set; }
        public double Length { get; set; }
        public double Depth { get; set; }
        public string Tc { get; set; }
        public string Bc { get; set; }

        public SoSummaryLine(string mark, int quantity, double length, double depth, string tc, string bc)
        {
            Mark = mark;
            Quantity = quantity;
            Length = length;
            Depth = depth;
            Tc = tc;
            Bc = bc;

        }
    }


    public class CoverSheetInfo
    {
        public List<SoSummaryLine> SoSummary { get; set; }
        public string JobNumber { get; set; }

        public string JobName { get; set; }

        public string JobLocation { get; set; }

        public string ListNumber { get; set; }

        public int TotalQuantity
        {
            get
            {
                return SoSummary.Select(s => s.Quantity).Sum();
            }
        }

        public List<string> Marks
        {
            get
            {
                return SoSummary.Select(s => s.Mark).ToList();
            }
        }

        public List<string> Tcs
        {
            get
            {
                return SoSummary.Select(s => s.Tc).ToList();
            }
        }

        public List<double> Lengths
        {
            get
            {
                return SoSummary.Select(s => s.Length).ToList();
            }
        }

        public CoverSheetInfo(List<SoSummaryLine> soSummary, string jobNumber, string jobName, string jobLocation, string listNumber)
        {
            SoSummary = soSummary;
            JobNumber = jobNumber;
            JobName = jobName;
            JobLocation = jobLocation;
            ListNumber = listNumber;

        }

        public static CoverSheetInfo FromV1So()
        {
            var soSummary = new List<SoSummaryLine>();


            Word.Selection selection = Globals.ThisAddIn.Application.Selection;

            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Find.Execute("MARK ");

            selection.MoveDown(Word.WdUnits.wdLine, 2);

            string[] joistDataArray = null;
            do
            {
                string mark = null;
                int quantity = 0;
                double length = 0.0;
                double depth = 0.0;
                string tc = null;
                string bc = null;

                selection.HomeKey(Word.WdUnits.wdLine);
                selection.EndKey(Extend: Word.WdUnits.wdLine);
                selection.Copy();
                string joistData = selection.Text;
                Char[] delimChars = { ' ', '\u00A0' };
                joistDataArray = joistData.Split(delimChars, StringSplitOptions.RemoveEmptyEntries);

                if (joistDataArray.Length > 7)
                {
                    mark = joistDataArray[0];
                    quantity = System.Int32.Parse(joistDataArray[1]);

                    bool lengthWithFraction = joistDataArray[4].Contains('/');
                    int subtractIndex = 0;
                    if (lengthWithFraction == true)
                    {
                        string joistLength = joistDataArray[3] + " " + joistDataArray[4];
                        length = StringManipulation.hyphenLengthToDecimal(joistLength);
                    }
                    else if (lengthWithFraction == false)
                    {
                        length = StringManipulation.hyphenLengthToDecimal(joistDataArray[3]);
                        subtractIndex = -1;
                    }

                    depth = StringManipulation.hyphenLengthToDecimal(joistDataArray[8 + subtractIndex]);

                    string TCBC = joistDataArray[10 + subtractIndex];
                    Char[] TCBCdelimChars = { '/' };
                    string[] arrayTCBC = TCBC.Split(TCBCdelimChars, StringSplitOptions.RemoveEmptyEntries);

                    tc = arrayTCBC[0];
                    bc = arrayTCBC[1];
                }
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selection.Find.Execute("\u00A0");

                var soSummaryLine = new SoSummaryLine(mark, quantity, length, depth, tc, bc);
                soSummary.Add(soSummaryLine);

            } while (joistDataArray.Length >= 7);

            //GATHERING OTHER INFORMATION
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.MoveDown(Word.WdUnits.wdLine, 2);
            selection.HomeKey(Word.WdUnits.wdLine, 1);
            selection.EndKey(Extend: Word.WdUnits.wdLine);
            selection.Copy();
            string jobTitle = selection.Text;
            string[] jobTitleArray = jobTitle.Split(new string[] { "  ", "\u00A0" }, StringSplitOptions.RemoveEmptyEntries);

            string jobNumber = null;
            string jobName = null;
            string jobLocation = null;
            if (jobTitleArray.Length == 4)
            {
                jobNumber = jobTitleArray[0];
                jobName = jobTitleArray[1];
                jobLocation = jobTitleArray[2];
            }
            else
            {
                jobNumber = jobTitleArray[0];
                jobName = jobTitleArray[1].Substring(0, 30);
                jobLocation = " ";
            }

            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.Find.Execute("LIST");
            selection.HomeKey(Word.WdUnits.wdLine, 1);
            selection.EndKey(Extend: Word.WdUnits.wdLine);
            selection.Copy();
            string listNumberLine = selection.Text;
            string[] ListNumberLineArray = listNumberLine.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            string listNumber = ListNumberLineArray[2];

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);

            //END OF GATHERING OTHER INFORMATION

            soSummary = soSummary.Where(s => s.Mark != null).ToList();
            var joistCoverSheetInfo = new CoverSheetInfo(soSummary, jobNumber, jobName, jobLocation, listNumber);
            return joistCoverSheetInfo;
        }
    }
}