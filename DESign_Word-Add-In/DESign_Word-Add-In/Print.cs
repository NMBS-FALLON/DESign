using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;

namespace DESign_WordAddIn
{
    class Print
    {
        public void PrintShopCopies()
        {
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;

            // Insert VBA code here. 

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            selection.Find.Execute("CHORD CUT SHEET");
            Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("chordCutSheet", selection.Range);

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            selection.EndKey(Word.WdUnits.wdStory, 1);
            Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("END", selection.Range);


            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            if (selection.Find.Execute("BASE PLATE CUT SHEET") == false)
            {
                selection.Find.Execute("WEB CUT SHEET");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                if (selection.Find.Execute("DETAIL FOLLOWS") == true)
                {
                    DialogResult dialogResult = MessageBox.Show("HAVE YOU ADDED ALL REQUIRED DETAILS TO THIS DOCUMENT?", "DETAIL CHECK", MessageBoxButtons.YesNo);
                    if (dialogResult != DialogResult.Yes)
                    {
                        goto ExitSub;
                    }
                }
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.Find.Execute("\f");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selection.GoTo(What: Word.WdGoToItem.wdGoToLine, Which: Word.WdGoToDirection.wdGoToNext, Count: 1);
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("BasePlate", selection.Range);

            }

            else
            {
                selection.Find.Execute("BASE PLATE CUT SHEET");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                if (selection.Find.Execute("DETAIL FOLLOWS") == true)
                {
                    DialogResult dialogResult = MessageBox.Show("HAVE YOU ADDED ALL REQUIRED DETAILS TO THIS DOCUMENT?", "DETAIL CHECK", MessageBoxButtons.YesNo);
                    if (dialogResult != DialogResult.Yes)
                    {
                        goto ExitSub;
                    }
                }
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.Find.Execute("\f");
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("BasePlate", selection.Range);
            }

            int intEnd;
            int intChordCutSheetPage;
            int intBasePlate;




            intEnd = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["END"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
            intChordCutSheetPage = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["chordCutSheet"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
            intBasePlate = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["BasePlate"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];

            string fistPage;
            string joistSheets;
            string shopCopies;
            string cutCopies;

            shopCopies = "1," + "2-" + Convert.ToString(intChordCutSheetPage - 1) + "," + Convert.ToString(intBasePlate) + "-" + Convert.ToString(intEnd) + "";
            cutCopies = String.Format("1, {0} - {1}", Convert.ToString(intChordCutSheetPage), Convert.ToString(intEnd));

            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            //Globals.ThisAddIn.Application.ActiveWindow.View.ShowRevisionsAndComments = true;
            //Globals.ThisAddIn.Application.ActiveWindow.View.ShowComments = false;
            //Globals.ThisAddIn.Application.ActiveWindow.View.ShowFormatChanges = false;
            Globals.ThisAddIn.Application.ActiveWindow.View.RevisionsMode = Word.WdRevisionsMode.wdInLineRevisions;




            //Globals.ThisAddIn.Application.ActivePrinter = "Xerox D110";
            Globals.ThisAddIn.Application.ActiveDocument.PrintOut(Range: Word.WdPrintOutRange.wdPrintRangeOfPages, Pages: shopCopies, Copies: 8);
            Globals.ThisAddIn.Application.ActiveDocument.PrintOut(Range: Word.WdPrintOutRange.wdPrintRangeOfPages, Pages: cutCopies, Copies: 6);

        ExitSub:
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);


        }
        public void SinglePrintShopCopies()
        {
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;

            // Insert VBA code here. 

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            selection.Find.Execute("CHORD CUT SHEET");
            Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("chordCutSheet", selection.Range);

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            selection.EndKey(Word.WdUnits.wdStory, 1);
            Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("END", selection.Range);


            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            if (selection.Find.Execute("BASE PLATE CUT SHEET") == false)
            {
                selection.Find.Execute("WEB CUT SHEET");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                if (selection.Find.Execute("DETAIL FOLLOWS") == true)
                {
                    DialogResult dialogResult = MessageBox.Show("HAVE YOU ADDED ALL REQUIRED DETAILS TO THIS DOCUMENT?", "DETAIL CHECK", MessageBoxButtons.YesNo);
                    if (dialogResult != DialogResult.Yes)
                    {
                        goto ExitSub;
                    }
                }
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.Find.Execute("\f");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selection.GoTo(What: Word.WdGoToItem.wdGoToLine, Which: Word.WdGoToDirection.wdGoToNext, Count: 1);
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("BasePlate", selection.Range);

            }

            else
            {
                selection.Find.Execute("BASE PLATE CUT SHEET");
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                if (selection.Find.Execute("DETAIL FOLLOWS") == true)
                {
                    DialogResult dialogResult = MessageBox.Show("HAVE YOU ADDED ALL REQUIRED DETAILS TO THIS DOCUMENT?", "DETAIL CHECK", MessageBoxButtons.YesNo);
                    if (dialogResult != DialogResult.Yes)
                    {
                        goto ExitSub;
                    }
                }
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.Find.Execute("\f");
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("BasePlate", selection.Range);
            }

            int intEnd;
            int intChordCutSheetPage;
            int intBasePlate;




            intEnd = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["END"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
            intChordCutSheetPage = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["chordCutSheet"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
            intBasePlate = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["BasePlate"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];

            string fistPage;
            string joistSheets;
            string shopCopies;
            string cutCopies;

            shopCopies = "1," + "2-" + Convert.ToString(intChordCutSheetPage - 1) + "," + Convert.ToString(intBasePlate) + "-" + Convert.ToString(intEnd) + "";
            cutCopies = String.Format("1, {0} - {1}", Convert.ToString(intChordCutSheetPage), Convert.ToString(intEnd));

            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            //Globals.ThisAddIn.Application.ActiveWindow.View.ShowRevisionsAndComments = true;
            //Globals.ThisAddIn.Application.ActiveWindow.View.ShowComments = false;
            //Globals.ThisAddIn.Application.ActiveWindow.View.ShowFormatChanges = false;
            Globals.ThisAddIn.Application.ActiveWindow.View.RevisionsMode = Word.WdRevisionsMode.wdInLineRevisions;




            //Globals.ThisAddIn.Application.ActivePrinter = "Xerox D110";
            Globals.ThisAddIn.Application.ActiveDocument.PrintOut(Range: Word.WdPrintOutRange.wdPrintRangeOfPages, Pages: shopCopies, Copies: 1);
            Globals.ThisAddIn.Application.ActiveDocument.PrintOut(Range: Word.WdPrintOutRange.wdPrintRangeOfPages, Pages: cutCopies, Copies: 1);

            ExitSub:
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);


        }
    }
}
