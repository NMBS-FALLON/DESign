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
        public void PrintShopCopies(int numShopCopies, int numCutCopies)
        {
            Word.Selection selection = Globals.ThisAddIn.Application.Selection; 

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

                /*
                if (selection.Find.Execute("\f") == true)
                {
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    selection.GoTo(What: Word.WdGoToItem.wdGoToLine, Which: Word.WdGoToDirection.wdGoToNext, Count: 1);
                    Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("BasePlate", selection.Range);
                }
                else
                {
                    MessageBox.Show("ERROR WITH PRINTING");
                    goto ExitSub;
                }
                */

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

                /*
                if (selection.Find.Execute("\f") == true)
                {
                    Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("BasePlate", selection.Range);
                }
                else
                {
                    MessageBox.Show("ERROR WITH PRINTING");
                    goto ExitSub;
                }
                */
            }

            int intEnd;
            int intChordCutSheetPage;
            int intBasePlate;




            intEnd = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["END"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
            intChordCutSheetPage = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["chordCutSheet"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
            intBasePlate = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["BasePlate"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];

            string shopCopies;
            string cutCopies;
            /*
            if (intChordCutSheetPage <= 1 && intBasePlate <= 1 && intEnd <= 1)
            {
                MessageBox.Show("ERROR WITH PRINTING");
                goto ExitSub;
            }
            */

            shopCopies = "1," + "2-" + Convert.ToString(intChordCutSheetPage - 1) + "," + Convert.ToString(intBasePlate) + "-" + Convert.ToString(intEnd) + "";
            cutCopies = String.Format("1, {0} - {1}", Convert.ToString(intChordCutSheetPage), Convert.ToString(intEnd));

            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            Globals.ThisAddIn.Application.ActiveWindow.View.RevisionsMode = Word.WdRevisionsMode.wdInLineRevisions;

            //Globals.ThisAddIn.Application.ActivePrinter = "Xerox D110";
            if (numShopCopies != 0)
            {
                Globals.ThisAddIn.Application.ActiveDocument.PrintOut(Range: Word.WdPrintOutRange.wdPrintRangeOfPages, Pages: shopCopies, Copies: numShopCopies);
            }
            if (numCutCopies != 0)
            {
                Globals.ThisAddIn.Application.ActiveDocument.PrintOut(Range: Word.WdPrintOutRange.wdPrintRangeOfPages, Pages: cutCopies, Copies: numCutCopies);
            }

        ExitSub:
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);

        }

        public void PrintShopCopiesJuarez()
        {
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.Find.Execute("\f");
            Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("endOfCoverSheet", selection.Range);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            selection.Find.Execute("CHORD CUT SHEET");
            Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("chordCutSheetStart", selection.Range);

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.Find.Execute("\f");
            Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("chordCutSheetEnd", selection.Range);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);



            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            if (selection.Find.Execute("END ROD CUT SHEET") == true)
            {
                selection.HomeKey(Word.WdUnits.wdStory, 0);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.Find.Execute("END ROD CUT SHEET");
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("endRodCutSheetStart", selection.Range);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.Find.Execute("\f");
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("endRodCutSheetEnd", selection.Range);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.HomeKey(Word.WdUnits.wdStory, 0);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            }

            if (selection.Find.Execute("BASE PLATE CUT SHEET") == false)
            {
                selection.Find.Execute("WEB CUT SHEET");
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("webCutSheetStart", selection.Range);
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
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("webCutSheetEnd", selection.Range);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                selection.GoTo(What: Word.WdGoToItem.wdGoToLine, Which: Word.WdGoToDirection.wdGoToNext, Count: 1);
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("BasePlateStart", selection.Range);

            }

            else
            {
                selection.Find.Execute("WEB CUT SHEET");
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("webCutSheetStart", selection.Range);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.Find.Execute("\f");
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("webCutSheetEnd", selection.Range);
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                selection.Find.Execute("BASE PLATE CUT SHEET");
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("basePlateCutSheetStart", selection.Range);
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


            }


            selection.EndKey(Word.WdUnits.wdStory, 1);
            Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add("END", selection.Range);

            int intEnd;
            int intChordCutSheetPage;
            int intBasePlate;




            intEnd = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["END"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
            intChordCutSheetPage = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["chordCutSheet"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];
            intBasePlate = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks["BasePlate"].Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber];

            string shopCopies;
            string cutCopies;
            /*
            if (intChordCutSheetPage <= 1 && intBasePlate <= 1 && intEnd <= 1)
            {
                MessageBox.Show("ERROR WITH PRINTING");
                goto ExitSub;
            }
            */

            shopCopies = "1," + "2-" + Convert.ToString(intChordCutSheetPage - 1) + "," + Convert.ToString(intBasePlate) + "-" + Convert.ToString(intEnd) + "";
            cutCopies = String.Format("1, {0} - {1}", Convert.ToString(intChordCutSheetPage), Convert.ToString(intEnd));

            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            Globals.ThisAddIn.Application.ActiveWindow.View.RevisionsMode = Word.WdRevisionsMode.wdInLineRevisions;

            //Globals.ThisAddIn.Application.ActivePrinter = "Xerox D110";
            
            // *** Globals.ThisAddIn.Application.ActiveDocument.PrintOut(Range: Word.WdPrintOutRange.wdPrintRangeOfPages, Pages: shopCopies, Copies: numShopCopies);
           // *** Globals.ThisAddIn.Application.ActiveDocument.PrintOut(Range: Word.WdPrintOutRange.wdPrintRangeOfPages, Pages: cutCopies, Copies: numCutCopies);

        ExitSub:
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);

        }
    }
}
