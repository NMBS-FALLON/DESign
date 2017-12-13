using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DESign_WordAddIn
{
    class MSwordOperations
    {
        public void addSection(Word.WdOrientation orientation)
        {
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Section section = selection.Sections.Add();
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.PageSetup.Orientation = orientation;
            selection.PageSetup.LeftMargin = (float)50;
            selection.PageSetup.RightMargin = (float)50;


            section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;

            Word.Range currentLocation = Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0);
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
        }
    }
}
