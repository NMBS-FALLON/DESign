using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using DESign_WordAddIn;

namespace DESign_WordAddIn.Insert_Blank_Sheets
{
    public partial class FormInsertBlankSheets : Form
    {
        public FormInsertBlankSheets()
        {
            InitializeComponent();
        }

        private void btnInsertBirdcage_Click(object sender, EventArgs e)
        {
            addDetail(Properties.Resources.blankBirdCage, "Bird Cage Detail");

        }

        private void btnInsertBlankTPlate_Click(object sender, EventArgs e)
        {
            addDetail(Properties.Resources.blankTPlate, "T-Plate Detail");
        }

        private void btnInsertBlankShimPlate_Click(object sender, EventArgs e)
        {
            addShim(Properties.Resources.blankShimPlate);
            GC.Collect();
        }






        private void addDetail(Bitmap detail, string detailType)
        {
            Bitmap birdCage = detail;

            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Section section = selection.Sections.Add();
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
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

            Word.Table tblDetailHeader = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 2, 6);


            tblDetailHeader.Cell(1, 1).Range.Text = "JOB NAME: ";
            tblDetailHeader.Cell(2, 1).Range.Text = "LOCATION: ";
            tblDetailHeader.Cell(1, 3).Range.Text = "JOB #: ";
            tblDetailHeader.Cell(2, 3).Range.Text = "LIST:  ";
            tblDetailHeader.Cell(1, 5).Range.Text = "SHEET #:";
            tblDetailHeader.Cell(1, 6).Range.Text = detailType;
            tblDetailHeader.Rows.SetHeight(10f, Word.WdRowHeightRule.wdRowHeightExactly);

            for (int i = 1; i <= 2; i++)
            {
                for (int col = 1; col <= 5; col = col + 2)
                {
                    tblDetailHeader.Cell(i, col).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    tblDetailHeader.Cell(i, col).Range.Underline = Word.WdUnderline.wdUnderlineSingle;
                    tblDetailHeader.Cell(i, col).Range.Font.Bold = 1;
                }

            }
            try
            {
                JoistCoverSheet joistCoverSheet = new JoistCoverSheet();
                List<List<string>> joistData = joistCoverSheet.JoistData();

                tblDetailHeader.Cell(1, 2).Range.Text = joistData[8][0];
                tblDetailHeader.Cell(2, 2).Range.Text = joistData[9][0];
                tblDetailHeader.Cell(1, 4).Range.Text = joistData[7][0];
                tblDetailHeader.Cell(2, 4).Range.Text = joistData[10][0];
            }
            catch { }


            for (int i = 1; i <= 2; i++)
            {
                for (int col = 2; col <= 4; col = col + 2)
                {
                    tblDetailHeader.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
            }




            selection.EndKey(Word.WdUnits.wdStory, 1);

            tblDetailHeader.Columns[1].Width = 65;
            tblDetailHeader.Columns[2].Width = 200;
            tblDetailHeader.Columns[3].Width = 50;
            tblDetailHeader.Columns[4].Width = 200;
            tblDetailHeader.Columns[5].Width = 60;


            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.EndKey(Word.WdUnits.wdStory);
            selection.Text = "\r\n";
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            IDataObject idat = null;
            Exception threadEx = null;
            Thread staThread = new Thread(
                delegate ()
                {
                    try
                    {

                        string sk1FileName = System.IO.Path.GetTempFileName();
                        Byte[] sk1ByteArray = ImageToByte(birdCage);
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

            Word.Table tblDetailInfo = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 6, 16);
            tblDetailInfo.AllowAutoFit = true;
            tblDetailInfo.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);


            //tblDetailInfo.Columns[3].Width = 120;



            tblDetailInfo.Borders.Enable = 1;

            for (int col = 1; col <= tblDetailInfo.Columns.Count; col++)
            {
                tblDetailInfo.Cell(1, col).Borders.Enable = 1;
            }

            tblDetailInfo.Cell(1, 1).Range.Text = "Qty";
            tblDetailInfo.Cell(1, 2).Range.Text = "Desc";
            tblDetailInfo.Cell(1, 3).Range.Text = "Mark(s)";
            tblDetailInfo.Cell(1, 4).Range.Text = "A";
            tblDetailInfo.Cell(1, 5).Range.Text = "B";
            tblDetailInfo.Cell(1, 6).Range.Text = "C";
            tblDetailInfo.Cell(1, 7).Range.Text = "D";
            tblDetailInfo.Cell(1, 8).Range.Text = "E";
            tblDetailInfo.Cell(1, 9).Range.Text = "F";
            tblDetailInfo.Cell(1, 10).Range.Text = "G";
            tblDetailInfo.Cell(1, 11).Range.Text = "H";
            tblDetailInfo.Cell(1, 12).Range.Text = "I";
            tblDetailInfo.Cell(1, 13).Range.Text = "J";
            tblDetailInfo.Cell(1, 14).Range.Text = "K";
            tblDetailInfo.Cell(1, 15).Range.Text = "L";
            tblDetailInfo.Cell(1, 16).Range.Text = "S";
            tblDetailInfo.Rows.SetHeight(10f, Word.WdRowHeightRule.wdRowHeightExactly);

            for (int col = 1; col <= 16; col++)
            {
                tblDetailInfo.Cell(1, col).Range.Font.Size = 8;
                tblDetailInfo.Cell(1, col).Range.Font.Bold = 1;
                tblDetailInfo.Cell(1, col).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            }

            tblDetailInfo.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            this.Close();
        }

        private void addShim(Bitmap detail)
        {
            Bitmap birdCage = detail;

            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Section section = selection.Sections.Add();
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
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

            Word.Table tblDetailHeader = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 2, 6);

            tblDetailHeader.Cell(1, 1).Range.Text = "JOB NAME: ";
            tblDetailHeader.Cell(2, 1).Range.Text = "LOCATION: ";
            tblDetailHeader.Cell(1, 3).Range.Text = "JOB #: ";
            tblDetailHeader.Cell(2, 3).Range.Text = "LIST:  ";
            tblDetailHeader.Cell(1, 5).Range.Text = "SHEET #:";
            tblDetailHeader.Cell(1, 6).Range.Text = "Shim Detail";
            tblDetailHeader.Rows.SetHeight(10f, Word.WdRowHeightRule.wdRowHeightExactly);

            for (int i = 1; i <= 2; i++)
            {
                for (int col = 1; col <= 5; col = col + 2)
                {
                    tblDetailHeader.Cell(i, col).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    tblDetailHeader.Cell(i, col).Range.Underline = Word.WdUnderline.wdUnderlineSingle;
                    tblDetailHeader.Cell(i, col).Range.Font.Bold = 1;
                }

            }
            try
            {
                JoistCoverSheet joistCoverSheet = new JoistCoverSheet();
                List<List<string>> joistData = joistCoverSheet.JoistData();

                tblDetailHeader.Cell(1, 2).Range.Text = joistData[8][0];
                tblDetailHeader.Cell(2, 2).Range.Text = joistData[9][0];
                tblDetailHeader.Cell(1, 4).Range.Text = joistData[7][0];
                tblDetailHeader.Cell(2, 4).Range.Text = joistData[10][0];
            }
            catch { }


            for (int i = 1; i <= 2; i++)
            {
                for (int col = 2; col <= 4; col = col + 2)
                {
                    tblDetailHeader.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
            }




            selection.EndKey(Word.WdUnits.wdStory, 1);

            tblDetailHeader.Columns[1].Width = 65;
            tblDetailHeader.Columns[2].Width = 200;
            tblDetailHeader.Columns[3].Width = 50;
            tblDetailHeader.Columns[4].Width = 200;
            tblDetailHeader.Columns[5].Width = 60;


            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.EndKey(Word.WdUnits.wdStory);
            selection.Text = "\r\n";
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            IDataObject idat = null;
            Exception threadEx = null;
            Thread staThread = new Thread(
                delegate ()
                {
                    try
                    {

                        string sk1FileName = System.IO.Path.GetTempFileName();
                        Byte[] sk1ByteArray = ImageToByte(birdCage);
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

            Word.Table tblDetailInfo = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(selection.Range, 6, 18);
            tblDetailInfo.AllowAutoFit = true;
            tblDetailInfo.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);


            //tblDetailInfo.Columns[3].Width = 120;



            tblDetailInfo.Borders.Enable = 1;

            for (int col = 1; col <= tblDetailInfo.Columns.Count; col++)
            {
                tblDetailInfo.Cell(1, col).Borders.Enable = 1;
            }

            tblDetailInfo.Cell(1, 1).Range.Text = "Qty";
            tblDetailInfo.Cell(1, 2).Range.Text = "Desc";
            tblDetailInfo.Cell(1, 3).Range.Text = "Mark(s)";
            tblDetailInfo.Cell(1, 4).Range.Text = "A";
            tblDetailInfo.Cell(1, 5).Range.Text = "B";
            tblDetailInfo.Cell(1, 6).Range.Text = "C";
            tblDetailInfo.Cell(1, 7).Range.Text = "D";
            tblDetailInfo.Cell(1, 8).Range.Text = "E";
            tblDetailInfo.Cell(1, 9).Range.Text = "F";
            tblDetailInfo.Cell(1, 10).Range.Text = "G";
            tblDetailInfo.Cell(1, 11).Range.Text = "H";
            tblDetailInfo.Cell(1, 12).Range.Text = "T1";
            tblDetailInfo.Cell(1, 13).Range.Text = "Cut";
            tblDetailInfo.Cell(1, 14).Range.Text = "T2";
            tblDetailInfo.Cell(1, 15).Range.Text = "T3";
            tblDetailInfo.Cell(1, 16).Range.Text = "K";
            tblDetailInfo.Cell(1, 17).Range.Text = "L";
            tblDetailInfo.Cell(1, 18).Range.Text = "S";
            tblDetailInfo.Rows.SetHeight(10f, Word.WdRowHeightRule.wdRowHeightExactly);

            for (int col = 1; col <= 18; col++)
            {
                tblDetailInfo.Cell(1, col).Range.Font.Size = 8;
                tblDetailInfo.Cell(1, col).Range.Font.Bold = 1;
                tblDetailInfo.Cell(1, col).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            }

            tblDetailInfo.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.EndKey(Word.WdUnits.wdStory, 1);
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            this.Close();
        }

        public static byte[] ImageToByte(Image img)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }

    }

}