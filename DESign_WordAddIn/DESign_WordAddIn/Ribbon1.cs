using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using DESign_WordAddIn.Insert_Blank_Sheets;
using System.Windows;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using Design.Docx_tools;


namespace DESign_WordAddIn
{
    public partial class RibbonNMBS
    {
        Print print = new Print();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void btnNailBacksheet_Click(object sender, RibbonControlEventArgs e)
        {
            
            new FormNailBacksheet().Show();            
        }

      
        private void btnHoldClear_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Range range = Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0);
            if (range.Find.Execute("COLOR CODE") == false)
            {
                MessageBox.Show("S.O. is missing a color code. Please fix and try again");
                return;
            }
            new FormHoldClear2().Show();
        }

        private void btnGetFiles_Click(object sender, RibbonControlEventArgs e)
        {
            new FormInsertJEDIImages().Show();         
        }

        private void btnExportToRisa_Click(object sender, RibbonControlEventArgs e)
        {
            
        }

        private void btnPrintJShopCopies_Click(object sender, RibbonControlEventArgs e)
        {

            print.PrintShopCopies(8, 6);
        }

        private void btnPrintGShopCopies_Click(object sender, RibbonControlEventArgs e)
        {

            print.PrintShopCopies(8, 4);
        }

        private void btnPrint1Master_Click(object sender, RibbonControlEventArgs e)
        {
            print.PrintShopCopies(1, 0);
        }

        private void btnBlankWorksheets_Click(object sender, RibbonControlEventArgs e)
        {
            new FormInsertBlankSheets().Show();
        }


        private void btnPrintJuarezShopCopies_Click(object sender, RibbonControlEventArgs e)
        {
            print.PrintShopCopiesJuarez();
        }

        private void btnPrint1Cut_Click(object sender, RibbonControlEventArgs e)
        {
            print.PrintShopCopies(0, 1);
        }

        private void btnV2HoldClear_Click(object sender, RibbonControlEventArgs e)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;
            var range = document.Range();
            var docxDoc = WordprocessingDocument.FromFlatOpcString(range.WordOpenXML);

            var summary = GetV2SoInfo.GetSoSummary(docxDoc);

            var summaryString = "";
            foreach (var item in summary)
            {
                summaryString += String.Format($"Mark: {item.Mark}, Qty: {item.Quantity}\r\n");
            }

            MessageBox.Show(summaryString);

        }
    }


}
