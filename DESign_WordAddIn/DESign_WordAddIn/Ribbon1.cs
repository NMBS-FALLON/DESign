using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using DESign_WordAddIn.Insert_Blank_Sheets;
using System.Windows;

namespace DESign_WordAddIn
{
    public partial class RibbonNMBS
    {
        SOtoRISA soToRISA = new SOtoRISA();
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
            new FormSOtoRISA().Show();
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
    }


}
