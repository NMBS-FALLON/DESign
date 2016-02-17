using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

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

        private void btnPrintShopCopies_Click(object sender, RibbonControlEventArgs e)
        {
            print.PrintShopCopies();
        }



    }


}
