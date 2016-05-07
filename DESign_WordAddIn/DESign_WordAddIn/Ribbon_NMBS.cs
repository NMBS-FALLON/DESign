using Microsoft.Office.Tools.Ribbon;

namespace DESign_WordAddIn
{
    public partial class Ribbon_NMBS
    {
        private void Ribbon_NMBS_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnNailerBackSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormNailBacksheet form1 = new FormNailBacksheet();
            form1.Show();
        }

        private void btnHoldClear_Click(object sender, RibbonControlEventArgs e)
        {
            FormHoldClear2 formhc2 = new FormHoldClear2();
            formhc2.Show();
        }

        private void btnPrintShopCopies_Click(object sender, RibbonControlEventArgs e)
        {
            Print print1 = new Print();
            print1.PrintShopCopies();
        }
    }
}
