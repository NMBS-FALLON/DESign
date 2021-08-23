using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DESign_Excel_Add_In
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var materials = ExtractFromSql.MarksInShoporder("Juarez", "0110-0043", "K1-1");
            
            foreach( var material in materials)
            {
                MessageBox.Show(String.Format("Mark = {0}", material));
            }
        }
    }
}
