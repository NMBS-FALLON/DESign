using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DESign_Sales_Excel_Add_in.Deck
{
    class DiaphragmTable
    {
        public string DeckType { get; set; }
        public int Gauge { get; set; }
        public double Ksi { get; set; }
        public string SupportFastener { get; set; }
        public string SidelapFastener { get; set; }
        public class Table
        {
            object[] table { get; set; }
        }
    }
}
