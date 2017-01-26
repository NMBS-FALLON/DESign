using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DESign_Sales_Excel_Add_in.Worksheet_Values
{
    public class UpdateCheck
    {
        public bool IsUpdated { get; set; } = false;
    }
    public class DoubleWithUpdateCheck : UpdateCheck
    {
        public double? Value { get; set; }

    }
    public class StringWithUpdateCheck : UpdateCheck
    {
        public string Text { get; set; }

    }
    public class IntWithUpdateCheck : UpdateCheck
    {
        public int? Value { get; set; }

    }
}
