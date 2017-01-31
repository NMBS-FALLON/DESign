using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DESign_Sales_Excel_Add_in.Worksheet_Values
{
    [Serializable]
    public class UpdateCheck
    {
        public bool IsUpdated { get; set; } = false;
    }

    [Serializable]
    public class DoubleWithUpdateCheck : UpdateCheck
    {
        public double? Value { get; set; }

    }
    [Serializable]
    public class StringWithUpdateCheck : UpdateCheck
    {
        public string Text { get; set; }

    }
    [Serializable]
    public class IntWithUpdateCheck : UpdateCheck
    {
        public int? Value { get; set; }

    }
}
