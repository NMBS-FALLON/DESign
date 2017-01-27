using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DESign_Sales_Excel_Add_in.Worksheet_Values
{
    public class Bridging
    {
        public StringWithUpdateCheck Sequence { get; set; }
        public StringWithUpdateCheck Size { get; set; }
        public StringWithUpdateCheck HorX { get; set; }
        public IntWithUpdateCheck Rows { get; set; }
        public DoubleWithUpdateCheck Length { get; set; }
        public DoubleWithUpdateCheck TotalLength { get; set; }
        public List<StringWithUpdateCheck> Notes { get; set; }
    }
}
