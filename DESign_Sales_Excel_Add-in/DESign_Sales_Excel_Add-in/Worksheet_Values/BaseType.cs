using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DESign_Sales_Excel_Add_in.Worksheet_Values
{
    public class BaseType
    {
        public StringWithUpdateCheck Name { get; set; }
        public StringWithUpdateCheck Description { get; set; }
        public DoubleWithUpdateCheck BaseLengthFt { get; set; }
        public DoubleWithUpdateCheck BaseLengthIn { get; set; }
        public IntWithUpdateCheck TcxlQuantity { get; set; }
        public DoubleWithUpdateCheck TcxlLengthFt { get; set; }
        public DoubleWithUpdateCheck TcxlLengthIn { get; set; }
        public IntWithUpdateCheck TcxrQuantity { get; set; }
        public DoubleWithUpdateCheck TcxrLengthFt { get; set; }
        public DoubleWithUpdateCheck TcxrLengthIn { get; set; }
        public DoubleWithUpdateCheck SeatDepthLE { get; set; }
        public DoubleWithUpdateCheck SeatDepthRE { get; set; }
        public IntWithUpdateCheck BcxQuantity { get; set; }
        public DoubleWithUpdateCheck Uplift { get; set; }
        public List<Load> Loads { get; set; }
        public StringWithUpdateCheck Erfos { get; set; }
        public DoubleWithUpdateCheck DeflectionTL { get; set; }
        public DoubleWithUpdateCheck DeflectionLL { get; set; }
        public StringWithUpdateCheck WnSpacing { get; set; }
        public List<StringWithUpdateCheck> Notes { get; set; }
    }


}
