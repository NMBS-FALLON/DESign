using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DESign_Sales_Excel_Add_in.Worksheet_Values
{
    public class BaseTypes_WS_Values
    {
        public string Mark { get; set; }
        public string Description { get; set; }
        public double BaseLengthFt { get; set; }
        public double BaseLengthIn { get; set; }
        public int TcxlQuantity { get; set; }
        public double TcxlLengthFt { get; set; }
        public double TcxlLengthIn { get; set; }
        public int TcxrQuantity { get; set; }
        public double TcxrLengthFt { get; set; }
        public double TcxrLengthIn { get; set; }
        public double SeatDepthLE { get; set; }
        public double SeatDepthRE { get; set; }
        public int BcxQuantity { get; set; }
        public double Uplift { get; set; }
        public List<Load> Loads { get; set; }
        public string Erfos { get; set; }
        public string DeflectionTL { get; set; }
        public string DeflectionLL { get; set; }
        public string WnSpacing { get; set; }
        public List<string> Remarks { get; set; }
    }

}
