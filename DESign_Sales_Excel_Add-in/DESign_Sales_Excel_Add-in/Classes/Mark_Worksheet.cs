using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DESign_Sales_Excel_Add_in.Classes
{
    public class MarksWsValues
    {
        public string Mark { get; set; }
        public string BaseTypes { get; set; }
        public int Quantity { get; set; }
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
        public string LoadInfoType { get; set; }
        public string LoadInfoCategory { get; set; }
        public string LoadInfoPosition { get; set; }
        public double Load1Value { get; set; }
        public double Load1DistanceFt { get; set; }
        public double Load1DistanceIn { get; set; }
        public double Load2Value { get; set; }
        public double Load2DistanceFt { get; set; }
        public double Load2DistanceIn { get; set; }
        public string CaseNumber { get; set; }
        public string Erfos { get; set; }
        public string DeflectionTL { get; set; }
        public string DeflectionLL { get; set; }
        public string WnSpacing { get; set; }
        public string Remarks { get; set; }
    }

}
