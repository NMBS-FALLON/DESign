using System;
using System.Text.RegularExpressions;
using DESign_BASE;


namespace DESign_BASE
{
    public class JoistBase
    {
        
        public int Sequence { get; set; }
        public string Mark { get; set; }
        public int Quantity { get; set; }
        public string Description { get; set; }
        public double BaseLength { get; set; }
        public string JoistType { get; set; }
        public double SeatsBDL { get; set; }
        public double SeatsBDR { get; set; }
        public double TCXL { get; set; }
        public double TCXR { get; set; }
        public double BCXL { get; set; }
        public double BCXR { get; set; }
        public string TC { get; set; }
        public string BC { get; set; }
        public double MaterialCost { get; set; }
        public double WeightInLBS { get; set; }
        public double TotalLH { get; set; }
        public double BLDecimal { get; set; }
        public double Time { get; set; }
        public bool UseWood { get; set; }
        private int strippedNumber;
        public int StrippedNumber
        {
            get
            {
                strippedNumber = Convert.ToInt32(Regex.Replace(Mark, "[^0-9]", ""));
                return strippedNumber;
            }
        }
    }

    public class Joist : JoistBase
    {
       
        private string tcWidth;
        public string TCWidth
        {
            get
            {
                tcWidth = QueryAngleData.WNtcWidth(TC);
                return tcWidth;
            }
           
        }

    }

    public class Girder : JoistBase
    {
        private string tcWidth;
        public string TCWidth
        {
            get
            {
                tcWidth = QueryAngleData.TypTCWidth(TC);
                return tcWidth;
            }
        }

    }
}