using System;
using System.Text.RegularExpressions;
using DESign_BASE_WPF;
using System.Collections.Generic;


namespace DESign_BASE_WPF
{
    public class JoistBase
    {
        
        public int Sequence { get; set; }
        public string Mark { get; set; }
        public int Quantity { get; set; }
        public string Description { get; set; }
        public double dblBaseLength { get; set; }
        public string strBaseLength { get; set; }
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
        public List<string> Notes { get; set; }
        public List<string> Loads { get; set; }
    }

    public class Joist : JoistBase
    {
        QueryAngleData queryAngleData = new QueryAngleData();

       
        private string tcWidth;
        public string TCWidth
        {
            get
            {
                tcWidth = queryAngleData.WNtcWidth(TC);
                return tcWidth;
            }
           
        }

    }

    public class Girder : JoistBase
    {
        QueryAngleData queryAngleData = new QueryAngleData();
        private string tcWidth;
        public string TCWidth
        {
            get
            {
                tcWidth = queryAngleData.TypTCWidth(TC);
                return tcWidth;
            }
        }

    }
}