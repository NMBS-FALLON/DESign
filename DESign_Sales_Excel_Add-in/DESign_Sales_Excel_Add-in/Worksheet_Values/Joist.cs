using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DESign_Sales_Excel_Add_in.Worksheet_Values
{
    [Serializable]
    public class Joist
    {
        public StringWithUpdateCheck Mark { get; set; }
        public List<StringWithUpdateCheck> BaseTypesOnMark { get; set; }
        public IntWithUpdateCheck Quantity { get; set; }
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
        private bool isGirder = false;
        public bool IsGirder
        {
            get
            {
                if (Description.Text.Contains("G") == true)
                {
                    isGirder = true;
                }
                return isGirder;
            }
        }
        private bool isLoadOverLoad = false;
        public bool IsLoadOverLoad
        {
            get
            {
                if (isGirder == false)
                {
                    string[] seperators = { "K", "LH", "DLH" };
                    string loadOrSeries = Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1];
                    if (loadOrSeries.Contains("/") == true)
                    {
                        isLoadOverLoad = true;
                    }
                }
                return isLoadOverLoad;
            }
        }
        private double tl;
        public double TL
        {
            get
            {
                if(IsLoadOverLoad == true)
                {
                    string[] seperators = { "K", "LH", "DLH" };

                    tl = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[0]);
                    
                }
                return tl;
            }
        }
        private double ll;
        public double LL
        {
            get
            {
                if(IsLoadOverLoad == true)
                {
                    string[] seperators = { "K", "LH", "DLH" };

                    ll = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[1]);

                }
                return ll;
            }
        }
        private double uDL;
        public double UDL
        {
            get
            {
                if (IsLoadOverLoad == true)
                {

                    uDL = TL - LL;
                }
                return uDL;
            }
        }
        

    }

}
