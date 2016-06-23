using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace DESign_BASE_WPF
{
    
    public class Load
    {
        StringManipulation stringManipulation = new StringManipulation();

        enum Types {U, NU, GU, C, CB, CP, CA, C3, CZ, AX, M, S};
        private string type = "";
        public string Type
        {
            get { return type; }
            set
            {
                if (Enum.GetNames(typeof(Types)).Contains(value) == true)
                {
                    type = value;
                }
            }
        }
        enum Categories { TL, DL, LL, WL, SL, CL, SM, IP }
        private string category = "";
        public string Category
        {
            get { return category; }
            set
            {
                if (Enum.GetNames(typeof(Categories)).Contains(value) == true)
                {
                    category = value;
                }
            }
        }

        enum Positions { TC, BC, BE, LE, RE }
        private string position = "";
        public string Position
        {
            get { return position; }
            set
            {
                if (Enum.GetNames(typeof(Positions)).Contains(value) == true)
                {
                    position = value;
                }
            }
        }

        public string Reference { get; set; }
        public int Group { get; set; }
        public double StartValue { get; set; }
        private double dblStartLocation = 0.0;
        public double DblStartLocation
        {
            get
            {
                if (StrStartLocation != "")
                {
                    dblStartLocation = stringManipulation.hyphenLengthToDecimal(StrStartLocation);
                }
                return dblStartLocation;
            }
            set
            {
                dblStartLocation = value;
            }
        }
        public string StrStartLocation { get; set; }
        public double EndValue { get; set; }
        private double dblEndLocation = 0.0;
        public double DblEndLocation
        {
            get
            {
                if (StrEndLocation != "")
                {
                    dblEndLocation = stringManipulation.hyphenLengthToDecimal(StrEndLocation);
                }
                return dblEndLocation;
            }
            set
            {
                dblEndLocation = value;
            }
        }
        public string StrEndLocation { get; set; }
    }

    
}
