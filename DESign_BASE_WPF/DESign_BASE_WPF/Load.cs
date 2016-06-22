using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace DESign_BASE_WPF
{
    public class Load
    {
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
        public double StartLocation { get; set; }
        public double EndValue { get; set; }
        public double EndLocation { get; set; }
    }

    
}
