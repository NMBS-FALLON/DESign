using System;
using System.Collections.Generic;
using System.Text;

namespace DESign_BASE
{
    public class Load
    {
        enum Types {U, NU, GU, C, CB, CP, CA, C3, CZ, AX, M, S};
        public string Type
        {
            get { return Type; }
            set
            {
                if(Enum.GetNames(typeof(Types)).Contains(value) == true)
                {
                    Type = value;
                }
            }
        }
        enum Categories { TL, DL, LL, WL, SL, CL, SM, IP }
        public string Category
        {
            get { return Category; }
            set
            {
                if (Enum.GetNames(typeof(Categories)).Contains(value) == true)
                {
                    Category = value;
                }
            }
        }

        enum Positions { TC, BC, BE, LE, RE }
        public string Position
        {
            get { return Position; }
            set
            {
                if (Enum.GetNames(typeof(Positions)).Contains(value) == true)
                {
                    Position = value;
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
