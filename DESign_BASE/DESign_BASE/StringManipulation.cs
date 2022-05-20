using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;




namespace DESign_BASE
{
    public class StringManipulation
    {
        public static double hyphenLengthToDecimal(string hyphenLength)
        {

            bool containsHyphen = hyphenLength.Contains("-");
            bool containsDivide = hyphenLength.Contains("/");

            double LengthinFt = 0;
            if (containsHyphen == true && containsDivide == true)
            {
                Char[] delimChars = { ' ', '-', '/' };
                string[] hyphenLengthArray = hyphenLength.Split(delimChars, StringSplitOptions.RemoveEmptyEntries);

                double[] hyphenLengthArrayDouble = new Double[hyphenLengthArray.Length];
                for (int i = 0; i < hyphenLengthArray.Length; i++)
                {
                    hyphenLengthArrayDouble[i] = Convert.ToDouble(hyphenLengthArray[i]);
                }
                LengthinFt = hyphenLengthArrayDouble[0] + ((hyphenLengthArrayDouble[1] + (hyphenLengthArrayDouble[2] / hyphenLengthArrayDouble[3])) / 12);
            }
            else if (containsHyphen == true && containsDivide == false)
            {
                Char[] delimChars = { ' ', '-', '/' };
                string[] hyphenLengthArray = hyphenLength.Split(delimChars, StringSplitOptions.RemoveEmptyEntries);

                double[] hyphenLengthArrayDouble = new Double[hyphenLengthArray.Length];
                for (int i = 0; i < hyphenLengthArray.Length; i++)
                {
                    hyphenLengthArrayDouble[i] = Convert.ToDouble(hyphenLengthArray[i]);
                }
                LengthinFt = hyphenLengthArrayDouble[0] + hyphenLengthArrayDouble[1] / 12;
            }
            else if (containsHyphen == false && containsDivide == true)
            {
                Char[] delimChars = { ' ', '-', '/' };
                string[] hyphenLengthArray = hyphenLength.Split(delimChars, StringSplitOptions.RemoveEmptyEntries);

                double[] hyphenLengthArrayDouble = new Double[hyphenLengthArray.Length];
                for (int i = 0; i < hyphenLengthArray.Length; i++)
                {
                    hyphenLengthArrayDouble[i] = Convert.ToDouble(hyphenLengthArray[i]);
                }
                LengthinFt = (hyphenLengthArrayDouble[0] + (hyphenLengthArrayDouble[1] / hyphenLengthArrayDouble[2])) / 12;
            }
            else if (containsHyphen == false && containsDivide == false)
            {
                Char[] delimChars = { ' ', '-', '/' };
                string[] hyphenLengthArray = hyphenLength.Split(delimChars, StringSplitOptions.RemoveEmptyEntries);

                double[] hyphenLengthArrayDouble = new Double[hyphenLengthArray.Length];
                for (int i = 0; i < hyphenLengthArray.Length; i++)
                {
                    hyphenLengthArrayDouble[i] = Convert.ToDouble(hyphenLengthArray[i]);
                }
                LengthinFt = hyphenLengthArrayDouble[0] / 12;
            }


            return LengthinFt;


        }
        public static string DecimilLengthToHyphen(double decimilLength)
        //9.84375 = 9-10 1/8
        {
            double feetDouble = Math.Floor(decimilLength);
            int feetInt = Convert.ToInt32(feetDouble);
            //9


            double inches = (decimilLength - feetDouble) * 12;
            //10.125

            double inchDouble = Math.Floor(inches);
            int inchInt = Convert.ToInt32(inchDouble);
            //10

            double fraction = (inches - inchDouble) * 8;
            //1 (for 1/8)

            double fractionRound = Math.Round(fraction, 0, MidpointRounding.AwayFromZero);

            int fractionInt = Convert.ToInt32(fractionRound);

            bool nullFraction = false;

            if (fractionInt == 0)
            {
                nullFraction = true;
            }
            if (fractionInt == 8)
            {
                inchInt = inchInt + 1;
                fractionInt = 0;
                nullFraction = true;

                if (inchInt == 12)
                {
                    feetInt = feetInt + 1;
                    inchInt = 0;
                }
            }

            string feetString = Convert.ToString(feetInt);
            string inchString = Convert.ToString(inchInt);
            string fractionString = Convert.ToString(fractionInt);

            string divisor = "8";

            if (fractionInt == 2) { fractionString = "1"; divisor = "4"; }
            if (fractionInt == 4) { fractionString = "1"; divisor = "2"; }
            if (fractionInt == 6) { fractionString = "3"; divisor = "4"; }

            string HyphenLength = feetString + "-" + inchString + " " + fractionString + "/" + divisor;


            if (nullFraction == true)
            {
                HyphenLength = feetString + "-" + inchString;
            }

            return HyphenLength;
        }

        public static string convertLengthStringtoHyphenLength(string lengthString)
        {
            bool containsHyphen = lengthString.Contains("-");
            bool containsBackSlash = lengthString.Contains("/");
            bool containsSpace = lengthString.Contains(" ");
            bool containsPeriod = lengthString.Contains(".");

            if (containsHyphen == true) { }
            else if (containsHyphen == false && containsSpace == true && containsBackSlash == true) { lengthString = "0-" + lengthString; }
            else if (containsHyphen == false && containsBackSlash == true) { lengthString = "0-0 " + lengthString; }
            else if (containsHyphen == false && containsPeriod == true) { lengthString = "0-" + lengthString; }
            else if (containsHyphen == false && containsPeriod == false && containsBackSlash == false) { lengthString = "0-" + lengthString; }
            else if (lengthString == "") { lengthString = "0-0"; }
            else { }

            return lengthString;

        }

        public static List<double> doubleListwithTolerance(List<double> doubleList, double tolerance)
        {

            int maxIndex = doubleList.Count - 1;
            double[] decendingDoubleList = doubleList.ToArray();
            Array.Sort(decendingDoubleList);

            for (int i = 0; i <= maxIndex; i++)
            {
                for (int k = i; k <= maxIndex; k++)

                    if (decendingDoubleList[k] - decendingDoubleList[i] <= tolerance)
                    {
                        decendingDoubleList[k] = decendingDoubleList[i];
                    }

            }


            double[] consolidatedDoubleArray = new double[doubleList.Count];

            for (int i = 0; i <= doubleList.Count - 1; i++)
            {

                int k = 0;
                double difference = doubleList[i] - decendingDoubleList[k];
                if (difference <= tolerance)
                {
                    consolidatedDoubleArray[i] = decendingDoubleList[k];
                }
                else
                {
                    bool withinTolerance = false;

                    while (withinTolerance == false)
                    {
                        k = k + 1;
                        difference = doubleList[i] - decendingDoubleList[k];
                        if (difference <= tolerance)
                        {
                            consolidatedDoubleArray[i] = decendingDoubleList[k];
                            withinTolerance = true;
                        }
                    }
                }


            }

            List<double> doubleListwithTolerance = new List<double>();
            for (int i = 0; i <= doubleList.Count - 1; i++)
            {
                doubleListwithTolerance.Add(consolidatedDoubleArray[i]);
            }

            return doubleListwithTolerance;
        }

        public static bool areStringElementsEqual(List<string> List)
        {
            bool has1Element = false;


            List<string> distincList = new List<string>(1);


            distincList.Add(List[0]);




            for (int i = 1; i <= List.Count - 1; i++)  //List.Count-2 ??
            {
                if (List[i] != List[0])
                {
                    distincList.Add(List[i]);
                }

            }

            if (distincList.Count == 1) { has1Element = true; }

            return has1Element;
        }

        public static double ConvertLengthtoDecimal(string length)
        {
            string lengthString = convertLengthStringtoHyphenLength(length);
            double dblLength = hyphenLengthToDecimal(lengthString);

            return dblLength;
        }
    }
}
