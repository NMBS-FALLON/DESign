using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using DESign_WordAddIn;


namespace DESign_WordAddIn
{
    
    class JoistCoverSheet
    {
        public List<List<string>> JoistData()
        {

            List<string> joistMarks = new List<string>();
            List<string> joistQuantities = new List<string>();
            List<string> joistLengths = new List<string>();
            List<string> joistDepth = new List<string>();
            List<string> joistTC = new List<string>();
            List<string> joistBC = new List<string>();

            Word.Selection selection = Globals.ThisAddIn.Application.Selection;

             selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.Find.Execute("MARK ");

            selection.MoveDown(Word.WdUnits.wdLine, 2);
            
            string[] joistDataArray = null;
            do
            {
            selection.HomeKey(Word.WdUnits.wdLine);
            selection.EndKey(Extend: Word.WdUnits.wdLine);
            selection.Copy();
            string joistData = selection.Text;
            Char[] delimChars = {' ', '\u00A0'};
            joistDataArray = joistData.Split(delimChars, StringSplitOptions.RemoveEmptyEntries);

            if (joistDataArray.Length > 7)
            {
                joistMarks.Add(joistDataArray[0]);
                joistQuantities.Add(joistDataArray[1]);

                bool lengthWithFraction = joistDataArray[4].Contains('/');
                int subtractIndex = 0;
                if (lengthWithFraction == true)
                {
                    string joistLength = joistDataArray[3] + " " + joistDataArray[4];
                    joistLengths.Add(joistLength);
                }
                else if (lengthWithFraction == false)
                {
                    joistLengths.Add(joistDataArray[3]);
                    subtractIndex = -1;
                }

                joistDepth.Add(joistDataArray[8 + subtractIndex]);

                string TCBC = joistDataArray[10 + subtractIndex];
                Char[] TCBCdelimChars = { '/' };
                string[] arrayTCBC = TCBC.Split(TCBCdelimChars, StringSplitOptions.RemoveEmptyEntries);

                joistTC.Add(arrayTCBC[0]);
                joistBC.Add(arrayTCBC[1]);
            }
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.Find.Execute("\u00A0");

            } while (joistDataArray.Length >= 7);


            


       
            //TOTAL QUANTITY OF JOISTS

            int intTotalJoistQuantity = 0;

    
            for (int i=0; i<=joistQuantities.Count-1; i++)
            {
                int joistQuantity = Convert.ToInt32(joistQuantities[i]);
                intTotalJoistQuantity = intTotalJoistQuantity + joistQuantity;
            }

            List<string> ListTotalJoistQuantity = new List<string>();

            ListTotalJoistQuantity.Add(Convert.ToString(intTotalJoistQuantity));

            //

            //GATHERING OTHER INFORMATION
            selection.HomeKey(Word.WdUnits.wdStory, 0);
            selection.MoveDown(Word.WdUnits.wdLine, 2);
            selection.HomeKey(Word.WdUnits.wdLine, 1);
            selection.EndKey(Extend: Word.WdUnits.wdLine);
            selection.Copy();
            string jobTitle = selection.Text;
            string[] jobTitleArray = jobTitle.Split(new string[] { "  ", "\u00A0" }, StringSplitOptions.RemoveEmptyEntries);

            List<string> jobNumber = new List<string>();
            List<string> jobName = new List<string>();
            List<string> jobLocation = new List<string>();
            if (jobTitleArray.Length == 4)
            {
                jobNumber.Add(jobTitleArray[0]);
                jobName.Add(jobTitleArray[1]);
                jobLocation.Add(jobTitleArray[2]);
            }
            else
            {
                jobNumber.Add(jobTitleArray[0]);
                jobName.Add(jobTitleArray[1].Substring(0, 30));
                jobLocation.Add(" ");
            }

            
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.Find.Execute("LIST"); 
            selection.HomeKey(Word.WdUnits.wdLine, 1);
            selection.EndKey(Extend: Word.WdUnits.wdLine);
            selection.Copy();
            string listNumberLine = selection.Text;
            string[] ListNumberLineArray = listNumberLine.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            List<string> listNumber = new List<string>();
            listNumber.Add(ListNumberLineArray[2]);

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);





            //END OF GATHERING OTHER INFORMATION
            
            
            List<List<string>> JoistData = new List<List<string>>();

            JoistData.Add(joistMarks);
            JoistData.Add(joistQuantities);
            JoistData.Add(joistLengths);
            JoistData.Add(joistDepth);
            JoistData.Add(joistTC);
            JoistData.Add(joistBC);
            JoistData.Add(ListTotalJoistQuantity);
            JoistData.Add(jobNumber);
            JoistData.Add(jobName);
            JoistData.Add(jobLocation);
            JoistData.Add(listNumber);


            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdStory, 0);

            return JoistData;


        }
    }
}
