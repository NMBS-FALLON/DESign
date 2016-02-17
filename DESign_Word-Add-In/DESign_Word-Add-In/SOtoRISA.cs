using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;

namespace DESign_WordAddIn
{
   
    class SOtoRISA
    {
        JoistCoverSheet joistCoverSheet = new JoistCoverSheet();
        StringManipulation stringManipulation = new StringManipulation();

        public void getNodes(string joistMark)
        {
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            //   List<List<string>> soJoists = joistCoverSheet.JoistData(); 
            // foreach (string joistMark in soJoists[0])
            //  {
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.Move(Word.WdUnits.wdStory, 0);
            selection.Find.Execute(joistMark + "             ");
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            selection.Find.Execute("DEPTH");
            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            selection.HomeKey(Word.WdUnits.wdLine);
            selection.MoveDown(Word.WdUnits.wdLine, 1);
            selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdExtend);
            string depthString = selection.Text;
            string[] depthStringArray = depthString.Split(new string[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
            double depthLE = stringManipulation.hyphenLengthToDecimal(depthStringArray[1]);
            double depthCenter = stringManipulation.hyphenLengthToDecimal(depthStringArray[2]);
            double depthRE = stringManipulation.hyphenLengthToDecimal(depthStringArray[3]);

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            selection.Find.Execute("WELD");
            selection.HomeKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);
            selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdExtend);
            string headerString = selection.Text;
            int charsToSectionStart = headerString.IndexOf("SECTION");
            int CharsInSection = headerString.IndexOf("DESC") - charsToSectionStart;
            int charsToDESCStart = headerString.IndexOf("DESC");
            int charsInDESC = headerString.IndexOf("USE") - charsToDESCStart;
            int charsToTCStart = headerString.IndexOf("TC");
            int charsInTC = headerString.IndexOf("BC") - charsToTCStart;
            int charsToBCStart = headerString.IndexOf("BC");

            selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            selection.MoveDown(Word.WdUnits.wdLine, 1);
            string lineText = "";
            List<string> combinedStrings = new List<string>();
            List<JoistMember> joistMembers = new List<JoistMember>();
            do
            {
                PointF tcNode = new PointF();
                PointF bcNode = new PointF();
                selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdExtend);
                lineText = selection.Text;
                int charsInBC = lineText.Length - charsToBCStart;
                string material = lineText.Substring(charsToSectionStart, CharsInSection).Trim();
                string member = lineText.Substring(charsToDESCStart, charsInDESC).Trim();
                double tcX = stringManipulation.hyphenLengthToDecimal(lineText.Substring(charsToTCStart, charsInTC).Trim());
                double bcX = stringManipulation.hyphenLengthToDecimal(lineText.Substring(charsToBCStart, charsInBC - 1).Trim());

                if(depthLE==depthCenter && depthLE == depthRE)
                {
                    tcNode.X = Convert.ToSingle(tcX)*76;
                    tcNode.Y = 0.0F; //subtract chord centroidal axis

                    bcNode.X = Convert.ToSingle(bcX)*76;
                    bcNode.Y = Convert.ToSingle(-1.0F * depthLE)*76;


                }
                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.MoveDown(Word.WdUnits.wdLine, 1);
                JoistMember joistMember = new JoistMember();
                joistMember.material = material;
                joistMember.member = member;
                joistMember.tcNode = tcNode;
                joistMember.bcNode = bcNode;

                joistMembers.Add(joistMember);
            } while (lineText.Contains("W2R") == false);

            Bitmap myBitmap = new Bitmap(10000,500);
            
            Graphics g = Graphics.FromImage(myBitmap);
            g.Clear(Color.White);
            
            Pen pen = new Pen(Color.Black, 30);
            using (var graphics = Graphics.FromImage(myBitmap))
            {
                foreach (JoistMember joistMember in joistMembers)
                {
                    graphics.DrawLine(pen, joistMember.tcNode, joistMember.bcNode);


                }
            }

            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            
            myBitmap.Save(@"C:\Users\darien.shannon\Documents\newJoist.png", System.Drawing.Imaging.ImageFormat.Png);
          

            //     }
        }
       

        
    }

    public class JoistMember
    {
        public string member;
        public string material;
        public PointF tcNode;
        public PointF bcNode;
    }
}
