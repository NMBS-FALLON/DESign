using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Microsoft.Win32;
using System.Windows;
using System.Text.RegularExpressions;
using DESign_Sales_Excel_Add_in.Worksheet_Values;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace DESign_Sales_Excel_Add_in.BlueBeam
{
    class ExtractBlueBeamMarkups
    {
        public Takeoff TakeoffFromBB()
        {
            Takeoff takeoff = new Takeoff();
            List<Takeoff.Sequence> sequences = new List<Takeoff.Sequence>();
            Takeoff.Sequence sequence = new Takeoff.Sequence();
            List<Joist> joists = new List<Joist>();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "SELECT BLUE BEAM MARKUP FILE (.XML)";
            openFileDialog.Filter = "Markup File (.xml)|*.xml";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (String fileName in openFileDialog.FileNames)
                {
                    //CREATE A COPY OF THE SELECTED FILE AND STORE IT IN A TEMPORARY FILE
                    string markupFileName = System.IO.Path.GetTempFileName();
                    Byte[] markupInByteArray = System.IO.File.ReadAllBytes(fileName);
                    System.IO.File.WriteAllBytes(markupFileName, markupInByteArray);

                    //LOAD THE BLUEBEAM MARKUP FILE INTO AN XElement
                    XElement markUpFile = XElement.Load(markupFileName);

                    //QUERY THE MARKUP XElement TO FIND ALL GIRDER MARKUPS
                    var joistMarkups =
                        from el in markUpFile.Elements("Markup")
                        where ((string)el.Element("Subject") == "GIRDER" || (string)el.Element("Subject") == "JOIST")
                        select el;


                    foreach (XElement joistMarkup in joistMarkups)
                    {
                        Joist joist = new Joist();

                        joist.Mark = new StringWithUpdateCheck { Text = (string)joistMarkup.Element("Label"), IsUpdated = false } ;
                        joist.Quantity = new IntWithUpdateCheck { Value = (int)joistMarkup.Element("Count"), IsUpdated = false };
                        joists.Add(joist);
                    }
                }
            }
            sequence.Joists = joists;
            sequences.Add(sequence);
            takeoff.Sequences = sequences;
            return takeoff;
        }
    }
}
