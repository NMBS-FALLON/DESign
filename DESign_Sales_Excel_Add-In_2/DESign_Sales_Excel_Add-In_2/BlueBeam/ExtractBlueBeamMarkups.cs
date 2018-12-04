using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DESign_Sales_Excel_Add_In_2.Worksheet_Values;
using Microsoft.Win32;

namespace DESign_Sales_Excel_Add_In_2.BlueBeam
{
  internal class ExtractBlueBeamMarkups
  {
    public Takeoff TakeoffFromBB()
    {
      var takeoff = new Takeoff();
      var sequences = new List<Takeoff.Sequence>();
      var sequence = new Takeoff.Sequence();
      var joists = new List<Joist>();

      var openFileDialog = new OpenFileDialog();
      openFileDialog.Title = "SELECT BLUE BEAM MARKUP FILE (.XML)";
      openFileDialog.Filter = "Markup File (.xml)|*.xml";
      openFileDialog.Multiselect = true;

      if (openFileDialog.ShowDialog() == true)
        foreach (var fileName in openFileDialog.FileNames)
        {
          //CREATE A COPY OF THE SELECTED FILE AND STORE IT IN A TEMPORARY FILE
          var markupFileName = Path.GetTempFileName();
          var markupInByteArray = File.ReadAllBytes(fileName);
          File.WriteAllBytes(markupFileName, markupInByteArray);

          //LOAD THE BLUEBEAM MARKUP FILE INTO AN XElement
          var markUpFile = XElement.Load(markupFileName);

          //QUERY THE MARKUP XElement TO FIND ALL GIRDER MARKUPS
          var joistMarkups =
            from el in markUpFile.Elements("Markup")
            where (string) el.Element("Subject") == "GIRDER" || (string) el.Element("Subject") == "JOIST"
            select el;


          foreach (var joistMarkup in joistMarkups)
          {
            var joist = new Joist();

            joist.Mark = new StringWithUpdateCheck {Text = (string) joistMarkup.Element("Label"), IsUpdated = false};
            joist.Quantity = new IntWithUpdateCheck {Value = (int) joistMarkup.Element("Count"), IsUpdated = false};
            joists.Add(joist);
          }
        }

      sequence.Joists = joists;
      sequences.Add(sequence);
      takeoff.Sequences = sequences;
      return takeoff;
    }
  }
}