using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data;
using Microsoft.Win32;
using System.Windows;
using System.Text.RegularExpressions;

namespace DESign_BASE_WPF
{
    class ExtractBlueBeamMarkups
    {
       
        public Job JobFromBlueBeamMarkups()
        {
            Job job = new Job();
            List<Joist> allJoists = new List<Joist>();
            List<Girder> allGirders = new List<Girder>();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "SELECT BLUE BEAM MARKUP FILE (.XML)";

            openFileDialog.Filter = "Markup File (.xml)|*.xml";
            if (openFileDialog.ShowDialog() == true)
            {
                string fileName = openFileDialog.FileName;

                //CREATE A COPY OF THE SELECTED FILE AND STORE IT IN A TEMPORARY FILE
                string markupFileName = System.IO.Path.GetTempFileName();
                Byte[] markupInByteArray = System.IO.File.ReadAllBytes(openFileDialog.FileName);
                System.IO.File.WriteAllBytes(markupFileName, markupInByteArray);

                //LOAD THE BLUEBEAM MARKUP FILE INTO AN XElement
                XElement markUpFile = XElement.Load(markupFileName);

                //QUERY THE MARKUP XElement TO FIND ALL GIRDER MARKUPS
                var girderMarkups =
                    from el in markUpFile.Elements("Markup")
                    where (string)el.Element("Subject") == "GIRDER" 
                    select el;

                //QUERY THE MARKUP XElement TO FIN ALL JOIST MARKUPS
                var joistMarkups =
                    from el in markUpFile.Elements("Markup")
                    where (string)el.Element("Subject") == "JOIST"
                    select el;


                //EXTRACT ALL GIRDERS AND THEIR VALUES FROM THE QUERRIED GIRDERS
                List<Girder> girders = new List<Girder>();  
                foreach (XElement girderMarkup in girderMarkups)
                {
                    Girder girder = new Girder();

                    girder.Mark = (string)girderMarkup.Element("Label");
                    girder.Description = (string)girderMarkup.Element("girder_DESC");
                    girder.Quantity = (int)girderMarkup.Element("Count");

                    string allNotes = (string)girderMarkup.Element("NOTES");
                    girder.Notes = Regex.Split(allNotes, "\r\n").ToList();

                    string allLoads = (string)girderMarkup.Element("LOADS");
                    girder.Loads = Regex.Split(allLoads, "\n").ToList();

                    girders.Add(girder);
                }

                //EXTRACT ALL JOISTS AND THEIR VALUES FROM THE QUERRIED JOISTS
                List<Joist> joists = new List<Joist>();
                foreach (XElement joistMarkup in joistMarkups)
                {
                    Joist joist = new Joist();

                    joist.Mark = (string)joistMarkup.Element("Label");
                    joist.Description = (string)joistMarkup.Element("JOIST_DESC");
                    joist.Quantity = (int)joistMarkup.Element("Count");

                    string allNotes = (string)joistMarkup.Element("NOTES");
                    joist.Notes = Regex.Split(allNotes, "\r\n").ToList();

                    string allLoads = (string)joistMarkup.Element("LOADS");
                    joist.Loads = Regex.Split(allLoads, "\n").ToList();

                    joists.Add(joist);
                }

                //TEST
                foreach (Joist joist in joists)
                {
                    foreach(string str in joist.Loads)
                    {
                        MessageBox.Show(str);
                    }
                    

                }
            }
            return job;
        }
    }
}
