﻿using System;
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
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == true)
            {
                List<Girder> girders = new List<Girder>();
                List<Joist> joists = new List<Joist>();

                foreach (String fileName in openFileDialog.FileNames)
                {

                    //CREATE A COPY OF THE SELECTED FILE AND STORE IT IN A TEMPORARY FILE
                    string markupFileName = System.IO.Path.GetTempFileName();
                    Byte[] markupInByteArray = System.IO.File.ReadAllBytes(fileName);
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

                    foreach (XElement girderMarkup in girderMarkups)
                    {
                        Girder girder = new Girder();

                        girder.Mark = (string)girderMarkup.Element("Label");
                        girder.Description = (string)girderMarkup.Element("JOIST_DESC");
                        girder.Quantity = (int)girderMarkup.Element("Count");

                        string allNotes = (string)girderMarkup.Element("NOTES");
                        girder.Notes = Regex.Split(allNotes, "\n").ToList();

                        string allLoads = (string)girderMarkup.Element("LOADS");
                        girder.Loads = Regex.Split(allLoads, "\n").ToList();


                        girder.strBaseLength = (string)girderMarkup.Element("J-G_Length");
                        girders.Add(girder);
                    }

                    //EXTRACT ALL JOISTS AND THEIR VALUES FROM THE QUERRIED JOISTS

                    foreach (XElement joistMarkup in joistMarkups)
                    {
                        Joist joist = new Joist();

                        joist.Mark = (string)joistMarkup.Element("Label");
                        joist.Description = (string)joistMarkup.Element("JOIST_DESC");

                        string joistQuantity = (string)joistMarkup.Element("Count");
                        int joistQuantity_int = Convert.ToInt32(joistQuantity.Replace(",", ""));
                        joist.Quantity = joistQuantity_int;

                        string allNotes = (string)joistMarkup.Element("NOTES");
                        joist.Notes = Regex.Split(allNotes, "\n").ToList();

                        string allLoads = (string)joistMarkup.Element("LOADS");
                        joist.Loads = Regex.Split(allLoads, "\n").ToList();

                        joist.strBaseLength = (string)joistMarkup.Element("J-G_Length");

                        joists.Add(joist);
                    }
                }

                //CREATE EXCEL COM OBJECT

                //TEST
                //foreach (Joist joist in joists)
                //{
                //    foreach(string str in joist.Loads)
                //    {
                //        MessageBox.Show(str);
                //    }


                //}

                //ADD JOISTS AND GIRDERS TO JOB
                joists = joists.OrderBy(x => x.StrippedNumber).ToList();
                girders = girders.OrderBy(x => x.StrippedNumber).ToList();
                job.Joists = joists;
                job.Girders = girders;

            }
            return job;
        }
    }
}
