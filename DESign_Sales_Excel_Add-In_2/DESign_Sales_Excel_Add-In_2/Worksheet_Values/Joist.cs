﻿using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace DESign_Sales_Excel_Add_In_2.Worksheet_Values
{
    [Serializable]
    public class Joist
    {
        public StringWithUpdateCheck Mark { get; set; }
        public List<StringWithUpdateCheck> BaseTypesOnMark { get; set; }
        public IntWithUpdateCheck Quantity { get; set; }
        private bool geometryAdded = false;
        private bool compositeJoistAdded = false;
        private StringWithUpdateCheck description = new StringWithUpdateCheck { };
        public StringWithUpdateCheck Description
        {
            get
            {
                StringWithUpdateCheck newDescription = DeepClone(description);
                if (newDescription.Text != null && newDescription.Text.Contains("<"))
                {
                    string[] seperators = { "<", ">", "-" };
                    string[] geometryValues = newDescription.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
                    string geometryNote = "";
                    string depth = "";
                    if (geometryValues.Length == 3) // single pitch
                    {
                        double leDepth = Convert.ToDouble(geometryValues[0]);
                        double reDepth = Convert.ToDouble(geometryValues[1]);
                        double centerDepth = (leDepth + reDepth) / 2.0;
                        depth = centerDepth.ToString("0.#");
                        geometryNote = string.Format("SP: {0}/{1}", geometryValues[0], geometryValues[1]);
                    }
                    if (geometryValues.Length == 4) // double pitch
                    {
                        depth = geometryValues[1];
                        geometryNote = string.Format("DP: {0}/{1}/{2}", geometryValues[0], geometryValues[1], geometryValues[2]);
                    }
                    newDescription.Text = depth + newDescription.Text.Substring(newDescription.Text.IndexOf('>') + 1);
                    if (geometryAdded == false)
                    {
                        Notes.Add(new StringWithUpdateCheck { Text = geometryNote, IsUpdated = newDescription.IsUpdated });
                        geometryAdded = true;
                    }
                }

                if (newDescription.Text != null && newDescription.Text.Contains("CJ"))
                {
                    string[] seperators = { "CJ", "/" };
                    string[] newDescriptionArray = newDescription.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
                    string depth = newDescriptionArray[0];
                    double factoredTL = Convert.ToDouble(newDescriptionArray[1]);
                    double factoredLL = Convert.ToDouble(newDescriptionArray[2]);
                    string cjLoading = newDescription.Text.Split(new string[] { "CJ" }, StringSplitOptions.RemoveEmptyEntries)[1];             
                    double DL = Math.Ceiling(((factoredTL - factoredLL) / 1.2)/5.0) * 5;
                    double LL = Math.Ceiling((factoredLL / 1.6)/5.0) * 5;
                    newDescription.Text = depth + "LH" + (DL + LL).ToString() + "/" + LL.ToString();
                    if (compositeJoistAdded == false)
                    {
                        Notes.Add(new StringWithUpdateCheck { Text = "CJ SERIES: " + cjLoading, IsUpdated = newDescription.IsUpdated });
                        compositeJoistAdded = true;
                    }
                }

                if (newDescription.Text != null)
                {
                    newDescription.Text = newDescription.Text.Replace("+", "K");
                    newDescription.Text = newDescription.Text.Replace("-", "LH");

                    var regex = new Regex(Regex.Escape("*"));
                    newDescription.Text = regex.Replace(newDescription.Text, "G", 1);
                    newDescription.Text = regex.Replace(newDescription.Text, "N", 1);
                    newDescription.Text = regex.Replace(newDescription.Text, "K", 1);
                }
                return newDescription;
            }
            set
            {
                description = value;
            }
        }
        private StringWithUpdateCheck descriptionAdjusted;
        public StringWithUpdateCheck DescriptionAdjusted
        {
            get
            {
                descriptionAdjusted = new StringWithUpdateCheck { Text = Description.Text, IsUpdated = Description.IsUpdated };
                if(IsGirder == true)
                {

                    string[] seperators = { "N", "K" };
                    string[] descriptionSplit = Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
                    string load = descriptionSplit[1];
                    if (load.Contains("/") == true)
                    {

                        string tl = load.Split('/')[0];
                        if (descriptionSplit.Length == 2)
                        {
                            descriptionAdjusted.Text = descriptionSplit[0] + "N" + tl + "K";
                        }
                        else
                        {
                            descriptionAdjusted.Text = descriptionSplit[0] + "N" + tl + "K" + descriptionSplit[2];
                        }
                    }
                }
                return descriptionAdjusted;
                    
            }
        }
        public DoubleWithUpdateCheck BaseLengthFt { get; set; }
        public DoubleWithUpdateCheck BaseLengthIn { get; set; }
        public IntWithUpdateCheck TcxlQuantity { get; set; }
        public DoubleWithUpdateCheck TcxlLengthFt { get; set; }
        public DoubleWithUpdateCheck TcxlLengthIn { get; set; }
        public IntWithUpdateCheck TcxrQuantity { get; set; }
        public DoubleWithUpdateCheck TcxrLengthFt { get; set; }
        public DoubleWithUpdateCheck TcxrLengthIn { get; set; }
        public DoubleWithUpdateCheck SeatDepthLE { get; set; }
        public DoubleWithUpdateCheck SeatDepthRE { get; set; }
        private IntWithUpdateCheck bcxQuantity = new IntWithUpdateCheck {Value = null, IsUpdated = false };
        public IntWithUpdateCheck BcxQuantity
        {
            get
            {
                if (bcxQuantity.Value == -1)
                {
                    bcxQuantity.Value = this.Quantity.Value * 2;
                }
                return bcxQuantity;
            }
            set
            {
                bcxQuantity = value;
            }
        }
        public DoubleWithUpdateCheck Uplift { get; set; }
        public List<Load> Loads { get; set; }
        public StringWithUpdateCheck Erfos { get; set; }
        public DoubleWithUpdateCheck DeflectionTL { get; set; }
        public DoubleWithUpdateCheck DeflectionLL { get; set; }
        public StringWithUpdateCheck WnSpacing { get; set; }
        public List<StringWithUpdateCheck> Notes { get; set; }
        private bool isGirder = false;
        public bool IsGirder
        {
            get
            {
                if (Description.Text != null && Description.Text.Contains("G") == true)
                {
                    isGirder = true;
                }
                return isGirder;
            }
        }
        private bool isLoadOverLoad = false;
        public bool IsLoadOverLoad
        {
            get
            {

                if (Description.Text != null)
                {
                    if (Description.Text.Contains("/") == true)
                    {
                        isLoadOverLoad = true;
                    }
                }

                return isLoadOverLoad;
            }
        }
        private double? tl;
        public double? TL
        {
            get
            {
                if (IsLoadOverLoad == true)
                {
                    if (IsGirder == false)
                    {
                        string[] seperators = { "K", "LH", "DLH" };
                        tl = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[0]);
                    }

                    else
                    {

                        tl = Convert.ToDouble(GirderLoad().Split('/')[0]);
                    }
                }
            
                return tl;
            }
        }
        private double? ll;
        public double? LL
        {
            get
            {
                if (IsLoadOverLoad == true)
                {
                    if (IsGirder == false)
                    {
                        string[] seperators = { "K", "LH", "DLH" };

                        ll = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[1]);

                    }
                    else
                    {
                        ll = Convert.ToDouble(GirderLoad().Split('/')[1]);

                    }
                }
                return ll;
            }
        }
        private double? uDL;
        public double? UDL
        {
            get
            {
                if (IsLoadOverLoad == true)
                {
                    if (IsGirder == false)
                    {
                        uDL = TL - LL;
                    }
                    else
                    {
                        try
                        {
                            string[] seperators = { "G", "N", "K" };

                            string[] descriptionSplit = Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
                            double? spaces = Convert.ToDouble(descriptionSplit[1]);
                            double? joistSpace = (BaseLengthFt.Value + BaseLengthIn.Value / 12.0) / spaces;
                            if (TL == null || LL == null)
                            {
                                uDL = null;
                            }
                            else
                            {
                                uDL = ((TL - LL) * 1000.0) / joistSpace;
                                uDL = 5 * (int)Math.Ceiling((float)(uDL / 5.0));
                            }
                        }
                        catch
                        {
                            MessageBox.Show(String.Format("Mark {0}: Error processing description.", Mark.Text));
                            throw;
                        }
                    }
                }
                return uDL;
            }
        }

        private string GirderLoad()
        {
            string[] seperators = { "N", "K" };
            string[] descriptionSplit = Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
            string girderLoad = girderLoad = descriptionSplit[1];
            return girderLoad;
        }

        public List<string> ImportErrors
        {
            get
            {
                List<string> importErrors = new List<string>();
                if (Mark.HasNoText == true) { importErrors.Add("Un-named mark."); }
                if (Quantity.Value == null) { importErrors.Add("No quantity."); }
                if (Description.HasNoText == true) { importErrors.Add("No Description."); }
                if (BaseLengthFt.Value == null) { importErrors.Add("No Base Length."); }
                if (IsGirder == true && Description.Text.Contains("K")== false) { importErrors.Add("Girder designation is missing a 'K'"); }
                foreach (Load load in Loads)
                {
                    importErrors.AddRange(load.Errors);
                }
                ////
                double? ll;
                if (IsLoadOverLoad == true)
                {
                    if (IsGirder == false)
                    {
                        string[] seperators = { "K", "LH", "DLH" };
                        try
                        {
                            ll = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[1]);
                        }
                        catch (Exception ex)
                        {
                            importErrors.Add("There is an issue with the description");
                        }
                    }
                    else
                    {
                        try
                        {
                            ll = Convert.ToDouble(GirderLoad().Split('/')[1]);
                            
                        }
                        catch (Exception ex)
                        {
                            importErrors.Add("There is an issue with the description");
                        }
                    }
                }
                double? tl;
                if (IsLoadOverLoad == true)
                {
                    if (IsGirder == false)
                    {
                        string[] seperators = { "K", "LH", "DLH" };
                        try
                        {
                            tl = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1].Split('/')[0]);
                        }
                        catch (Exception ex)
                        {
                            importErrors.Add("There is an issue with the description");
                        }
                    }

                    else
                    {
                        try
                        {
                            tl = Convert.ToDouble(GirderLoad().Split('/')[0]);
                        }
                        catch (Exception ex)
                        {
                            importErrors.Add("There is an issue with the description");

                        }
                    }
                }
                ////

                return importErrors;
            }
        }

        private bool errorsHaveBeenAdded = false;
        private List<string> errors = new List<string>();
        public List<string> Errors
        {
            get
            {
                if (errorsHaveBeenAdded == false)
                {
                    errors.AddRange(ImportErrors);
                    errorsHaveBeenAdded = true;
                }

                return errors;
            }
        }
        public void AddError(string error)
        {
            errors.Add(error);
        }

        public static T DeepClone<T>(T obj)
        {
            using (var ms = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(ms, obj);
                ms.Position = 0;

                return (T)formatter.Deserialize(ms);
            }
        }

    }

    

}