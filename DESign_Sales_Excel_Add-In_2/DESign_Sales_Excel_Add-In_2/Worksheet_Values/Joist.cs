using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DESign_Sales_Excel_Add_In_2.Worksheet_Values
{
  [Serializable]
  public class Joist
  {
    private IntWithUpdateCheck bcxQuantity = new IntWithUpdateCheck {Value = null, IsUpdated = false};
    private bool compositeJoistAdded;
    private StringWithUpdateCheck description = new StringWithUpdateCheck();
    private StringWithUpdateCheck descriptionAdjusted;
    private readonly List<string> errors = new List<string>();

    private bool errorsHaveBeenAdded;
    private bool geometryAdded;
    public bool isComposite;
    private bool isGirder;
    private bool isLoadOverLoad;
    private double? ll;
    private double? tl;
    private double? uDL;
    public StringWithUpdateCheck Mark { get; set; }
    public List<StringWithUpdateCheck> BaseTypesOnMark { get; set; }
    public IntWithUpdateCheck Quantity { get; set; }

    public StringWithUpdateCheck Depth { get; set; }
    public StringWithUpdateCheck Series { get; set; }
    public StringWithUpdateCheck D1 { get; set; }
    public StringWithUpdateCheck D2 { get; set; }
    public StringWithUpdateCheck D3 { get; set; }
    public StringWithUpdateCheck D4 { get; set; }

        public StringWithUpdateCheck Description
    {
      get
      {
        var newDescription = DeepClone(description);
        if (newDescription.Text == "NEW SHEET")
                {
                    newDescription.IsUpdated = Depth.IsUpdated || Series.IsUpdated || D1.IsUpdated || D2.IsUpdated || D3.IsUpdated || D4.IsUpdated;
                    
                    var depth = Depth.Text == null ? "": Depth.Text;
                    var series = Series.Text == null ? "": Series.Text;
                    if (series == "*") { series = "G";  }
                    if (series == "+") { series = "K"; }
                    if (series == "-") { series = "LH";  }
                    if (series.Contains("G"))
                    {
                        var d1 = D1.Text == null ? "" : D1.Text;
                        var d2 = D2.Text == null ? "" : D2.Text;
                        var d3 = D3.Text == null ? "" : D3.Text;
                        var d4 = D4.Text == null ? "" : D4.Text;
                        var extra1 = d3 != "" ? "/" : "";
                        newDescription.Text = depth + series + d1 + "N" + d2 + extra1 + d3 + "K" + d4;
                    }
                    else
                    {
                        var d1 = D1.Text == null ? "" : D1.Text;
                        var d2 = D2.Text == null ? "" : D2.Text;
                        var extra1 = d2 != "" ? "/" : "";
                        newDescription.Text = depth + series + d1 + extra1 + d2;
                    }
                    
                }
        if (newDescription.Text != null && newDescription.Text.Contains("<"))
        {
          string[] seperators = {"<", ">", "-"};
          var geometryValues = newDescription.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
          var geometryNote = "";
          var depth = "";
          if (geometryValues.Length == 3) // single pitch
          {
            var leDepth = Convert.ToDouble(geometryValues[0]);
            var reDepth = Convert.ToDouble(geometryValues[1]);
            var centerDepth = (leDepth + reDepth) / 2.0;
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
            Notes.Add(new StringWithUpdateCheck {Text = geometryNote, IsUpdated = newDescription.IsUpdated});
            geometryAdded = true;
          }
        }

        if (newDescription.Text != null && newDescription.Text.Contains("CJ"))
        {
          isComposite = true;
          string[] seperators = {"CJ", "/"};
          var newDescriptionArray = newDescription.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
          var depth = newDescriptionArray[0];
          var factoredTL = Convert.ToDouble(newDescriptionArray[1]);
          var factoredLL = Convert.ToDouble(newDescriptionArray[2]);
          var cjLoading = newDescription.Text.Split(new[] {"CJ"}, StringSplitOptions.RemoveEmptyEntries)[1];
          var DL = Math.Ceiling((factoredTL - factoredLL) / 1.2 / 5.0) * 5;
          var LL = Math.Ceiling(factoredLL / 1.6 / 5.0) * 5;
          newDescription.Text = depth + "LH" + (DL + LL) + "/" + LL;
          if (compositeJoistAdded == false)
          {
            Notes.Add(new StringWithUpdateCheck
              {Text = "CJ SERIES: " + cjLoading, IsUpdated = newDescription.IsUpdated});
            compositeJoistAdded = true;
          }
        }

        if (newDescription.Text != null)
        {
          newDescription.Text = newDescription.Text.Replace("+-", "KCS");
          newDescription.Text = newDescription.Text.Replace("+", "K");
          newDescription.Text = newDescription.Text.Replace("-", "LH");

          var regex = new Regex(Regex.Escape("*"));
          newDescription.Text = regex.Replace(newDescription.Text, "G", 1);
          newDescription.Text = regex.Replace(newDescription.Text, "N", 1);
          newDescription.Text = regex.Replace(newDescription.Text, "K", 1);
        }

        return newDescription;
      }
      set { description = value; }
    }

    public StringWithUpdateCheck DescriptionAdjusted
    {
      get
      {
        descriptionAdjusted = new StringWithUpdateCheck {Text = Description.Text, IsUpdated = Description.IsUpdated};
        if (IsGirder)
        {
          string[] seperators = {"N", "K"};
          var descriptionSplit = Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
          var load = descriptionSplit[1];
          if (load.Contains("/"))
          {
            var tl = load.Split('/')[0];
            if (descriptionSplit.Length == 2)
              descriptionAdjusted.Text = descriptionSplit[0] + "N" + tl + "K";
            else
              descriptionAdjusted.Text = descriptionSplit[0] + "N" + tl + "K" + descriptionSplit[2];
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

    public IntWithUpdateCheck BcxQuantity
    {
      get
      {
        if (bcxQuantity.Value == -1) bcxQuantity.Value = Quantity.Value * 2;
        if (bcxQuantity.Value == -2) bcxQuantity.Value = Quantity.Value;
        return bcxQuantity;
      }
      set { bcxQuantity = value; }
    }

    public DoubleWithUpdateCheck Uplift { get; set; }
    public List<Load> Loads { get; set; }
    public StringWithUpdateCheck Erfos { get; set; }
    public StringWithUpdateCheck DeflectionTL { get; set; }
    public StringWithUpdateCheck DeflectionLL { get; set; }
    public StringWithUpdateCheck WnSpacing { get; set; }

    public DoubleWithUpdateCheck MinInertia { get; set; }
    public List<StringWithUpdateCheck> Notes { get; set; }

    public bool IsGirder
    {
      get
      {
        if (Description.Text != null && Description.Text.Contains("G")) isGirder = true;
        return isGirder;
      }
    }

    public bool IsLoadOverLoad
    {
      get
      {
        if (Description.Text != null)
          if (Description.Text.Contains("/"))
            isLoadOverLoad = true;

        return isLoadOverLoad;
      }
    }

    public double? TL
    {
      get
      {
        if (IsLoadOverLoad)
        {
          if (IsGirder == false)
          {
            string[] seperators = {"K", "LH", "DLH"};
            tl = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1]
              .Split('/')[0]);
          }

          else
          {
            tl = Convert.ToDouble(GirderLoad().Split('/')[0]);
          }
        }

        return tl;
      }
    }

    public double? LL
    {
      get
      {
        if (IsLoadOverLoad)
        {
          if (IsGirder == false)
          {
            string[] seperators = {"K", "LH", "DLH"};

            ll = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1]
              .Split('/')[1]);
          }
          else
          {
            ll = Convert.ToDouble(GirderLoad().Split('/')[1]);
          }
        }

        return ll;
      }
    }

    public double? UDL
    {
      get
      {
        if (IsLoadOverLoad)
        {
          if (IsGirder == false)
            uDL = TL - LL;
          else
            try
            {
              string[] seperators = {"G", "N", "K"};

              var descriptionSplit = Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
              double? spaces = Convert.ToDouble(descriptionSplit[1]);
              var joistSpace = (BaseLengthFt.Value + BaseLengthIn.Value / 12.0) / spaces;
              if (TL == null || LL == null)
              {
                uDL = null;
              }
              else
              {
                uDL = (TL - LL) * 1000.0 / joistSpace;
                uDL = 5 * (int) Math.Ceiling((float) (uDL / 5.0));
              }
            }
            catch
            {
              MessageBox.Show(string.Format("Mark {0}: Error processing description.", Mark.Text));
              throw;
            }
        }

        return uDL;
      }
    }

    public List<string> ImportErrors
    {
      get
      {
        var importErrors = new List<string>();
        if (Mark.HasNoText) importErrors.Add("Un-named mark.");
        if (Quantity.Value == null) importErrors.Add("No quantity.");
        if (Description.HasNoText) importErrors.Add("No Description.");
        if (BaseLengthFt.Value == null) importErrors.Add("No Base Length.");
        if (IsGirder && Description.Text.Contains("K") == false)
          importErrors.Add("Girder designation is missing a 'K'");
        foreach (var load in Loads) importErrors.AddRange(load.Errors);
        ////
        double? ll;
        if (IsLoadOverLoad)
        {
          if (IsGirder == false)
          {
            string[] seperators = {"K", "LH", "DLH"};
            try
            {
              ll = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1]
                .Split('/')[1]);
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
        if (IsLoadOverLoad)
        {
          if (IsGirder == false)
          {
            string[] seperators = {"K", "LH", "DLH"};
            try
            {
              tl = Convert.ToDouble(Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries)[1]
                .Split('/')[0]);
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

    private string GirderLoad()
    {
      string[] seperators = {"N", "K"};
      var descriptionSplit = Description.Text.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
      string girderLoad = girderLoad = descriptionSplit[1];
      return girderLoad;
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

        return (T) formatter.Deserialize(ms);
      }
    }
  }
}