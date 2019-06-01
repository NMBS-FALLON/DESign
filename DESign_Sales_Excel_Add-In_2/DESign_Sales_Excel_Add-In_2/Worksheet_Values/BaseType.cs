using System;
using System.Collections.Generic;

namespace DESign_Sales_Excel_Add_In_2.Worksheet_Values
{
  [Serializable]
  public class BaseType
  {
    public StringWithUpdateCheck Name { get; set; }
    public List<string> BaseTypeStrings { get; set; }
    public StringWithUpdateCheck Depth { get; set; }
    public StringWithUpdateCheck Series { get; set; }
    public StringWithUpdateCheck D1 { get; set; }
    public StringWithUpdateCheck D2 { get; set; }
    public StringWithUpdateCheck D3 { get; set; }
    public StringWithUpdateCheck D4 { get; set; }

    public StringWithUpdateCheck Description { get; set; }
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
    public IntWithUpdateCheck BcxQuantity { get; set; }
    public DoubleWithUpdateCheck Uplift { get; set; }
    public List<Load> Loads { get; set; }
    public StringWithUpdateCheck Erfos { get; set; }
    public StringWithUpdateCheck DeflectionTL { get; set; }
    public StringWithUpdateCheck DeflectionLL { get; set; }
    public StringWithUpdateCheck WnSpacing { get; set; }

    public DoubleWithUpdateCheck MinInertia { get; set; }
    public List<StringWithUpdateCheck> Notes { get; set; }

    public List<string> Errors { get; } = new List<string>();

    public void AddError(string error)
    {
      Errors.Add(error);
    }
  }
}