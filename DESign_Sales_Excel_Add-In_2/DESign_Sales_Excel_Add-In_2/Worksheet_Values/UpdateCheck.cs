using System;

namespace DESign_Sales_Excel_Add_In_2.Worksheet_Values
{
  [Serializable]
  public class UpdateCheck
  {
    public bool IsUpdated { get; set; } = false;
  }

  [Serializable]
  public class DoubleWithUpdateCheck : UpdateCheck
  {
    public double? Value { get; set; }
  }

  [Serializable]
  public class StringWithUpdateCheck : UpdateCheck
  {
    private bool hasNoText;
    private string text;

    public string Text
    {
      get
      {
        if (text != null) text = text.Trim();
        return text;
      }
      set { text = value; }
    }

    public bool HasNoText
    {
      get
      {
        if (Text == null || Text == "") hasNoText = true;
        return hasNoText;
      }
    }
  }

  [Serializable]
  public class IntWithUpdateCheck : UpdateCheck
  {
    public int? Value { get; set; }
  }
}