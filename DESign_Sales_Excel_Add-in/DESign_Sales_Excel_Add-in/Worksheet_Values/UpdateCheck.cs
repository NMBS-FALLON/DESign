using System;


namespace DESign_Sales_Excel_Add_in.Worksheet_Values
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
        public string Text { get; set; }
        private bool hasNoText = false;
        public bool HasNoText
        {
            get
            {
                if (Text == null || Text == "")
                {
                    hasNoText = true;
                }
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
