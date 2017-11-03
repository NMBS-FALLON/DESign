using System;
using System.Linq;
using System.Collections.Generic;

namespace DESign_Sales_Excel_Add_in.Worksheet_Values
{
    [Serializable]
    public class Load
    {
        public StringWithUpdateCheck LoadInfoType { get; set; }
        public StringWithUpdateCheck LoadInfoCategory { get; set; }
        public StringWithUpdateCheck LoadInfoPosition { get; set; }
        private DoubleWithUpdateCheck load1Value = null;
        private bool hasBeenReduced = false;
        public DoubleWithUpdateCheck Load1Value
        {
            get
            {
                if (hasBeenReduced == false)
                {
                    if (load1Value == null || load1Value.Value == null) { }
                    else
                    {
                        if (LoadInfoCategory.Text == "SMU")
                        {
                            load1Value.Value = 1 * (int)Math.Ceiling((decimal)(load1Value.Value * 0.7 / 1.0));
                            hasBeenReduced = true;
                            LoadInfoCategory.Text = "SM";
                        }
                        if(LoadInfoCategory.Text == "WLU")
                        {
                            load1Value.Value = 1 * (int)Math.Ceiling((decimal)(load1Value.Value * 0.6 / 1.0));
                            hasBeenReduced = true;
                            LoadInfoCategory.Text = "WL";
                        }
                    }
                }
                return load1Value;
            }
            set
            {
                load1Value = value;
            }
        }
        public StringWithUpdateCheck Load1DistanceFt { get; set; }
        public DoubleWithUpdateCheck Load1DistanceIn { get; set; }
        public DoubleWithUpdateCheck Load2Value { get; set; }
        public StringWithUpdateCheck Load2DistanceFt { get; set; }
        public DoubleWithUpdateCheck Load2DistanceIn { get; set; }
        public DoubleWithUpdateCheck CaseNumber { get; set; }
        public StringWithUpdateCheck LoadNote { get; set; }
        private bool isNull = true;
        public bool IsNull
        {
            get
            {
                if (LoadInfoType.Text == null &&
                    LoadInfoType.IsUpdated == false &&
                    LoadInfoCategory.Text == null &&
                    LoadInfoCategory.IsUpdated == false &&
                    LoadInfoPosition.Text == null &&
                    LoadInfoPosition.IsUpdated == false &&
                    Load1Value.Value == null &&
                    Load1Value.IsUpdated == false &&
                    Load1DistanceFt.Text == null &&
                    Load1DistanceFt.IsUpdated == false &&
                    Load1DistanceIn.Value == null &&
                    Load1DistanceIn.IsUpdated == false &&
                    Load2Value.Value == null &&
                    Load2Value.IsUpdated == false &&
                    Load2DistanceFt.Text == null &&
                    Load2DistanceFt.IsUpdated == false &&
                    Load2DistanceIn.Value == null &&
                    Load2DistanceIn.IsUpdated == false &&
                    CaseNumber.Value == null &&
                    CaseNumber.IsUpdated == false &&
                    LoadNote.Text == null &&
                    LoadNote.IsUpdated == false)
                {
                    isNull = true;
                }
                else
                {
                    isNull = false;
                }
                return isNull;
            }
  
        }
        public List<string> Errors
        {
            get
            {
                List<string> errors = new List<string>();
                string[] concentratedLoads = { "C", "CB", "CZ", "C3" };
                if (concentratedLoads.Contains(LoadInfoType.Text) && Load1DistanceFt.HasNoText == true)
                {
                    errors.Add("Concentated load without a distance.");
                }
                if (LoadInfoType.HasNoText == true || LoadInfoCategory.HasNoText == true || LoadInfoPosition.HasNoText == true)
                {
                    errors.Add("'Load Info.' column is incomplete.");
                }
                if (Load1Value.Value == null)
                {
                    errors.Add("No value given in 'Load 1' column.");
                }
                if (Load2Value.Value == null && Load2DistanceFt.HasNoText == false)
                {
                    errors.Add("Value missing in 'Load 2' column.");
                }
                if (Load2Value.Value != null && Load2DistanceFt.HasNoText == true)
                {
                    errors.Add("Distance missing in 'Load 2' column.");
                }
                bool isWL = LoadInfoCategory.Text == "WL";
                bool isNegative = Load1Value.Value < 0.0;
                bool isBackedOutLoad = (LoadInfoType.Text == "CP" && LoadInfoCategory.Text == "CL" && isNegative == true);
                bool isInLC1 = (CaseNumber.Value == 1 || CaseNumber.Value == null);
                
                if (isWL == false && isNegative == true && isInLC1 == true && isBackedOutLoad == false)
                {
                    errors.Add("Non-WL negative value in LC1; Please confirm.");
                }
                return errors;
            }
            
        }

    }



}
