namespace DESign_Sales_Excel_Add_In_2.Worksheet_Values
{
  public class Bridging
  {
    public string Sequence { get; set; }

        private string size;
    public string Size
        {
            get
            {
                switch (size)
                {
                    case "1 1/4":
                        return "A16B";
                    case "1 1/2":
                        return "1510";
                    case "1 3/4":
                        return "2012";
                    case "2":
                        return "2012";
                    case "2 1/2":
                        return "A44A";
                    case "3":
                        return "3022";
                    case "3 1/2":
                        return "3528";
                    default:
                        return size;
                }
            }
            set
            {
                size = value;
            }
        }
    public string HorX { get; set; }
    public double PlanFeet { get; set; }
  }
}