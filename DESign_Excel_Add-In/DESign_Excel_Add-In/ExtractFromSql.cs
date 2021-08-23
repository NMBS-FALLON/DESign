using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace DESign_Excel_Add_In
{
    class Helpers
    {
        static public double FractionToDecimal(string fraction)
        {
            var splitFraction = fraction.Split(new char[] { ' ' });
            var inches = Convert.ToDouble(splitFraction[0]);
            var numerator = 0.0;
            var denominator = 0.0;
            if (splitFraction.Length > 1)
            {
                var splitFraction_ = splitFraction[1].Split(new char[] { '/' });
                numerator = Convert.ToDouble(splitFraction_[0]);
                denominator = Convert.ToDouble(splitFraction_[1]);
            }

            var decimalLength = inches + numerator / denominator;
            return decimalLength;
        }

        static public (double Feet, double Inch) HyphenToFtIn(string hyphenLength)
        {
            var splitFtIn = hyphenLength.Split(new char[] { '-' });
            var feet = Convert.ToDouble(splitFtIn[0]);
            var inch = FractionToDecimal(splitFtIn[1]);
            return (Feet: feet, Inch: inch);
        }
    }
    class Joists2
    {
        public string Mark { get; set; }
        public int Quantity { get; set; }
        public double SeatsBdl { get; set; }
        public double SeatsBdr { get; set; }
        public double TcPanelLeFt { get; set; }
        public double TcPanelLeIn { get; set; }
        public double TcPanelReFt { get; set; }
        public double TcPanelReIn { get; set; }
        public double FirstDiagLeFt { get; set; }
        public double FirstDiagLeIn { get; set; }
        public double FirstDiagReFt { get; set; }
        public double FirstDiagReIn { get; set; }
        public double DepthLe { get; set; }
        public double DepthRe { get; set; }
        public double BcPanelLeFt { get; set; }
        public double BcPanelLeIn { get; set; }
        public double BcPanelReFt { get; set; }
        public double BcPanelReIn { get; set; }
        public string TcSection { get; set; }
        public string BcSection { get; set; }
        public double Axial { get; set; }
        public double TcxlIn { get; set; }
        public double TcxrIn { get; set; }

        public Joists2(string mark, int quantity, double seatsBdl, double seatsBdr, double tcPanelLeFt,
            double tcPanelLeIn, double tcPanelReFt, double tcPanelReIn, double firstDiagLeFt, double firstDiagLeIn,
            double firstDiagReFt, double firstDiagReIn, double depthLe, double depthRe, double bcPanelLeFt, double bcPanelLeIn,
            double bcPanelReFt, double bcPanelReIn,
            string tcSection, string bcSection, double axial, double tcxlIn, double tcxrIn)
        {
            Mark = mark; Quantity = quantity; SeatsBdl = seatsBdl; SeatsBdr = seatsBdr; TcPanelLeFt = tcPanelLeFt;
            TcPanelLeIn = tcPanelLeIn; TcPanelReFt = tcPanelReFt; TcPanelReIn = tcPanelReIn; FirstDiagLeFt = firstDiagLeFt;
            FirstDiagLeIn = firstDiagLeIn; FirstDiagReFt = firstDiagReFt; FirstDiagReIn = firstDiagReIn;
            DepthLe = depthLe; DepthRe = depthRe; BcPanelLeFt = bcPanelLeFt; BcPanelLeIn = bcPanelLeIn;
            BcPanelReFt = bcPanelReFt; BcPanelReIn = bcPanelReFt;
            TcSection = tcSection; BcSection = bcSection; Axial = axial; TcxlIn = tcxlIn; TcxrIn = tcxrIn;

        }

    }
    class ExtractFromSql
    {
        static public Dictionary<string, string> PlantAbbreviation()
        {
            var plantAbbreviation = new Dictionary<string, string>();
            plantAbbreviation.Add("Juarez", "JARZ");
            plantAbbreviation.Add("Fallon", "FALN");

            return plantAbbreviation;
        }

        static public string ConnectionString(string plant)
        {
            var server = string.Format("NMBS{0}-SQL", PlantAbbreviation()[plant]);
            var initialCatalog = string.Format("NMBS_{0}", plant);

            var connectionString =
                String.Format("Server={0}; Initial Catalog={1}; Integrated Security = true", server, initialCatalog);

            return connectionString;
        }

        static public Dictionary<int, string> Sections(string plant)
        {
            var vwMaterials = new Dictionary<int, string>();

            var connectionString = ConnectionString(plant);
            string queryString = "SELECT SectionIDX, Name FROM dbo.Section;";


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                var command = new SqlCommand(queryString, connection);
                connection.Open();
                var reader = command.ExecuteReader();


                try
                {
                    while (reader.Read())
                    {
                        vwMaterials.Add(
                                Convert.ToInt32(reader["SectionIDX"]),
                                Convert.ToString(reader["Name"])
                            );
                    }
                }
                finally
                {
                    reader.Close();
                }
            }
            return vwMaterials;
        }

        static public List<(string Mark, string Side, string Section)> EngSeats(string plant, string jobNumber)
        {
            var engSeats = new List<(string Mark, string Side, string Section)>();

            var connectionString = ConnectionString(plant);
            string queryString =
                String.Format("SELECT Mark, SectionIDX, Side FROM dbo.EngSeats WHERE JobNumber='{0}'", jobNumber);

            var sections = Sections(plant);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                var command = new SqlCommand(queryString, connection);
                connection.Open();
                var reader = command.ExecuteReader();


                try
                {
                    while (reader.Read())
                    {
                        var section = sections[Convert.ToInt32(reader["SectionIDX"])];
                        engSeats.Add(
                                (Mark: Convert.ToString(reader["Mark"]),
                                    Side: Convert.ToString(reader["Side"]),
                                    Section: section)
                            );
                    }
                }
                finally
                {
                    reader.Close();
                }
            }
            return engSeats;
        }

        static public List<(string Mark, string Web, int ComponentNo, string Section)> EngWebs(string plant, string jobNumber)
        {
            var engWebs = new List<(string Mark, string Web, int ComponentNo, string Section)>();

            var connectionString = ConnectionString(plant);
            string queryString =
                String.Format("SELECT Mark, SectionIDX, Web, ComponentNo FROM dbo.EngWebs WHERE JobNumber='{0}'", jobNumber);

            var sections = Sections(plant);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                var command = new SqlCommand(queryString, connection);
                connection.Open();
                var reader = command.ExecuteReader();


                try
                {
                    while (reader.Read())
                    {
                        var section = sections[Convert.ToInt32(reader["SectionIDX"])];
                        engWebs.Add(
                                (Mark: Convert.ToString(reader["Mark"]),
                                 Web: Convert.ToString(reader["Web"]),
                                 ComponentNo: Convert.ToInt32(reader["ComponentNo"]),
                                 Section: section)
                            );
                    }

                }
                finally
                {
                    reader.Close();
                }
            }
            return engWebs;
        }

        static public List<Joists2> Joist22(string plant, string jobNumber)
        {
            var joists2 = new List<Joists2>();

            var connectionString = ConnectionString(plant);
            string queryString =
                String.Format("SELECT [Seats BDL], [Seats BDR], [TCXL], [TCXR], [TC Panels LE], [TC Panels RE], [First Diag LE], [First Diag RE], [Quantity], [Mark], [BC Panels LE], [BC Panels RE], [Depth LE], [Depth RE], [TopChord_IDX], [BottomChord_IDX], [Extras] FROM dbo.Joists2 WHERE [Job Number]='{0}'", jobNumber);

            var sections = Sections(plant);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                var command = new SqlCommand(queryString, connection);
                connection.Open();
                var reader = command.ExecuteReader();


                try
                {
                    while (reader.Read())
                    {
                        var seatsBdl = Helpers.FractionToDecimal(Convert.ToString(reader["Seats BDL"]));
                        var seatsBdr = Helpers.FractionToDecimal(Convert.ToString(reader["Seats BDR"]));
                        var (tcxlFt, tcxlIn) = Helpers.HyphenToFtIn(Convert.ToString(reader["TCXL"]));
                        var (tcxrFt, tcxrIn) = Helpers.HyphenToFtIn(Convert.ToString(reader["TCXR"]));
                        var (tcPanelLeFt, tcPanelLeIn) = Helpers.HyphenToFtIn(Convert.ToString(reader["TC Panels LE"]));
                        var (tcPanelReFt, tcPanelReIn) = Helpers.HyphenToFtIn(Convert.ToString(reader["TC Panels RE"]));
                        var (FirstDiagLeFt, FirstDiagLeIn) = Helpers.HyphenToFtIn(Convert.ToString(reader["First Diag LE"]));
                        var (FirstDiagReFt, FirstDiagReIn) = Helpers.HyphenToFtIn(Convert.ToString(reader["First Diag RE"]));
                        var quantity = Convert.ToInt32(reader["Quantity"]);
                        var mark = Convert.ToString(reader["Mark"]);
                        var (bcPanelLeFt, bcPanelLeIn) = Helpers.HyphenToFtIn(Convert.ToString(reader["BC Panels LE"]));
                        var (bcPanelReFt, bcPanelReIn) = Helpers.HyphenToFtIn(Convert.ToString(reader["BC Panels RE"]));
                        var depthLe = Convert.ToDouble(reader["Depth LE"]);
                        var depthRe = Convert.ToDouble(reader["Depth RE"]);
                        var tcSection = sections[Convert.ToInt32(reader["TopChord_IDX"])];
                        var bcSection = sections[Convert.ToInt32(reader["BottomChord_IDX"])];
                        var axial = 1000.0;
                        var joist2 = new Joists2(
                                mark, quantity, seatsBdl, seatsBdr, tcPanelLeFt,
                                tcPanelLeIn, tcPanelReFt, tcPanelReIn, FirstDiagLeFt, FirstDiagLeIn,
                                FirstDiagReFt, FirstDiagReIn, depthLe, depthRe, bcPanelLeFt, bcPanelLeIn, bcPanelReFt, bcPanelReIn,
                                tcSection, bcSection, axial, tcxlFt * 12.0 + tcxlIn, tcxrFt * 12.0 + tcxrIn);
                        joists2.Add(joist2);
                    }

                }
                finally
                {
                    reader.Close();
                }
            }
            return joists2;
        }

        static public List<string> MarksInShoporder(string plant, string jobNumber, string shopOrderNumber)
        {
            var marks = new List<string>();

            var connectionString = ConnectionString(plant);
            string queryString =
                String.Format("SELECT Mark FROM dbo.ShopordList WHERE [Job Number]='{0}' AND [List Number]='{1}'", jobNumber, shopOrderNumber);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                var command = new SqlCommand(queryString, connection);
                connection.Open();
                var reader = command.ExecuteReader();


                try
                {
                    while (reader.Read())
                    {
                        marks.Add(Convert.ToString(reader["Mark"]));                     
                    }
                }
                finally
                {
                    reader.Close();
                }
            }
            return marks;
        }

    }
}
