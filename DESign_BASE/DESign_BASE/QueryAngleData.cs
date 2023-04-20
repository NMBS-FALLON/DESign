using System.Linq;
using System.Xml.Linq;
using System.Reflection;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Collections.Generic;
using System;

namespace DESign_BASE
{
    public class Angle
    {
        public Int32 Idx { get; }
        public string SectionName { get; }
        public double LegHorizontal { get; }
        public double LegVertical { get; }
        public double Thickness { get; }

        public Angle(Int32 idx, string sectionName, double legHorizontal, double legVertical, double thickness)
        {
            Idx = idx;
            SectionName = sectionName;
            LegHorizontal = legHorizontal;
            LegVertical = legVertical;
            Thickness = thickness;
        }

    }

    public class GetConnectionString
    {
        public static string FromPlant(string plant)
        {
            var server = "";
            var catalog = "";
            switch (plant)
            {
                case "Fallon":
                    server = "NMBSFALN-SQL";
                    catalog = "NMBS_Fallon";
                    break;
                case "Juarez":
                    server = "NMBSJARZ-SQL";
                    catalog = "NMBS_Juarez";
                    break;
                default:
                    break;
            }

            var connectionString = String.Format("Server={0}; Initial Catalog={1}; Integrated Security = true", server, catalog);

            return connectionString;
        }
    }

    public class QueryAngleData
    {
        static public List<Angle> AnglesFromSql(string plant)
        {

            var angles = new List<Angle>();
            var connectionString = GetConnectionString.FromPlant(plant);
            string queryString = "SELECT SectionIDX, SectionName, LegHorizontal, LegVertical, Thickness FROM dbo.vwMaterials;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                var command = new SqlCommand(queryString, connection);
                connection.Open();
                var reader = command.ExecuteReader();


                try
                {
                    while (reader.Read())
                    {
                        angles.Add(
                            new Angle(
                                Convert.ToInt32(reader["SectionIdx"]),
                                Convert.ToString(reader["SectionName"]),
                                Convert.ToDouble(reader["LegHorizontal"]),
                                Convert.ToDouble(reader["LegVertical"]),
                                Convert.ToDouble(reader["Thickness"])
                                )
                            ); ;
                    }
                }
                finally
                {
                    reader.Close();
                }
            }
            return angles;
        }

        static public string SectionName(List<Angle> angles, Int32 sectionIdx)
        {
            var sectionName =
                angles
                .Where(a => a.Idx == sectionIdx)
                .Select(a => a.SectionName)
                .First();
            return sectionName;
        }

        static public double DblThickness(List<Angle> angles, string tc)
        {
            var thickness =
                angles
                .Where(a => a.SectionName == tc)
                .Select(a => a.Thickness)
                .First();
            return thickness;

        }

        /*
        static public double DblThickness(string tc)
        {
            var asm = Assembly.GetExecutingAssembly();
            var stream = asm.GetManifestResourceStream(AngleData());
            XDocument angleDataXML = XDocument.Load(stream);

            var angleThickness =
                from element in angleDataXML.Descendants("Angle")
                where (string)element.Attribute("section") == tc
                select (double)element.Attribute("thickness");

            double thickness = 0.0;
            foreach (double element in angleThickness)
                thickness = element;

            return thickness;
        }
        */

        static public double DblVleg(List<Angle> angles, string tc)
        {
            var legVertical =
                angles
                .Where(a => a.SectionName == tc)
                .Select(a => a.LegVertical)
                .First();
            return legVertical;

        }

        static public double DblHleg(List<Angle> angles, string tc)
        {
            var legHorizontal =
                angles
                .Where(a => a.SectionName == tc)
                .Select(a => a.LegHorizontal)
                .First();
            return legHorizontal;

        }

        /*
        static public double DblVleg(string tc)
        {
            var asm = Assembly.GetExecutingAssembly();
            var stream = asm.GetManifestResourceStream(AngleData());
            XDocument angleDataXML = XDocument.Load(stream);

            var verticalLeg =
                from element in angleDataXML.Descendants("Angle")
                where (string)element.Attribute("section") == tc
                select (double)element.Attribute("vLeg");

            double vLeg = 0.0;
            foreach (double element in verticalLeg)
                vLeg = element;

            return vLeg;
        }
        */

        static public string WNtcWidth(List<Angle> angles, string tc)
        {
            var legHorizontal =
                angles
                .Where(a => a.SectionName == tc)
                .Select(a => a.LegHorizontal)
                .First();

            var wnWidth = legHorizontal * 2.0 + 1;
            if(legHorizontal == 1.875) { wnWidth = 5.0; }
            if(legHorizontal == 2.875) { wnWidth = 7.0; }
            if (tc == "A32B") { wnWidth = 5.0; }

            return Convert.ToString(wnWidth);

        }

        /*
        static public string WNtcWidth(string tc)
        {
            var asm = Assembly.GetExecutingAssembly();
            var stream = asm.GetManifestResourceStream(AngleData());
            XDocument angleDataXML = XDocument.Load(stream);

            var tcWidth =
                from angle in angleDataXML.Descendants("Angle")
                where (string)angle.Attribute("section") == tc
                select (string)angle.Attribute("wnTCWidth");

            string sTCWidth = "";
            foreach (string element in tcWidth)
                sTCWidth = element;

            return sTCWidth;
        }
        */
        /*
        static public string TypTCWidth(string tc)
        {
            var asm = Assembly.GetExecutingAssembly();
            var stream = asm.GetManifestResourceStream(AngleData());
            XDocument angleDataXML = XDocument.Load(stream);

            var tcWidth =
                from angle in angleDataXML.Descendants("Angle")
                where (string)angle.Attribute("section") == tc
                select (string)angle.Attribute("typTCWidth");

            string sTCWidth = "";
            foreach (string element in tcWidth)
                sTCWidth = element;

            return sTCWidth;
        }
        */
        static public string TypTCWidth(List<Angle> angles, string tc)
        {
            var legHorizontal =
                angles
                .Where(a => a.SectionName == tc)
                .Select(a => a.LegHorizontal)
                .First();

            var typTcWidth = legHorizontal * 2.0 + 1.0;
            return Convert.ToString(typTcWidth);
        }

        static public bool Requres1InchGap(List<Angle> angles, string tc)
        {
            var legHorizontal =
                angles
                .Where(a => a.SectionName == tc)
                .Select(a => a.LegHorizontal)
                .First();

            var requres1InchGap = true;
            if (legHorizontal == 1.875 || legHorizontal == 2.875) { requres1InchGap = false; }

            return requres1InchGap;
        }

        static public object QuerryObject(string inputAttribute, string inputAttributeValue, string returnAttribute)
        {
            var asm = Assembly.GetExecutingAssembly();
            var stream = asm.GetManifestResourceStream(AngleData());
            XDocument angleDataXML = XDocument.Load(stream);
            var tcWidth =
                from angle in angleDataXML.Descendants("Angle")
                where (string)angle.Attribute(inputAttribute) == inputAttributeValue
                select (string)angle.Attribute(returnAttribute);
            object sTCWidth = "";
            foreach (string element in tcWidth)
                sTCWidth = element;
            return sTCWidth;
        }

        static private string AngleData()
        {

            string[] stringArray = Assembly.GetExecutingAssembly().GetManifestResourceNames();
            string angleData = stringArray.FirstOrDefault(s => s.Contains("AngleData.xml"));
            return angleData;
        }

    }
}
