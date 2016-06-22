using System.Linq;
using System.Xml.Linq;
using System.Reflection;

namespace DESign_BASE_WPF_WPF
{
    public class QueryAngleData
    {

        public double DblThickness(string tc)
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
        public double DblVleg(string tc)
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
        public string WNtcWidth(string tc)
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

        public string TypTCWidth(string tc)
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

        public object QuerryObject (string inputAttribute, string inputAttributeValue, string returnAttribute)
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

        private string AngleData()
        {
            
            string[] stringArray = Assembly.GetExecutingAssembly().GetManifestResourceNames();
            string angleData = stringArray.FirstOrDefault(s => s.Contains("AngleData.xml"));
            return angleData;
        }

    }
}
