using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using System.Xml.Linq;
using System.Reflection;

namespace DESign_WordAddIn
{
    class QueryAngleDataXML
    {
        public double DblThickness (string tc)
        {
                var asm = Assembly.GetExecutingAssembly();
                var stream = asm.GetManifestResourceStream("DESign_WordAddIn.Resources.AngleData.xml");
                XDocument angleDataXML = XDocument.Load(stream);

                var angleThickness =
                    from element in angleDataXML.Descendants("Angle")
                    where (string)element.Attribute("section") == tc
                    select (double) element.Attribute("thickness");

                double thickness = 0.0;
                foreach (double element in angleThickness)
                    thickness=element;
               
            return thickness;
        }
        public double DblVleg(string tc)
        {
            var asm = Assembly.GetExecutingAssembly();
            var stream = asm.GetManifestResourceStream("DESign_WordAddIn.Resources.AngleData.xml");
            XDocument angleDataXML = XDocument.Load(stream);

            var verticalLeg =
                from element in angleDataXML.Descendants("Angle")
                where (string)element.Attribute("section") == tc
                select (double) element.Attribute("vLeg");

            double vLeg = 0.0;
            foreach (double element in verticalLeg)
                vLeg = element;

            return vLeg;
        }

    }
}
