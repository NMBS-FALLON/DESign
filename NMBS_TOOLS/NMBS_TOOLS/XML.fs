
namespace NMBS_Tools.XML

module XML =
    open System
    open System.Text
    open System.Xml
    open System.Xml.Linq

    let xmlTest() =
        let xmlDocument = new XmlDocument()
        xmlDocument.Load(@"C:\Users\darien.shannon\Desktop\First Nandina\FIRST NANDINA LOGISTICS CENTER_MORENO VALLEY_CA_BB _8_21_17_MSF - Copy.xml")

        ()

    