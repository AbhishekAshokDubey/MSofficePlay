using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;

/// <summary>
/// https://msdn.microsoft.com/en-us/magazine/ff796230.aspx
/// </summary>

namespace ConsoleApplication3
{

    class Program
    {
        static void Main(string[] args)
        {
            var onenoteApp = new Application();

            string notebookXml;
            onenoteApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;

          foreach (var notebookNode in from node in doc.Descendants(ns +
          "Page") select node)
            {
               /// Console.WriteLine(notebookNode.Attribute("name").Value);
                if (notebookNode.Attribute("name").Value == "Now me")
                {
                    string pageXml;
                    onenoteApp.GetPageContent(notebookNode.Attribute("ID").Value, out pageXml);

                    var parsedXML = XDocument.Parse(pageXml);
                    ///var root = parsedXML.Root;
                    for (int i = 0; i < parsedXML.Descendants(ns + "OCRText").Count(); i++)
                    {
                        ///XElement OCRTXML = parsedXML.Descendants(ns + "OCRText").FirstOrDefault();
                        XElement OCRTXML = parsedXML.Descendants(ns + "OCRText").ElementAt(i);
                        Console.Write(OCRTXML.Value);
                        Console.Write("==========================================================");
                    }                   
                    
                    Console.Write("Done");
                }
            }
        }
    }

}
