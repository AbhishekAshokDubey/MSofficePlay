using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;

/// <summary>
/// https://msdn.microsoft.com/en-us/magazine/ff796230.aspx
/// </summary>

namespace ConsoleApplication3
{
    class Program
    {

        /*
        static Application onenoteApp = new Application();
        static XNamespace ns = null;
        //static string imgDirPath = @"C:\Users\ad12183\Desktop\images";
        */

        public static class Globals
        {
            //public static string imgDirPath = @"C:\Users\ad12183\Desktop\images"; // Modifiable in Code
            public const string imgDirPath = @"C:\Users\ad12183\Desktop\images\images\extractedImages"; // Unmodifiable
            public static Int16 imageCount = 0;
            public const string oneNoteName = "OCRImages"; // Unmodifiable
            public static List<string> imageNames = new List<string>();
        }

        /*
        static void GetNamespace()
        {
            string xml;
            onenoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);

            var doc = XDocument.Parse(xml);
            ns = doc.Root.Name.Namespace;
        }
        */

        static void waitScreen(int secondsCount) {
            for (int i = secondsCount; i > 0 ; i--) {
                Console.Clear();
                Console.SetCursorPosition(1, 0);
                Console.Write("waiting for " + i.ToString());
                System.Threading.Thread.Sleep(1000);
            }
            Console.Clear();
            Console.Write("Reading and saving the text from uploaded OneNote images");
        }

        static void Main(string[] args)
        {
            //cleanAllUp("a");
            //CreatePage("sectionId", "pageName");
            writeNote();
            waitScreen(Globals.imageCount); // It takes oneNote time to do OCR before we could read it.
            System.Threading.Thread.Sleep(Globals.imageCount * 1000);
            List<string> slideTextList = readNote();
            saveListToFile(slideTextList, "slideText.txt");
        }



        static void cleanAllUp(string a)
        {
            var onenoteApp = new Application();

            string notebookXml;
            onenoteApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            var pageNode = doc.Descendants(ns + "Page").Where(n => n.Attribute("name").Value == "OCRImages").FirstOrDefault();
            var existingPageId = pageNode.Attribute("ID").Value;
            pageNode.RemoveAll();

        }


        static void saveListToFile(List<string> slideTextList, string textFileName)
        {
            for (int i = 0; i < slideTextList.Count(); i++)
            {
                //slideTextList[i] = slideTextList.ElementAt(i).Replace(System.Environment.NewLine, " ");
                //slideTextList[i] = Regex.Replace(slideTextList.ElementAt(i), @"\r\n?|\n", " ");
                slideTextList[i] = slideTextList.ElementAt(i).Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
            }
            System.IO.File.WriteAllLines(Path.Combine(Globals.imgDirPath, textFileName), slideTextList);
        }

        static void writeNote()
        {

            string strNamespace = "http://schemas.microsoft.com/office/onenote/2010/onenote";
            string m_xmlImageContent = "<one:Image><one:Size width=\"{1}\" height=\"{2}\" isSetByUser=\"true\" /><one:Data>{0}</one:Data></one:Image>";
            string m_xmlNewOutline = "<?xml version=\"1.0\"?><one:Page xmlns:one=\"{2}\" ID=\"{1}\"><one:Title><one:OE><one:T><![CDATA[{3}]]></one:T></one:OE></one:Title>{0}</one:Page>";
            string pageName = Globals.oneNoteName;

            var onenoteApp = new Application();

            string notebookXml;
            onenoteApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            var pageNode = doc.Descendants(ns + "Page").Where(n => n.Attribute("name").Value == Globals.oneNoteName).FirstOrDefault();
            var existingPageId = pageNode.Attribute("ID").Value;
            if (pageNode != null)
            {
                string[] fileEntries = Directory.GetFiles(Globals.imgDirPath);
                foreach (string fileName in fileEntries)
                {
                    Bitmap bitmap = new Bitmap(fileName);
                    MemoryStream stream = new MemoryStream();
                    bitmap.Save(stream, ImageFormat.Jpeg);
                    string fileString = Convert.ToBase64String(stream.ToArray());


                    string imageXmlStr = string.Format(m_xmlImageContent, fileString, bitmap.Width / 10, bitmap.Height / 10);
                    string pageChangesXml = string.Format(m_xmlNewOutline, new object[] { imageXmlStr, existingPageId, strNamespace, pageName });
                    onenoteApp.UpdatePageContent(pageChangesXml.ToString(), DateTime.MinValue);
                    Globals.imageNames.Add(fileName);
                    Globals.imageCount++;
                }
            }
            else {
                Console.WriteLine(pageName + " Notebook not found.");
                Environment.Exit(-1);

            }
        }

        static List<string> readNote()
        {
            var onenoteApp = new Application();

            string notebookXml;
            onenoteApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml);

            var doc = XDocument.Parse(notebookXml);
            var ns = doc.Root.Name.Namespace;
            var notebookNode = doc.Descendants(ns + "Page").Where(n => n.Attribute("name").Value == Globals.oneNoteName).FirstOrDefault();

            string pageXml;
            onenoteApp.GetPageContent(notebookNode.Attribute("ID").Value, out pageXml);

            var parsedXML = XDocument.Parse(pageXml);
            List<string> slideTextList = new List<string>();
            string textData;
            for (int i = 0; i < Globals.imageCount; i++)
            {
                ///XElement OCRTXML = parsedXML.Descendants(ns + "OCRText").FirstOrDefault();
                XElement OCRTXML = parsedXML.Descendants(ns + "Image").ElementAt(i).Descendants(ns + "OCRData").FirstOrDefault();
                if (OCRTXML != null) {
                    textData = OCRTXML.Value;
                }
                else {
                    textData = "No text detected";
                }
                slideTextList.Add(Globals.imageNames.ElementAt(i) +": "+ textData);
            }
            return slideTextList;
        }
    }

}

/*
foreach (var notebookNode in from node in doc.Descendants(ns +
"Page")
                             select node)
{
 Console.WriteLine(notebookNode.Attribute("name").Value);
*/

/*if (notebookNode.Attribute("name").Value == "OCRImages")
{ */

/*
static string CreatePage(string sectionId, string pageName)
{
    // Create the new page
    string pageId;
    onenoteApp.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

    // Get the title and set it to our page name
    string xml;
    onenoteApp.GetPageContent(pageId, out xml, PageInfo.piAll);
    var doc = XDocument.Parse(xml);
    var title = doc.Descendants(ns + "T").First();
    title.Value = pageName;

    // Update the page
    onenoteApp.UpdatePageContent(doc.ToString());

    return pageId;
}
*/
