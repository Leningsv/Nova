
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Shark.PdfConvert;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Nova.WordHtmlConverterBusiness.Services
{
    public class DocumentConverterService
    {
        public bool WordToHTML()
        {
            bool res= false;
            try
            {
                byte[] byteArray = File.ReadAllBytes("C:/Users/lsamaniego/Documents/Sprint Backlog.docx");
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    memoryStream.Write(byteArray, 0, byteArray.Length);
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
                    {                        
                        HtmlConverterSettings settings = new HtmlConverterSettings()
                        {
                            PageTitle = "My Page Title"
                        };
                        XElement html = HtmlConverter.ConvertToHtml(doc, settings);

                        File.WriteAllText("C:/Users/lsamaniego/Documents/Sprint Backlog.html", html.ToStringNewLineOnAttributes());
                    }
                }
                res = true;
            }
            catch (Exception ex)
            {
                throw;
            }
            return res;
        }
        public bool HTMLToPDF()
        {

            bool res = false;
            try
            {
                PdfConvert.Convert(new PdfConversionSettings
                {
                    Title = "My Static Content from URL",
                    ContentUrl = "C:/Users/lsamaniego/Documents/Sprint Backlog.html",
                    OutputPath = @"C:/Users/lsamaniego/Documents/aaa.pdf"
                });               

            }
            catch (Exception ex)
            {

                throw;
            }
            return res;            
        }
    }
}
