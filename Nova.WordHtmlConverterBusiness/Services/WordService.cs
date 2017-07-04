using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nova.WordHtmlConverterBusiness.Services
{
    public class WordService
    {
        public void WordReader()
        {
            //Apertura Conexion Documento Office
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open("C:/Users/lsamaniego/Documents/test.docx", true);
            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            var a = body;


            //Cierre conexion Documento Office
            wordprocessingDocument.Close();
        }
        public void WordTagReplace()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(String.Format("C:/Users/lsamaniego/Documents/Ejemplo Plantilla.docx"), true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                var a = body.Descendants<Text>();
                foreach (var textItem in body.Descendants<Text>().Where(textItem => textItem.Text.Contains("«afCiudad»")))
                {
                    textItem.Text = textItem.Text.Replace("«afCiudad»", "Quito");
                }
                //foreach (var para in body.Elements<Paragraph>())
                //{
                //    foreach (var run in para.Elements<Run>())
                //    {
                //        foreach (var text in run.Elements<Text>())
                //        {
                //            if (text.Text.Contains("«afCiudad»"))
                //                text.Text = text.Text.Replace("«afCiudad»", "Quito");
                //        }
                //    }
                //}
            }
        }
    }
}
