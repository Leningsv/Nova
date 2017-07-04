using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nova.WordHtmlConverterBusiness.Services
{
    public class ConverterService
    {
        public void ConvertWordToPDF()
        {
            Document wordDocument;            
            try
            {
                Application appWord = new Application();
                wordDocument = appWord.Documents.Open("C:/Users/lsamaniego/Documents/EjemploPlantilla.docx");
                wordDocument.ExportAsFixedFormat("C:/Users/lsamaniego/Documents/EjemploPlantilla.pdf", WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception ex)
            {

                throw;
            }
        }
    }
}
