using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Nova.WordHtmlConverterBusiness.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
namespace Nova.WordHtmlConverterBusiness.Services
{
    public class FormularyService
    {
        private void ValidateDirectory(string directory)
        {
            try
            {
                //Verifica la Existencia del Directorio
                if (Directory.Exists(directory))
                {
                    Console.WriteLine("That path exists already.");
                }
                else
                {
                    DirectoryInfo di = Directory.CreateDirectory(directory);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        private bool GenerateTempFile(AuxFile auxFile)
        {
            bool res = false;
            try
            {
                Byte[] bytes = Convert.FromBase64String(auxFile.Base64);
                File.WriteAllBytes(auxFile.FullTempPath, bytes);
                res = true;
            }
            catch (Exception ex)
            {
                throw;
            }
            return res;
        }
        /// <summary>
        /// Librerias Nesesarias
        /// using DocumentFormat.OpenXml.Packaging;
        /// using DocumentFormat.OpenXml.Wordprocessing;
        /// </summary>
        /// <param name="fullPathTempFile">Path del directorio del documento a convertir</param>
        /// <param name="listKeyWord">Palabras claves a remplazar en el documento</param>
        /// <returns></returns>
        public bool ProcessFile(string fullPathTempFile, List<KeyWordFile> listKeyWord)
        {
            bool res = false;
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPathTempFile, true))
                {
                    var body = doc.MainDocumentPart.Document.Body;
                    foreach (var item in listKeyWord)
                    {
                        foreach (var textItem in body.Descendants<Text>().Where(textItem => textItem.Text.Contains(item.Code)))
                        {
                            textItem.Text = textItem.Text.Replace(item.Code, item.Value);
                        }
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
        public void GeretatePDF(AuxFile auxFile, List<KeyWordFile> listKeyWord)
        {
            string res = string.Empty;
            try
            {
                //Validacion del Directorio
                ValidateDirectory(ConfigurationManager.AppSettings["DirectoryTempFile"]);
                auxFile.FullTempPath = ConfigurationManager.AppSettings["DirectoryTempFile"] + auxFile.Name + DateTime.Now.ToString("yyyyMMddhhmmss") + "." + auxFile.Extention;
                //Creacion del Archivo Temporal
                if (GenerateTempFile(auxFile))
                {
                    //Se Procesa el archivo Temporal 
                    if (ProcessFile(auxFile.FullTempPath, listKeyWord))
                    {
                        //Se Genera el archivo PDF
                        ValidateDirectory(ConfigurationManager.AppSettings["DirectoryPDFFile"]);
                        //Generacion Archivo PDF
                        Microsoft.Office.Interop.Word.Document wordDocument;
                        Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                        wordDocument = appWord.Documents.Open(auxFile.FullTempPath);
                        var auxRes = ConfigurationManager.AppSettings["DirectoryPDFFile"] + auxFile + DateTime.Now.ToString("yyyyMMddhhmmss") + ".pdf";
                        wordDocument.ExportAsFixedFormat(auxRes, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                        //appWord.Documents.Close();
                        wordDocument.Close();
                        res = auxRes;
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
