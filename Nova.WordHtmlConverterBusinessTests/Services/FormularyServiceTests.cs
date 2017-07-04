using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nova.WordHtmlConverterBusiness.Models;
using System;
using System.Collections.Generic;
using System.IO;

namespace Nova.WordHtmlConverterBusiness.Services.Tests
{
    [TestClass()]
    public class FormularyServiceTests
    {
        FormularyService formularyService = new FormularyService();
        [TestMethod()]
        public void ProcessFileTest()
        {
            try
            {
                var listKeyWord = new List<KeyWordFile>() {
                    new KeyWordFile()
                    {
                        Code="«afDia»",
                        Value="01"
                    },new KeyWordFile()
                    {
                        Code="«afMes»",
                        Value="Octubre"
                    },new KeyWordFile()
                    {
                        Code="«afAnio»",
                        Value="17"
                    }
                };
                formularyService.ProcessFile("C:/Users/lsamaniego/Documents/EjemploPlantilla.docx", listKeyWord);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail();
            }
        }

        [TestMethod()]
        public void GeretatePDFTest()
        {
            try
            {
                byte[] bytes = File.ReadAllBytes("C:/Users/lsamaniego/Documents/EjemploPlantilla.docx");
                string file64 = Convert.ToBase64String(bytes);
                AuxFile auxFile = new AuxFile()
                {
                    Base64 = file64,
                    Extention = "docx",
                    FullTempPath = "",
                    Name = "EjemploPlantilla"
                };
                var listKeyWord = new List<KeyWordFile>() {
                    new KeyWordFile()
                    {
                        Code="«afDia»",
                        Value="01"
                    },new KeyWordFile()
                    {
                        Code="«afMes»",
                        Value="Octubre"
                    },new KeyWordFile()
                    {
                        Code="«afAnio»",
                        Value="17"
                    }
                };
                formularyService.GeretatePDF(auxFile, listKeyWord);
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail();
            }
        }
    }
}