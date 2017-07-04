using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nova.WordHtmlConverterBusiness.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nova.WordHtmlConverterBusiness.Services.Tests
{
    [TestClass()]
    public class ConverterServiceTests
    {
        ConverterService converterService = new ConverterService();
        [TestMethod()]
        public void ConvertWordToPDFTest()
        {
            try
            {
                converterService.ConvertWordToPDF();
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail();
            }            
        }
    }
}