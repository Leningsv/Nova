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
    public class WordServiceTests
    {
        private WordService wordService = new WordService();
        [TestMethod()]
        public void WordReaderTest()
        {
            try
            {
                wordService.WordReader();
                Assert.IsTrue(true);
            }
            catch (Exception)
            {
                Assert.Fail();
            }
        }

        [TestMethod()]
        public void WordTagReplaceTest()
        {
            try
            {
                wordService.WordTagReplace();
                Assert.IsTrue(true);
            }
            catch (Exception)
            {
                Assert.Fail();
            }
        }
    }
}