using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AsposeUnitTestProject
{
    [TestClass]
    public class Util_PDF_Aspose_UnitTest
    {
        [TestInitialize]
        public void Init()
        {
            Util.PDF.PDFUtils_Aspose.InitializeAsposePDF();
        }

        [TestMethod]
        public void TestCreate()
        {
            Util.PDF.PDFUtils_Aspose.TestCreate();
        }

        [TestMethod]
        public void TestReadDocInfo()
        {
            var r = Util.PDF.PDFUtils_Aspose.TestReadDocInfo();
            Assert.AreEqual<string>("main text", r);
        }

        [TestMethod]
        public void TestReadText()
        {
            var r = Util.PDF.PDFUtils_Aspose.TestRead();
            Assert.AreEqual<string>("main text", r);
        }

        [TestMethod]
        public void TestReadGAC()
        {
            string pdfPath = @"D:\SC_Github\TestAspose\AsposeUnitTestProject\PDF\TestReadGAC.pdf";
            var r = Util.PDF.PDFUtils_Aspose.TestReadGAC(pdfPath);
        }

        [TestMethod]
        public void TestReadGACV2()
        {
            string pdfPath = @"D:\SC_Github\TestAspose\AsposeUnitTestProject\PDF\TestReadGAC.pdf";
            var r = Util.PDF.PDFUtils_Aspose.TestReadGAC_BAK(pdfPath);
        }

        [TestMethod]
        public void TestReadGACV3()
        {
            string pdfPath = @"D:\SC_Github\TestAspose\AsposeUnitTestProject\PDF\TestReadGAC.pdf";
            var r = Util.PDF.PDFUtils_Aspose.TestRead();
        }
    }
}
