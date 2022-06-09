using Exceldatascript;

namespace Mstestdatascipt
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethodPoc()
        {
            Assert.AreEqual(2 ,2 );
        }

        [TestMethod]
        public void TestLenghtOfPdfDokuments()
        {
            ExcelDataScriptExecute excelDataScriptExecute = new ExcelDataScriptExecute(); 

            var listoutput = excelDataScriptExecute.GetDataTableFromExcel(38);

            Assert.AreEqual(listoutput.Count, 7689);
        }
    }
}