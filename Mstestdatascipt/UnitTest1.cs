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
        public void TestMetadata2006_2016()
        {  
            
            ExcelDataScriptExecute excelDataScriptExecute = new ExcelDataScriptExecute();  

            var listoutput = excelDataScriptExecute.Metadata2006_2016();

            Assert.IsNotNull(listoutput.Count);
        }
    }
}