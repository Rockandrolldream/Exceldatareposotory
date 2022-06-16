using Exceldatascript;
using Moq;

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
        public void ShouldTestPropertyInInterfaces()
        {
            Mock<ExcelInterface> excelmock = new Mock<ExcelInterface>();
            excelmock.SetupProperty(p => p.Pdf_URL, "http://arpeissig.at/wp-content/uploads/2016/02/D7_NHB_ARP_Final_2.pdf"); 
            var obj = excelmock.Object;
            Assert.IsNotNull(obj.Pdf_URL);
        }

        [TestMethod]
        public void ShouldTestExcelObjectConstructorProperty()
        {
            var constructorobject = new Mock<ExcelObject>(MockBehavior.Strict, new object[] { "http://arpeissig.at/wp-content/uploads/2016/02/D7_NHB_ARP_Final_2.pdf", "IsDownloaded", 20, "BR50060" });
            var constructorobjectvalue = constructorobject.Object.Rownumber; 
            Assert.AreEqual(constructorobjectvalue, 20);
        }
    }
}