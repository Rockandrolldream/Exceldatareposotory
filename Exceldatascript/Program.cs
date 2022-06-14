using System.Data;

namespace Exceldatascript
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelDataScriptExecute exceldatascriptexecute = new ExcelDataScriptExecute();
            // var result = exceldatascriptexecute.GetDataTableFromExcel(38); 


            // opgave 1
            //var listforMetadata2006_2016 = exceldatascriptexecute.Metadata2006_2016();
            //foreach (var item in listforMetadata2006_2016)
            //{
            //    Console.WriteLine(item.Pdf_URL + "  " + item.Isdownloaded);
            //}


            // opgave 2 


            var listresults = exceldatascriptexecute.GetDataFromExcel(38);

            //foreach (var item in listresults)
            //{
            //    Console.WriteLine(item.BRnum + " " + item.Pdf_URL);
            //}

            //exceldatascriptexecute.LookIntoAnotherColoumn(listresults);

            exceldatascriptexecute.Downloadfiles(listresults);


        }
    }
}