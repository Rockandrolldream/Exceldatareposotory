using System.Data;

namespace Exceldatascript
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = "C:/Users/KOM/Desktop/Exceldatascriptopgave/GRI_2017_2020.xlsx";
            ExcelDataScriptExecute exceldatascriptexecute = new ExcelDataScriptExecute();
            exceldatascriptexecute.GetDataTableFromExcel(38);

        }
    }
}