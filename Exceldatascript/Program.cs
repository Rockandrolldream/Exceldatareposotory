using System.Data;

namespace Exceldatascript
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelDataScriptExecute exceldatascriptexecute = new ExcelDataScriptExecute();

            // opgave 1
            exceldatascriptexecute.UpdateMetaData();

            // opgave 2
            exceldatascriptexecute.GetDataFromExcel();           
        }
    }
}