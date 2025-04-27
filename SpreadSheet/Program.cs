using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheet
{
    public class Program
    {
        static void Main(string[] args)
        {
            Data data = new Data();
            Create.CreateSpreadsheetWorkbook("E:\\Spreadsheet\\Create.xlsx");
            Create.InsertData("E:\\Spreadsheet\\Create.xlsx", data.myData, "mySheet");
            Create.ReadExcelFileDOM("E:\\Spreadsheet\\Create.xlsx");
        }
    }
}
