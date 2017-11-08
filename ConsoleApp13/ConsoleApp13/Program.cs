using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp13
{
    class Program
    {
        static Excel.Application excel = new Excel.Application();
        static Excel.Workbook wbook = excel.Workbooks.Open(@"C:\Users\VIP\Downloads\aaa.xlsx");

        static void Main(string[] args)
        {

            var wbook = excel.Workbooks.Open(@"C:\Users\VIP\Downloads\aaa.xlsx",
                Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing);

            var worksheets = wbook.Worksheets;
            



            SearchData((Excel.Worksheet)worksheets["원본데이터"]);
        }

        static void SearchData(Excel.Worksheet sheet)
        {
            Console.Write("Enter data : ");
            string index1 = Console.ReadLine();
            Console.Write("Enter data : ");
            string index2 = Console.ReadLine();

            Excel.Range range = sheet.get_Range(index1, index2);
            for (int i = 2; i <= range.Rows.Count; i++)
            {
                for (int j = 1; j <= range.Columns.Count; j++)
                {
                    Console.WriteLine((range.Cells[i, j] as Excel.Range).Value2);
                }
                Console.WriteLine("====");
            }

            SearchData(sheet);
        }
    }

    class StockData

    {
        public string code_event { get; set; }
        public string name_event { get; set; }
        public string date { get; set; }
        public string name_business { get; set; }
        public int val_start { get; set; }
        public int val_max { get; set; }
        public int val_min { get; set; }
        public int val_end { get; set; }
        public int val_total { get; set; }
        public int countstock { get; set; }

        public StockData(string EventName, string EventCode, string Date, 
            int StartPS, int MaxPS, int MinPS, 
            int EndPS, int CountStock, int Total, string BusinessName)
        {
            name_event = EventName;
            code_event = EventCode;
            date = Date;
            val_start = StartPS;
            val_max = MaxPS;
            val_min = MinPS;
            val_end = EndPS;
            countstock = CountStock;
            val_total = Total;
            name_business = BusinessName;

        }
    }
}
