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
        //static Excel.Application excel = new Excel.Application();
        //static Excel.Workbook wbook = excel.Workbooks.Open(@"C:\Users\VIP\Downloads\aaa.xlsx");
        static List<StockData> list = new List<StockData>();

        static void Main(string[] args)
        {
            
            var wbook = excel.Workbooks.Open(@"C:\Users\VIP\Downloads\aaa.xlsx",
                Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing);

            var range = ((Excel.Worksheet)wbook.Worksheets["원본데이터"]).UsedRange;

            for (int i = 2; i <= range.Rows.Count; i++)
            {
                StockList.Add(new StockData(
                    (string)(range.Cells[i, 1] as Excel.Range).Value2,
                    (string)(range.Cells[i, 2] as Excel.Range).Value2,
                    (string)(range.Cells[i, 3] as Excel.Range).Value2,
                    (int)(range.Cells[i, 4] as Excel.Range).Value2,
                    (int)(range.Cells[i, 5] as Excel.Range).Value2,
                    (int)(range.Cells[i, 6] as Excel.Range).Value2,
                    (int)(range.Cells[i, 7] as Excel.Range).Value2,
                    (int)(range.Cells[i, 8] as Excel.Range).Value2,
                    (int)(range.Cells[i, 9] as Excel.Range).Value2,
                    (string)(range.Cells[i, 10] as Excel.Range).Value2));
                    
            }
            List<List<int>> array = new List<List<int>>();
            array.Add(new List<int>());
            array.Add(new List<int>(new int[] { 1, 3, 5 }));

            int aa = 1;
            bool IsWork = false;
            array.ForEach(arr =>
            {

                if (arr.Count > 0 && arr[0] == aa)
                {
                    Console.WriteLine("Done");
                    
                    IsWork = true;
                    return;
                }
            });

            Console.WriteLine(IsWork);
            Console.ReadLine();
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

    static class StockList
    {
        static List<List<StockData>> list = new List<List<StockData>>();

        public static void Add(StockData newdata)
        {
            bool IsAdded = false;
            list.ForEach(array =>
            {
                if (array.Count > 0 && array[0].name_business == newdata.name_business)
                {
                    array.Add(newdata);
                    IsAdded = true;
                    return;
                }
            });

            if (!IsAdded) 
                list.Add(new List<StockData>(new StockData[] { newdata }));
            
        }

        public static StockData[] GetStockData(string BusinessName)
        {
            return (from array in list where array[0].name_business == BusinessName select array).First().ToArray();
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
