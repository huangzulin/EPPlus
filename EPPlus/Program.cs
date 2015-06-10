using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace EPPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            IList<int> list= new List<int>();
            var package = new ExcelPackage(new FileInfo(@"C:\workspace\EPPlus\EPPlus\test.xlsx"));
            ExcelWorksheet workSheet = package.Workbook.Worksheets[1];
            for (int i = 1; i < 50; i++)
            {
                object cellValue = workSheet.Cells[i, 1].Value;
                list.Add(cellValue==null?0:Convert.ToInt32(cellValue));
            }

            Console.WriteLine(list.Count);
            list = list.Distinct().ToList();
            foreach (var i in list)
            {
                Console.WriteLine(i);
            }
        }
    }
}
