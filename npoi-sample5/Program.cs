using NPOI.SS.UserModel;
using System;
using System.IO;


namespace WorkbookFactoryDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var stream= File.OpenRead("fakeExcel.xls"))
            {
                IWorkbook workbook = WorkbookFactory.Create(stream);

                Console.WriteLine("workbook type: " + workbook.GetType().Name);
                Console.ReadLine();
            }
        }
    }
}
