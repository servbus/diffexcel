using OfficeOpenXml;
using System;
using System.IO;

namespace DiffExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //args = new[] { @"C:\Users\cnryb\Desktop\diffexcel\test\t1.xlsx", @"C:\Users\cnryb\Desktop\diffexcel\test\t2.xlsx" };

            if (args.Length == 0 || args[0] == "-h" || args[0] == "--help")
            {
                Console.WriteLine("Compare the differences between two excel files");
                return;
            }
            var filename1 = args[0];
            var filename2 = args[1];
            ExcelPackage package1 = new ExcelPackage();
            ExcelPackage package2 = new ExcelPackage();
            package1.Load(File.OpenRead(filename1));
            package2.Load(File.OpenRead(filename2));

            if (package1.Workbook.Worksheets.Count != package2.Workbook.Worksheets.Count)
            {
                Different($"Worksheets count");
                return;
            }

            for (int i = 0; i < package1.Workbook.Worksheets.Count; i++)
            {
                var sheet1 = package1.Workbook.Worksheets[i];
                var sheet2 = package2.Workbook.Worksheets[i];
                if (sheet1.Name != sheet2.Name)
                {
                    Different($"sheet name. index {i}");
                    return;
                }
                if (sheet1.Dimension.End.Row != sheet2.Dimension.End.Row)
                {
                    Different($"row number. sheet name {sheet1.Name}, index {i}");
                    return;
                }
                if (sheet1.Dimension.End.Column != sheet2.Dimension.End.Column)
                {
                    Different($"column number. sheet name {sheet1.Name}, index {i}");
                    return;
                }

                for (int m = 1; m <= sheet1.Dimension.End.Row; m++)
                {
                    for (int n = 1; n <= sheet1.Dimension.End.Column; n++)
                    {
                        var v1 = sheet1.Cells[m, n];
                        var v2 = sheet2.Cells[m, n];
                        if (v1.Value == null)
                        {
                            if (v2.Value != null)
                            {
                                Different($"cell value. address {v1.Address}");
                                return;
                            }
                        }
                        else if (v1.Value.Equals(v2.Value) == false)
                        {
                            Different($"cell value. address {v1.Address}");
                            return;
                        }
                    }
                }
            }

            Console.WriteLine("same");
        }

        static void Different(string msg)
        {
            Console.WriteLine($"different: {msg}");
        }
    }
}
