using System;
using System.Collections;
using System.IO;
using OfficeOpenXml;

namespace Scanner
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Are you ready?");
            Console.ReadLine();
            Scanner s = new Scanner();
            s.CollectInternationalEmails();
            Console.ReadLine();
        }
    }

    class Scanner
    {
        public string InternationalStudentsExcelFilePath { get; set; }
        public string ResultsFolder { get; set; }

        public ArrayList InternationalStudentEmails { get; set; }
        public int IntStudentsRowStart { get; set; }
        public int IntStudentsColumnStart { get; set; }

        public Scanner()
        {
            InternationalStudentsExcelFilePath = @"D:\Google Drive\UCN Work\InternshipScannerV2\InternationalStudentsList\International students - B&T.xlsx";
            InternationalStudentEmails = new ArrayList();
            IntStudentsRowStart = 2;
            IntStudentsColumnStart = 1;
        }

        public void CollectInternationalEmails()
        {
            Console.WriteLine("Collecting emails of all registered International Students");
            try
            {
                var fileInfo = new FileInfo(InternationalStudentsExcelFilePath);
                using (var p = new ExcelPackage(fileInfo))
                {
                    ExcelWorkbook wb = p.Workbook;
                    ExcelWorksheets workSheets = wb.Worksheets;
                    foreach (ExcelWorksheet worksheet in workSheets)
                    {
                        var end = worksheet.Dimension.End;
                        for(int row = IntStudentsRowStart; row <= end.Row; row++)
                        {
                            var data = worksheet.Cells[row, IntStudentsColumnStart].Value;
                            if (data == null)
                            {
                                continue;
                            }
                            string studentemail = data.ToString().Trim();
                            InternationalStudentEmails.Add(studentemail);
                            Console.WriteLine(studentemail);
                        }
                    }
                    Console.WriteLine();
                    Console.WriteLine("All international emails collected...");
                    Console.WriteLine($"There are {InternationalStudentEmails.Count} international students on record...");
                }
            }

            catch (PathTooLongException e)
            {
                Console.WriteLine("The path is too long; Steven did not consider this... Try again with another folder that's not so deeply barried in all of your junk, maybe?");
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                throw e;
            }
            
        }

    }
}
