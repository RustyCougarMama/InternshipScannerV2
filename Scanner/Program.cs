using System;
using System.Collections;
using System.Collections.Generic;
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
            Console.WriteLine("Print all Files?");
            Console.ReadLine();
            s.GetAllExcelFiles();
            foreach (string name in s.DataFileNames)
            {
                Console.WriteLine(name);

            }
            s.ProcessExcelFiles();
            foreach (var education in s.IntStudentsDictionary.Keys)
            {
                Console.WriteLine("");
                Console.WriteLine(education);
                Console.WriteLine("##############################################################");
                IDictionary students = (IDictionary)s.IntStudentsDictionary[education];
                foreach (var student in students.Keys)
                {

                    Console.WriteLine(student);
                }
                //Console.WriteLine(dic);
            }
            Console.WriteLine("You are now done...");
            Console.ReadKey();

        }
    }
    public class Student
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string JobTitle { get; set; }
        public Student(string _name, string _email, string _jobTitle)
        {
            Name = _name;
            Email = _email;
            JobTitle = _jobTitle;
        }
        public Student()
        {

        }
    }

    class Scanner
    {
        public string InternationalStudentsExcelFilePath { get; set; }
        public string DataFilePath { get; set; }
        public string ResultsFolder { get; set; }
        public ArrayList InternationalStudentEmails { get; set; }
        public ArrayList DataFileNames { get; set; }
        public int DataRowStart { get; set; }
        public int IntStudentsRowStart { get; set; }
        public int IntStudentsColumnStart { get; set; }
        public IDictionary IntStudentsDictionary { get; set; }



        public Scanner()
        {
            InternationalStudentsExcelFilePath = @"D:\Google Drive\UCN Work\InternshipScannerV2\InternationalStudentsList\International students - B&T.xlsx";
            DataFilePath = @"D:\Google Drive\UCN Work\InternshipScannerV2\Data";
            DataFileNames = new ArrayList();
            DataRowStart = 2;
            InternationalStudentEmails = new ArrayList();
            IntStudentsRowStart = 2;
            IntStudentsColumnStart = 1;
            IntStudentsDictionary = new Dictionary<string, Dictionary<string, Student>>();
        }

        public string CollectInternationalEmails()
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
                        for (int row = IntStudentsRowStart; row <= end.Row; row++)
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
                    Console.WriteLine(p.File.Name);
                }
                string s = $"All international emails collected... /n There are {InternationalStudentEmails.Count} international students on record...";
                return s;
            }

            catch (PathTooLongException e)
            {
                Console.WriteLine("The path is too long; Steven did not consider this... Try again with another folder that's not so deeply barried in all of your junk, maybe?");
                return "The path is too long; Steven did not consider this... Try again with another folder that's not so deeply barried in all of your junk, maybe?";
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                return e.Message;
                throw e;
            }

        }

        public void GetAllExcelFiles()
        {
            DataFileNames.AddRange(Directory.GetFiles(DataFilePath));
        }

        public string ProcessExcelFiles()
        {
            if (DataFileNames.Count > 0)
            {
                foreach (string fileString in DataFileNames)
                {
                    var fileInfo = new FileInfo(fileString);
                    using (var p = new ExcelPackage(fileInfo))
                    {
                        ExcelWorkbook wb = p.Workbook;
                        ExcelWorksheet ws = wb.Worksheets[0];
                        IDictionary Education = new Dictionary<string, Student>();
                        var end = ws.Dimension.End;
                        for (int row = DataRowStart; row <= end.Row; row++)
                        {
                            Student s = new Student
                            {
                                Name = ws.Cells[row, 3].Value.ToString(),
                                Email = ws.Cells[row, 4].Value.ToString(),
                                JobTitle = ws.Cells[row, 2].Value.ToString()

                            };
                            //TODO Check is student is international
                            if (!Education.Contains(s.Name))
                            {
                                Education.Add(s.Name, s);
                            }
                        }
                        IntStudentsDictionary.Add(p.File.Name, Education);
                    }
                }
            }
            return "All students in the selected educations added to collection...";
        }
    }
}

