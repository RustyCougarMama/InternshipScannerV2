using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace InternshipScannerV2.Controllers
{
    //class Program
    //{
    //    static void Main(string[] args)
    //    {
    //        Console.WriteLine("Are you ready?");
    //        Console.ReadLine();
    //        Scanner s = new Scanner();
    //        s.CollectInternationalEmails();
    //        Console.WriteLine("Print all Files?");
    //        Console.ReadLine();
    //        s.GetAllExcelFiles();
    //        foreach (string name in s.DataFileNames)
    //        {
    //            Console.WriteLine(name);
    //        }
    //        s.ProcessExcelFiles();
    //        foreach (var education in s.IntStudentsDictionary.Keys)
    //        {
    //            Console.WriteLine("");
    //            Console.WriteLine(education);
    //            Console.WriteLine("##############################################################");
    //            IDictionary students = (IDictionary)s.IntStudentsDictionary[education];
    //            foreach (var student in students.Keys)
    //            {
    //                Console.WriteLine(student);
    //            }
    //            //Console.WriteLine(dic);
    //        }
    //        Console.WriteLine("You are now done...");
    //        Console.ReadKey();
    //    }
    //}
    public class Student
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string JobTitle { get; set; }
        public bool isStudyingiInDK { get; set; }
        public Student(string _name, string _email, string _jobTitle)
        {
            Name = _name;
            Email = _email;
            JobTitle = _jobTitle;
            isStudyingiInDK = false;
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
        public string ResultName { get; set; }
        public ArrayList InternationalStudentEmails { get; set; }
        public ArrayList DataFileNames { get; set; }
        public int DataRowStart { get; set; }
        public int IntStudentsRowStart { get; set; }
        public int IntStudentsColumnStart { get; set; }
        public IDictionary IntStudentsDictionary { get; set; }
        private List<string> listOfEducations;
        private FileInfo curResultFileInfo;
        private ArrayList curStudentEmails;
        private ArrayList studentsInCurEducation;
        private int curStudentIndex;
        private int curEducationIndex;
        private int curEducationSize;
        private int numberOfEducations;
        public TextBox tb;



        public Scanner()
        {
            InternationalStudentsExcelFilePath = @"D:\Google Drive\UCN Work\InternshipScannerV2\InternationalStudentsList\International students - B&T.xlsx";
            DataFilePath = @"D:\Google Drive\UCN Work\InternshipScannerV2\Data";
            ResultsFolder = @"D:\Google Drive\UCN Work\InternshipScannerV2\Result\";
            DataFileNames = new ArrayList();
            DataRowStart = 2;
            InternationalStudentEmails = new ArrayList();
            IntStudentsRowStart = 2;
            IntStudentsColumnStart = 1;
            IntStudentsDictionary = new Dictionary<string, Dictionary<string, Student>>();
            curStudentEmails = new ArrayList();
            curStudentIndex = 0;
            curEducationIndex = 0;
            curEducationSize = 0;
            numberOfEducations = 0;
            ResultName = "result1.xlsx";
            listOfEducations = new List<string>();
        }

        public Scanner(string resultName, string resultFolderFilePath, string intStudentExcelFilePath, string dataFilePath, TextBox textBox)
        {
            InternationalStudentsExcelFilePath = intStudentExcelFilePath;
            DataFilePath = dataFilePath;
            ResultsFolder = resultFolderFilePath;
            ResultName = resultName;
            DataFileNames = new ArrayList();
            DataRowStart = 2;
            InternationalStudentEmails = new ArrayList();
            IntStudentsRowStart = 2;
            IntStudentsColumnStart = 1;
            IntStudentsDictionary = new Dictionary<string, Dictionary<string, Student>>();
            listOfEducations = new List<string>();
            tb = textBox;

        }

        #region Initialization


        private void appendTextBox(string text)
        {
            if (tb != null)
            {
                tb.AppendText(text);
                tb.AppendText(Environment.NewLine);
                //tb.SelectionStart = tb.Text.Length;
                tb.ScrollToEnd();
            }

        }

        public int CollectInternationalEmails()
        {
            Console.WriteLine("Collecting emails of all registered International Students");
            appendTextBox("Collecting emails of all registered International Students");
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
                            appendTextBox(studentemail);
                        }
                    }
                    Console.WriteLine();
                    appendTextBox(" ");
                    Console.WriteLine("All international emails collected...");
                    appendTextBox("All international emails collected...");
                    string text = $"There are {InternationalStudentEmails.Count} international students on record...";
                    Console.WriteLine(text);
                    appendTextBox(text);
                }
                int s = InternationalStudentEmails.Count;
                return s;
            }

            catch (PathTooLongException e)
            {
                appendTextBox("The path is too long; Steven did not consider this... Try again with another folder that's not so deeply barried in all of your junk, maybe?");
                Console.WriteLine("The path is too long; Steven did not consider this... Try again with another folder that's not so deeply barried in all of your junk, maybe?");
                return -2;
            }

            catch (Exception e)
            {
                appendTextBox(e.Message);
                Console.WriteLine(e.Message);
                appendTextBox(e.StackTrace);
                Console.WriteLine(e.StackTrace);
                return -1;
                throw e;
            }

        }

        public void GetAllExcelFiles()
        {
            DataFileNames.AddRange(Directory.GetFiles(DataFilePath));
        }

        public int ProcessExcelFiles()
        {
            appendTextBox("  ");
            appendTextBox("Collecting all educations and students assigned...");
            int noInternationStudents = 0;
            if (DataFileNames.Count > 0)
            {
                foreach (string fileString in DataFileNames)
                {
                    var fileInfo = new FileInfo(fileString);
                    appendTextBox(fileString);
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
                            if (InternationalStudentEmails.Contains(s.Email))
                            {
                                if (!Education.Contains(s.Name))
                                {
                                    noInternationStudents++;
                                    Education.Add(s.Name, s);
                                }
                            }
                            else
                            {
                                appendTextBox(s.Name + " is not an international.");
                            }

                        }
                        IntStudentsDictionary.Add(p.File.Name, Education);
                        listOfEducations.Add(p.File.Name);
                    }
                }
            }

            return noInternationStudents;
        }
        /// <summary>
        /// Creates the starting excel workbook to be used for the results.
        /// </summary>
        //private void CreateResultExcel()
        //{
        //    string[] files = Directory.GetFiles(ResultsFolder);
        //    if (files.Length == 0)
        //    {
        //        using (ExcelPackage ex = new ExcelPackage())
        //        {
        //            ex.Workbook.Properties.Title = ResultName;
        //            curResultFileInfo = new FileInfo(ResultsFolder + ResultName + ".xlsx");
        //            ex.Workbook.Worksheets.Add("config");
        //            ex.SaveAs(curResultFileInfo);
        //        }
        //        return;
        //    }
        //    else
        //    {
        //        using (ExcelPackage ex = new ExcelPackage())
        //        {
        //            ResultName = ResultName + (files.Length + 1);
        //            ex.Workbook.Properties.Title = ResultName;
        //            curResultFileInfo = new FileInfo(ResultsFolder + ResultName + ".xlsx");
        //            ex.SaveAs(curResultFileInfo);
        //        }
        //        return;
        //    }
        //}

        #endregion
        /// <summary>
        /// DEPRICATED METHOD, DO NOT USE.
        /// </summary>
        /// <param name="lblStudentName"></param>
        /// <param name="lblEmail"></param>
        /// <param name="lblWorkplace"></param>
        /// <param name="lblEducation"></param>
        public void ScanStudents(TextBlock lblStudentName, TextBlock lblEmail, TextBlock lblWorkplace, TextBlock lblEducation)
        {
            appendTextBox("Starting Scan...");
            foreach (string education in IntStudentsDictionary.Keys)
            {
                appendTextBox("###########################################");
                appendTextBox(education);
                appendTextBox("###########################################");
                using (ExcelPackage ex = new ExcelPackage())
                {
                    string educationName = education;
                    educationName = educationName.Substring(0, educationName.Length - 5);
                    ex.Workbook.Worksheets.Add(educationName);
                    IDictionary currentEducation = (IDictionary)IntStudentsDictionary[education];
                    lblEducation.Text = educationName;
                    foreach (Student s in currentEducation.Values)
                    {
                        //GOT THIS FAR, SHE WORKS!!!
                        string name = s.Name;
                        string email = s.Email;
                        string workPlace = s.JobTitle;

                        lblStudentName.Text = name;
                        lblEmail.Text = email;
                        lblWorkplace.Text = workPlace;
                        //wait for user to make decision.

                        //break the loop instead, and wait for button press.
                    }
                }
            }
        }
        /// <summary>
        /// Takes the data from the current student in the dictionary, and displays their information.
        /// Also ensures that once the scanner reaches the end of an education, it will move onto the 
        /// next one.
        /// </summary>
        /// <param name="lblStudentName">ui element</param>
        /// <param name="lblEmail">ui element</param>
        /// <param name="lblWorkplace">ui element</param>
        /// <param name="lblEducation">ui element</param>
        public bool StepScan(TextBlock lblStudentName, TextBlock lblEmail, TextBlock lblWorkplace, TextBlock lblEducation)
        {
            //TODO Add a check to see when to finish scan

            if (listOfEducations.Count == curEducationIndex)
            {
                //TODO Create and run the convert to excel method
                return false;
            }


            string curEducation = listOfEducations[curEducationIndex];
            lblEducation.Text = curEducation;
            IDictionary currentEducation = (IDictionary)IntStudentsDictionary[curEducation];

            // Here we check to see if we have reached the end of the current education. If so, it increases
            // the index of the education, and resets the curStudentIndex to 0

            if (studentsInCurEducation == null)
            {
                studentsInCurEducation = new ArrayList(currentEducation.Keys);
                appendTextBox(curEducation);
            }

            if (curStudentIndex == studentsInCurEducation.Count)
            {
                curEducationIndex++;
                if (curEducationIndex == listOfEducations.Count)
                {
                    return false;
                }
                curEducation = listOfEducations[curEducationIndex];
                currentEducation = (IDictionary)IntStudentsDictionary[curEducation];
                curStudentIndex = 0;
                studentsInCurEducation = new ArrayList(currentEducation.Keys);
                appendTextBox(curEducation);
                lblEducation.Text = curEducation;
            }

            appendTextBox("Number of Students in this education: " + studentsInCurEducation.Count);

            Student s = (Student)currentEducation[studentsInCurEducation[curStudentIndex]];

            string name = s.Name;
            string email = s.Email;
            string workPlace = s.JobTitle;

            lblStudentName.Text = name;
            lblEmail.Text = email;
            lblWorkplace.Text = workPlace;

            return true;
        }
        /// <summary>
        /// The action that occurs when the user click "Approve"
        /// Increases the curStudentIndex by 1
        /// </summary>
        public void approveStudent()
        {
            string curEducation = listOfEducations[curEducationIndex];
            IDictionary currentEducation = (IDictionary)IntStudentsDictionary[curEducation];
            Student s = (Student)currentEducation[studentsInCurEducation[curStudentIndex]];

            s.isStudyingiInDK = true;
            currentEducation[s.Name] = s;
            IntStudentsDictionary[curEducation] = currentEducation;
            appendTextBox("-Approved!");
            curStudentIndex++;
        }

        /// <summary>
        /// The action that occurs when the user click "Reject"
        /// Increases the curStudentIndex by 1
        /// </summary>
        public void rejectStudent()
        {
            string curEducation = listOfEducations[curEducationIndex];
            IDictionary currentEducation = (IDictionary)IntStudentsDictionary[curEducation];
            Student s = (Student)currentEducation[studentsInCurEducation[curStudentIndex]];

            s.isStudyingiInDK = false;
            currentEducation[s.Name] = s;
            IntStudentsDictionary[curEducation] = currentEducation;
            appendTextBox("-Denied!");
            curStudentIndex++;
        }

        public void saveEducationResults()
        {
            ArrayList educations = new ArrayList(IntStudentsDictionary.Keys);
            using (ExcelPackage ex = new ExcelPackage(curResultFileInfo))
            {

                string[] files = Directory.GetFiles(ResultsFolder);
                if (files.Length == 0)
                {
                    ex.Workbook.Properties.Title = ResultName;
                    curResultFileInfo = new FileInfo(ResultsFolder + ResultName + ".xlsx");
                }
                else
                {

                    ResultName = ResultName.Substring(0, 6) + (files.Length + 1) + ".xlsx";
                    ex.Workbook.Properties.Title = ResultName;
                    curResultFileInfo = new FileInfo(ResultsFolder + ResultName );

                }

                foreach (string education in educations)
                {
                    IDictionary eDic = (IDictionary)IntStudentsDictionary[education];
                    ExcelWorksheet ws = ex.Workbook.Worksheets.Add(education);
                    //ExcelWorksheet ws =
                    //    ex.Workbook.Worksheets.FirstOrDefault(x => x.Name == education);
                    //if (ws == null)
                    //{
                    //    ws = ex.Workbook.Worksheets.Add(education);
                    //}

                    var cellEdu = ws.Cells[1, 1];
                    cellEdu.IsRichText = true;
                    var Edu = cellEdu.RichText.Add(education);
                    Edu.Bold = true;
                    Edu.Size = 24;
                    cellEdu.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cellEdu.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                    var cellStudent = ws.Cells[3, 1];
                    cellStudent.IsRichText = true;
                    var Stu = cellStudent.RichText.Add("Student");

                    var cellEmail = ws.Cells[3, 2];
                    cellEmail.IsRichText = true;
                    var Ema = cellEmail.RichText.Add("Email");

                    var cellWorkplace = ws.Cells[3, 3];
                    cellWorkplace.IsRichText = true;
                    var wor = cellWorkplace.RichText.Add("Workplace");

                    var cellInDK = ws.Cells[3, 4];
                    cellInDK.IsRichText = true;
                    var dk = cellInDK.RichText.Add("Study in DK?");

                    var headingCells = ws.Cells[3, 1, 3, 4];
                    headingCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headingCells.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                    int curRow = 4;

                    foreach (DictionaryEntry student in eDic)
                    {
                        Student s = (Student)student.Value;

                        ws.Cells[curRow, 1].Value = s.Name;
                        ws.Cells[curRow, 2].Value = s.Email;
                        ws.Cells[curRow, 3].Value = s.JobTitle;
                        ws.Cells[curRow, 4].Value = s.isStudyingiInDK;
                        //TODO Make thing red or green depending on studying in DK or not.
                        curRow++;
                    }

                }

                ex.SaveAs(curResultFileInfo);
                appendTextBox("Excel file saved to: " + ResultsFolder);
            }
        }
    }
}
