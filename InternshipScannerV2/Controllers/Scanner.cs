﻿using System;
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
        private ArrayList studentsInCurEducation;
        private int curStudentIndex;
        private int curEducationIndex;
        public TextBox tb;

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
            string currentDirectory = Directory.GetCurrentDirectory();
            appendTextBox(currentDirectory);
            currentDirectory = currentDirectory.Replace(@"InternshipScannerV2\bin\Debug\netcoreapp3.0", "");
            appendTextBox(currentDirectory);
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

        /// <summary>
        /// This function will take all of the data gained from running the program and the users inputs,
        /// and will then build a nice excel sheet out of it. 
        /// </summary>
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
                ExcelWorksheet wsg = ex.Workbook.Worksheets.Add("Overview");

                //Make the Title
                var cellOve = wsg.Cells[1, 1];
                cellOve.IsRichText = true;
                var Ove = cellOve.RichText.Add("General Overview");
                Ove.Bold = true;
                Ove.Size = 24;
                cellOve.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cellOve.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);
                wsg.Cells[1, 1, 1, 10].Merge = true;

                //Make the columns
                wsg.Cells[3, 1].Value = "Education";
                wsg.Cells[3, 2].Value = "Total Students";
                wsg.Cells[3, 3].Value = "Students in DK";
                wsg.Cells[3, 4].Value = "% of Students in DK";
                var headingCells2 = wsg.Cells[3, 1, 3, 4];
                headingCells2.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headingCells2.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                headingCells2.Style.Border.BorderAround(ExcelBorderStyle.Medium);

                int curOverviewCell = 4;
                int finalEducationRow = 0;

                foreach (string education in educations)
                {
                    IDictionary eDic = (IDictionary)IntStudentsDictionary[education];
                    ExcelWorksheet ws = ex.Workbook.Worksheets.Add(education);

                    var cellEdu = ws.Cells[1, 1];
                    cellEdu.IsRichText = true;
                    var Edu = cellEdu.RichText.Add(education);
                    Edu.Bold = true;
                    Edu.Size = 24;
                    cellEdu.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cellEdu.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);
                    ws.Cells[1, 1, 1, 10].Merge = true;

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
                    headingCells.Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    int curRow = 4;
                    int totalStudents = 0;
                    int dkStudents = 0;

                    foreach (DictionaryEntry student in eDic)
                    {
                        Student s = (Student)student.Value;

                        ws.Cells[curRow, 1].Value = s.Name;
                        ws.Cells[curRow, 2].Value = s.Email;
                        ws.Cells[curRow, 3].Value = s.JobTitle;

                        var dkCell = ws.Cells[curRow, 4];
                        dkCell.IsRichText = true;
                        dkCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        dkCell.Value = s.isStudyingiInDK;
                        if (s.isStudyingiInDK)
                        {
                            dkCell.Style.Fill.BackgroundColor.SetColor(Color.PaleGreen);
                            dkStudents++;
                        }
                        else dkCell.Style.Fill.BackgroundColor.SetColor(Color.PaleVioletRed);

                        totalStudents++;
                        curRow++;
                    }

                    //Total up all of the students and display it
                    curRow++;
                    ws.Cells[curRow, 1].Value = "Total Students";
                    ws.Cells[curRow, 2].Value = "Students in DK";
                    ws.Cells[curRow, 3].Value = "% of Students in DK";
                    var headingCells3 = ws.Cells[curRow, 1, curRow, 3];
                    headingCells3.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headingCells3.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    headingCells3.Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    curRow++;

                    var totalCell = ws.Cells[curRow, 1];
                    totalCell.Value = totalStudents;

                    var dkStudentCell = ws.Cells[curRow, 2];
                    dkStudentCell.Value = dkStudents;

                    var percentageCell = ws.Cells[curRow, 3];
                    percentageCell.Formula = "=(" + dkStudentCell.Address + "/" + totalCell.Address + ")";
                    percentageCell.Style.Numberformat.Format = "0%";
                    percentageCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    percentageCell.Style.Fill.BackgroundColor.SetColor(Color.PaleTurquoise);

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();

                    //Add these statistics to the 'Overview' tab
                    wsg.Cells[curOverviewCell, 1].Value = education;
                    
                    var genTotalStudents = wsg.Cells[curOverviewCell, 2];
                    genTotalStudents.Value = totalStudents;

                    var genDKStudents = wsg.Cells[curOverviewCell, 3];
                    genDKStudents.Value = dkStudents;

                    var genPercentageCell = wsg.Cells[curOverviewCell, 4];
                    genPercentageCell.Formula = "=(" + genDKStudents.Address + "/" + genTotalStudents.Address + ")";
                    genPercentageCell.Style.Numberformat.Format = "0%";
                    genPercentageCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    genPercentageCell.Style.Fill.BackgroundColor.SetColor(Color.PaleTurquoise);

                    //Log the last row where an education was imputed.
                    finalEducationRow = curOverviewCell;

                    curOverviewCell++;
                }
                //Colour all of the Education titles
                var educationTitles = wsg.Cells[4, 1, (curOverviewCell-1), 1];
                educationTitles.Style.Fill.PatternType = ExcelFillStyle.Solid;
                educationTitles.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                curOverviewCell++;

                wsg.Cells[curOverviewCell, 2].Value = "All Students";
                wsg.Cells[curOverviewCell, 3].Value = "All Students in DK";
                wsg.Cells[curOverviewCell, 4].Value = "% of All Students in DK";
                var headingCells4 = wsg.Cells[curOverviewCell, 1, curOverviewCell, 3];
                headingCells4.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headingCells4.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                headingCells4.Style.Border.BorderAround(ExcelBorderStyle.Medium);
                curOverviewCell++;

                var allTotalStudents = wsg.Cells[curOverviewCell, 2];
                allTotalStudents.Formula = "=SUM(B4:B"+ finalEducationRow + ")";

                var allDKStudents = wsg.Cells[curOverviewCell, 3];
                allDKStudents.Formula = "=SUM(C4:C" + finalEducationRow + ")";

                var allDKStudentsPer = wsg.Cells[curOverviewCell, 4];
                allDKStudentsPer.Formula = "=(" + allDKStudents.Address + "/" + allTotalStudents.Address + ")";

                wsg.Cells[wsg.Dimension.Address].AutoFitColumns();

                ex.SaveAs(curResultFileInfo);
                appendTextBox("Excel file saved to: " + ResultsFolder);
            }
        }
    }
}
