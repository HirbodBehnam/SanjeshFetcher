using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Color = System.Drawing.Color;

namespace SanjeshFetcher
{
    class Program
    {
        /// <summary>
        /// The student struct
        /// </summary>
        struct Student
        {
            /// <summary>
            /// True if student is a math
            /// </summary>
            public bool IsMathStudent;
            /// <summary>
            /// Ture if student is male
            /// </summary>
            public bool IsMale;
            /// <summary>
            /// شماره پرونده
            /// </summary>
            public string Parvande;
            /// <summary>
            /// The name of the student
            /// </summary>
            public string Name;
            /// <summary>
            /// Local rank of student in separate groups (rotbe mantaghe grooh 1,2,3)
            /// </summary>
            public int[] Rank;
            /// <summary>
            /// Overall rank of student
            /// </summary>
            public int RankOverall;
            /// <summary>
            /// Global rank of student in country in separate groups
            /// </summary>
            public int[] RankCountry;
            /// <summary>
            /// Global rank of student in country
            /// </summary>
            public int RankCountryOverall;
            /// <summary>
            /// رتبه سهمیه
            /// </summary>
            public int[] RankSpecial;
            /// <summary>
            /// رتبه سهمیه کلی
            /// </summary>
            public int RankSpecialOverall;
            /// <summary>
            /// Normalized score of student in different groups (taraz)
            /// </summary>
            public int[] NormalizedScore;
            /// <summary>
            /// Overall normalized score of student
            /// </summary>
            public int NormalizedScoreOverall;
            /// <summary>
            /// Grades of the student.
            /// The first four are literature, arabic, religion and english
            /// The next ones for math students are math, physics and chemistry
            /// For the other ones are TODO
            /// </summary>
            public float[] Grades;
            /// <summary>
            /// Answers that the student has given.
            /// 0 Means no answer
            /// </summary>
            public byte[] Answers;
            /// <summary>
            /// The key to the answers
            /// </summary>
            public byte[] AnswersKey;
            /// <summary>
            /// The code of omoomi answer sheet
            /// </summary>
            public string OmoomiCode;
            /// <summary>
            /// The code of ekhtesasi answer sheet
            /// </summary>
            public string EkhtesasiCode;
        }

        private const string MainResultUrl =
            "http://srv2.sanjesh.org/p_krn/index.php/krn_ntj_sar990629/sar_ntj1_99arm/krn/";

        private const string AnswerSheetUrl =
            "http://srv2.sanjesh.org/p_krn/index.php/krn_ntj_sar990629/sar_ntj1_99arm/krn/pasokh/";

        private const string DataCompiled =
            "number_p={0}&c_rah={1}&id_number={2}&hidecod=127711917198&form=3&captcha=12557&submit22=%CC%D3%CA%CC%E6";

        private const string UserAgent =
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36";

        private static readonly HttpClient Client = new HttpClient();

        private const string HeaderOfExcel = "تحلیل کنکور 99 ";
        public static void Main(string[] args)
        {
            // Setup the client and start get results
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Client.DefaultRequestHeaders.Add("User-Agent", UserAgent);
            Client.DefaultRequestHeaders.Add("Referer", MainResultUrl);
            // Get the username and passwords from config
            ConcurrentBag<Student> students;
            {
                var lines = File.ReadAllLines("students.txt");
                students = new ConcurrentBag<Student>();
                int i = 0;
                Parallel.ForEach(lines, (line) =>
                {
                    Console.Write("\rFetching {0}%", (float)i / lines.Length * 100f);
                    var splitData = line.Split(' ');
                    if (splitData.Length < 3) // ignore invalid lines
                    {
                        i++;
                        return;
                    } 
                    var student = GetStudentData(splitData[0], splitData[1], splitData[2]);
                    if (student.Answers != null) // check incorrect username
                        students.Add(student);
                    i++;
                });
                Console.WriteLine("\rFetching Done");
            }
            // Now do some statics stuff for math students
            Console.WriteLine("Doing calculations for math students...");
            using (var excel = new ExcelPackage(new FileInfo("ریاضی.xlsx")))
            {
                // Create each subject
                foreach (var (title, ranges) in StudentHelpers.MathSubjects)
                {
                    var worksheet = excel.Workbook.Worksheets.Add(title + " موضوعی");
                    worksheet.View.RightToLeft = true;
                    worksheet.Column(1).Width = 27.5d;
                    // Setup the header
                    {
                        worksheet.Cells["A1"].Value = HeaderOfExcel + title;
                        worksheet.Cells["A1:D1"].Merge = true;
                        worksheet.Cells["A2"].Value = "نام و نام خانوادگی";
                        worksheet.Cells["B2"].Value = "رشته";
                        worksheet.Cells["C2"].Value = "جنس";
                        worksheet.Cells["D2"].Value = "کد";
                    }
                    // Add header of sub-subjects
                    for (int i = 0; i < ranges.Length; i++)
                    {
                        (string subStat, _) = ranges[i];
                        // Add header
                        worksheet.Cells[1, 4 * i + 5, 1, 4 * i + 8].Merge = true;
                        worksheet.Cells[1, 4 * i + 5].Value = subStat;
                        worksheet.Cells[1, 4 * i + 5, 1, 4 * i + 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 4 * i + 5, 1, 4 * i + 8].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[1, 4 * i + 5, 1, 4 * i + 8].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        // Add const of "سفید" & ...
                        worksheet.Cells[2, 4 * i + 5].Value = "صحیح";
                        worksheet.Cells[2, 4 * i + 5].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 5].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 6].Value = "غلط";
                        worksheet.Cells[2, 4 * i + 6].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 6].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 7].Value = "سفید";
                        worksheet.Cells[2, 4 * i + 7].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 7].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 7].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 8].Value = "درصد";
                        worksheet.Cells[2, 4 * i + 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 8].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    }
                    // Add students
                    int rowCounter = 3; // start from row 3
                    foreach (var student in students)
                    {
                        if (!student.IsMathStudent)
                            continue;
                        // Add normal info
                        worksheet.Cells[rowCounter, 1].Value = student.Name;
                        worksheet.Cells[rowCounter, 2].Value = "ریاضی";
                        worksheet.Cells[rowCounter, 3].Value = student.IsMale ? "مرد" : "زن";
                        worksheet.Cells[rowCounter, 4].Value = student.Parvande;
                        // Add subjects info
                        for (int i = 0; i < ranges.Length; i++)
                        {
                            (_, int[] questions) = ranges[i];
                            int correct = 0, white = 0, mistake = 0;
                            foreach (var question in questions) // TODO: maybe marge statics here?
                            {
                                // also note question - 1
                                if (student.Answers[question - 1] == 0) // check empty answer
                                    white++;
                                else if (student.Answers[question - 1] == student.AnswersKey[question - 1])
                                    correct++;
                                else
                                    mistake++;
                            }
                            worksheet.Cells[rowCounter, 4 * i + 5].Value = correct;
                            worksheet.Cells[rowCounter, 4 * i + 6].Value = mistake;
                            worksheet.Cells[rowCounter, 4 * i + 7].Value = white;
                            worksheet.Cells[rowCounter, 4 * i + 8].Formula = ExcelFormula.SubjectFormula;
                            // styling
                            for (int j = 0; j < 4; j++)
                                worksheet.Cells[rowCounter, 4 * i + 5 + j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells[rowCounter, 4 * i + 8].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            worksheet.Cells[rowCounter, 4 * i + 5].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        }
                        // Add a row to row counter
                        rowCounter++;
                    }
                    // Add average
                    worksheet.Cells[rowCounter, 1].Value = "میانگین";
                    worksheet.Cells[rowCounter, 2].Value = "ریاضی";
                    worksheet.Cells[rowCounter, 3].Value = "-";
                    worksheet.Cells[rowCounter, 4].Value = "-";
                    for (int i = 5; i <= worksheet.Dimension.Columns; i++) // note that it must be <=
                        worksheet.Cells[rowCounter, i].Formula = ExcelFormula.AverageBottomFormula;
                    // Do styling
                    // Add borders to some headers; Note that these must be added at last to override other borders of the other cells
                    worksheet.Cells[1, 1, 1, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    for (int i = 1; i <= 4; i++) // add for second row
                        worksheet.Cells[2, i].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    for (int i = 3; i <= worksheet.Dimension.Rows; i++) // add for names and average
                    {
                        for (int j = 1; j <= 4; j++)
                            worksheet.Cells[i, j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells[i, 4].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    }
                    for (int i = 1; i <= worksheet.Dimension.Columns; i++) // Footer (average)
                    {
                        worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        if (i % 4 == 0) // Note that there is no need for checking the (i % 4 == 1) and editing the left side
                            worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Right.Style =
                                ExcelBorderStyle.Medium;
                    }
                    //worksheet.Cells.Style.Border.BorderAround(ExcelBorderStyle.Thin); 
                    worksheet.Cells[1, 1, 2, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, 2, worksheet.Dimension.Columns].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    // Footer style
                    worksheet.Cells[rowCounter, 1, rowCounter, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[rowCounter, 1, rowCounter, worksheet.Dimension.Columns].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    // Center everything
                    worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                // Create question specific sheets
                for (int qIndex = 0; qIndex < StudentHelpers.MathQuestions; qIndex++) // Note 
                {
                    // At first check where does this question belong
                    string subjectName = "";
                    foreach (var (title, ranges) in StudentHelpers.MathSubjects)
                        if (ranges.Last().Item2.Contains(qIndex + 1)) // note qIndex + 1
                        {
                            subjectName = title;
                            break;
                        }

                    if (subjectName == "")
                    {
                        Console.WriteLine("Error on getting subject name for question {0}", qIndex);
                        continue;
                    }

                    ExcelWorksheet worksheet = excel.Workbook.Worksheets[subjectName + " سوال به سوال"];
                    if (worksheet == null)
                    {
                        worksheet = excel.Workbook.Worksheets.Add(subjectName + " سوال به سوال");
                        // do some styling
                        worksheet.Cells[1, 1].Value = "عنوان";
                        worksheet.Cells[2, 1].Value = "شماره";
                        worksheet.Cells[3, 1].Value = "تعداد گزینه صحیح";
                        worksheet.Cells[4, 1].Value = "درصد گزینه صحیح";
                        worksheet.Cells[5, 1].Value = "تعداد گزینه اشتباه";
                        worksheet.Cells[6, 1].Value = "درصد گزینه اشتباه";
                        worksheet.Cells[7, 1].Value = "تعداد گزینه سفید";
                        worksheet.Cells[8, 1].Value = "درصد گزینه سفید";
                        worksheet.Cells[9, 1].Value = "درصد سوال";

                        worksheet.Column(1).Width = 15d;
                        worksheet.Row(1).Height = 150d;

                        worksheet.View.RightToLeft = true;
                    }
                    // Get the statics
                    int correct = 0, white = 0, wrong = 0;
                    foreach (var student in students)
                    {
                        if (!student.IsMathStudent) // ignore non math students
                            continue;
                        if (student.Answers[qIndex] == 0) // question is white
                            white++;
                        else if (student.Answers[qIndex] == student.AnswersKey[qIndex]) // correct answer
                            correct++;
                        else // wrong answer
                            wrong++;
                    }
                    // Fill the style sheet
                    // Get the title
                    string questionTitle = subjectName;
                    foreach (var (name, ranges) in StudentHelpers.MathSubjects)
                    {
                        // ReSharper disable once InvertIf Do not check other subjects
                        if (name == subjectName)
                        {
                            foreach (var (candidateTitle, range) in ranges)
                            {
                                // ReSharper disable once InvertIf We must not check other titles
                                if (range.Contains(qIndex + 1))
                                {
                                    questionTitle = candidateTitle;
                                    break;
                                }
                            }
                            break;
                        }
                    }
                    // Now fill the sheet
                    int column = worksheet.Dimension.Columns + 1;
                    worksheet.Cells[1, column].Value = questionTitle;
                    worksheet.Cells[1, column].Style.TextRotation = 180;
                    worksheet.Cells[2, column].Value = (qIndex + 1).ToString();
                    worksheet.Cells[3, column].Value = correct;
                    worksheet.Cells[4, column].Formula = ExcelFormula.QuestionsCorrectPercentage;
                    worksheet.Cells[5, column].Value = wrong;
                    worksheet.Cells[6, column].Formula = ExcelFormula.QuestionsWrongPercentage;
                    worksheet.Cells[7, column].Value = white;
                    worksheet.Cells[8, column].Formula = ExcelFormula.QuestionsWhitePercentage;
                    worksheet.Cells[9, column].Formula = ExcelFormula.QuestionsPercentage;
                }
                // Final styling for question specific sheets
                foreach (var (title, _) in StudentHelpers.MathSubjects)
                {
                    var worksheet = excel.Workbook.Worksheets[title + " سوال به سوال"];
                    // Do coloring
                    worksheet.Cells[1, 1, 2, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, 2, worksheet.Dimension.Columns].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    // Do borders
                    for (int i = 1; i <= worksheet.Dimension.Rows; i++)
                        for (int j = 1; j <= worksheet.Dimension.Columns; j++)
                        {
                            worksheet.Cells[i, j].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            worksheet.Cells[i, j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        }

                    for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                    {
                        worksheet.Cells[2, i].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; // title rows
                        worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; // last row
                    }
                    // Do text alignment
                    worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
                // Create overall sheet
                {
                    var worksheet = excel.Workbook.Worksheets.Add("نتایج کلی");
                    worksheet.View.RightToLeft = true;
                    worksheet.Column(1).Width = 27.5d; // names are big
                    // Add headers
                    worksheet.Cells[1, 1].Value = "نام و نام خانوادگی";
                    worksheet.Cells[1, 2].Value = "ادبیات";
                    worksheet.Cells[1, 3].Value = "عربی";
                    worksheet.Cells[1, 4].Value = "دینی";
                    worksheet.Cells[1, 5].Value = "زبان";
                    worksheet.Cells[1, 6].Value = "ریاضی";
                    worksheet.Cells[1, 7].Value = "فیزیک";
                    worksheet.Cells[1, 8].Value = "شیمی";
                    worksheet.Cells[1, 9].Value = "رتبه منطقه";
                    worksheet.Cells[1, 10].Value = "رتبه کشور";
                    worksheet.Cells[1, 11].Value = "تراز کل";
                    worksheet.Cells[1, 12].Value = "رتبه سهمیه";
                    // Add students
                    int rowCounter = 2;
                    foreach (var student in students)
                    {
                        if (!student.IsMathStudent)
                            continue;
                        // Add data
                        worksheet.Cells[rowCounter, 1].Value = student.Name;
                        for (int i = 0; i < student.Grades.Length; i++)
                            worksheet.Cells[rowCounter, i + 2].Value = student.Grades[i];
                        worksheet.Cells[rowCounter, 9].Value = student.RankOverall;
                        worksheet.Cells[rowCounter, 10].Value = student.RankCountryOverall;
                        worksheet.Cells[rowCounter, 11].Value = student.NormalizedScoreOverall;
                        if (student.RankSpecialOverall == 0)
                            worksheet.Cells[rowCounter, 12].Value = "-";
                        else
                            worksheet.Cells[rowCounter, 12].Value = student.RankSpecialOverall;
                        // add to row counter
                        rowCounter++;
                    }
                    // Do styling
                    worksheet.Cells[1, 1, 1, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, 1, 12].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    // Draw some borders
                    for (int i = 1; i <= worksheet.Dimension.Rows; i++)
                    {
                        worksheet.Cells[i, 1].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[i, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    }
                    for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                    {
                        worksheet.Cells[1, i].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[1, i].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    }
                    for (int i = 2; i <= worksheet.Dimension.Rows; i++)
                        for (int j = 2; j <= worksheet.Dimension.Columns; j++)
                        {
                            worksheet.Cells[i, j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[i, j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        }
                }
                excel.Save();
            }
            // Do it again for tajrobi students
            Console.WriteLine("Doing calculations for tajrobi students...");
            using (var excel = new ExcelPackage(new FileInfo("تجربی.xlsx")))
            {
                // Create each subject
                foreach (var (title, ranges) in StudentHelpers.TajrobiSubjects)
                {
                    var worksheet = excel.Workbook.Worksheets.Add(title + " موضوعی");
                    worksheet.View.RightToLeft = true;
                    worksheet.Column(1).Width = 27.5d;
                    // Setup the header
                    {
                        worksheet.Cells["A1"].Value = HeaderOfExcel + title;
                        worksheet.Cells["A1:D1"].Merge = true;
                        worksheet.Cells["A2"].Value = "نام و نام خانوادگی";
                        worksheet.Cells["B2"].Value = "رشته";
                        worksheet.Cells["C2"].Value = "جنس";
                        worksheet.Cells["D2"].Value = "کد";
                    }
                    // Add header of sub-subjects
                    for (int i = 0; i < ranges.Length; i++)
                    {
                        (string subStat, _) = ranges[i];
                        // Add header
                        worksheet.Cells[1, 4 * i + 5, 1, 4 * i + 8].Merge = true;
                        worksheet.Cells[1, 4 * i + 5].Value = subStat;
                        worksheet.Cells[1, 4 * i + 5, 1, 4 * i + 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 4 * i + 5, 1, 4 * i + 8].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[1, 4 * i + 5, 1, 4 * i + 8].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        // Add const of "سفید" & ...
                        worksheet.Cells[2, 4 * i + 5].Value = "صحیح";
                        worksheet.Cells[2, 4 * i + 5].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 5].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 6].Value = "غلط";
                        worksheet.Cells[2, 4 * i + 6].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 6].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 7].Value = "سفید";
                        worksheet.Cells[2, 4 * i + 7].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 7].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 7].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 4 * i + 8].Value = "درصد";
                        worksheet.Cells[2, 4 * i + 8].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 4 * i + 8].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    }
                    // Add students
                    int rowCounter = 3; // start from row 3
                    foreach (var student in students)
                    {
                        if (student.IsMathStudent)
                            continue;
                        // Add normal info
                        worksheet.Cells[rowCounter, 1].Value = student.Name;
                        worksheet.Cells[rowCounter, 2].Value = "تجربی";
                        worksheet.Cells[rowCounter, 3].Value = student.IsMale ? "مرد" : "زن";
                        worksheet.Cells[rowCounter, 4].Value = student.Parvande;
                        // Add subjects info
                        for (int i = 0; i < ranges.Length; i++)
                        {
                            (_, int[] questions) = ranges[i];
                            int correct = 0, white = 0, mistake = 0;
                            foreach (var question in questions) // TODO: maybe marge statics here?
                            {
                                // also note question - 1
                                if (student.Answers[question - 1] == 0) // check empty answer
                                    white++;
                                else if (student.Answers[question - 1] == student.AnswersKey[question - 1])
                                    correct++;
                                else
                                    mistake++;
                            }
                            worksheet.Cells[rowCounter, 4 * i + 5].Value = correct;
                            worksheet.Cells[rowCounter, 4 * i + 6].Value = mistake;
                            worksheet.Cells[rowCounter, 4 * i + 7].Value = white;
                            worksheet.Cells[rowCounter, 4 * i + 8].Formula = ExcelFormula.SubjectFormula;
                            // styling
                            for (int j = 0; j < 4; j++)
                                worksheet.Cells[rowCounter, 4 * i + 5 + j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells[rowCounter, 4 * i + 8].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            worksheet.Cells[rowCounter, 4 * i + 5].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        }
                        // Add a row to row counter
                        rowCounter++;
                    }
                    // Add average
                    worksheet.Cells[rowCounter, 1].Value = "میانگین";
                    worksheet.Cells[rowCounter, 2].Value = "تجربی";
                    worksheet.Cells[rowCounter, 3].Value = "-";
                    worksheet.Cells[rowCounter, 4].Value = "-";
                    for (int i = 5; i <= worksheet.Dimension.Columns; i++) // note that it must be <=
                        worksheet.Cells[rowCounter, i].Formula = ExcelFormula.AverageBottomFormula;
                    // Do styling
                    // Add borders to some headers; Note that these must be added at last to override other borders of the other cells
                    worksheet.Cells[1, 1, 1, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    for (int i = 1; i <= 4; i++) // add for second row
                        worksheet.Cells[2, i].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    for (int i = 3; i <= worksheet.Dimension.Rows; i++) // add for names and average
                    {
                        for (int j = 1; j <= 4; j++)
                            worksheet.Cells[i, j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells[i, 4].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    }
                    for (int i = 1; i <= worksheet.Dimension.Columns; i++) // Footer (average)
                    {
                        worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        if (i % 4 == 0) // Note that there is no need for checking the (i % 4 == 1) and editing the left side
                            worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Right.Style =
                                ExcelBorderStyle.Medium;
                    }
                    //worksheet.Cells.Style.Border.BorderAround(ExcelBorderStyle.Thin); 
                    worksheet.Cells[1, 1, 2, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, 2, worksheet.Dimension.Columns].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    // Footer style
                    worksheet.Cells[rowCounter, 1, rowCounter, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[rowCounter, 1, rowCounter, worksheet.Dimension.Columns].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    // Center everything
                    worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                // Create question specific sheets
                for (int qIndex = 0; qIndex < StudentHelpers.TajrobiQuestions; qIndex++) // Note 
                {
                    // At first check where does this question belong
                    string subjectName = "";
                    foreach (var (title, ranges) in StudentHelpers.TajrobiSubjects)
                        if (ranges.Last().Item2.Contains(qIndex + 1)) // note qIndex + 1
                        {
                            subjectName = title;
                            break;
                        }

                    if (subjectName == "")
                    {
                        Console.WriteLine("Error on getting subject name for question {0}", qIndex);
                        continue;
                    }

                    ExcelWorksheet worksheet = excel.Workbook.Worksheets[subjectName + " سوال به سوال"];
                    if (worksheet == null)
                    {
                        worksheet = excel.Workbook.Worksheets.Add(subjectName + " سوال به سوال");
                        // do some styling
                        worksheet.Cells[1, 1].Value = "عنوان";
                        worksheet.Cells[2, 1].Value = "شماره";
                        worksheet.Cells[3, 1].Value = "تعداد گزینه صحیح";
                        worksheet.Cells[4, 1].Value = "درصد گزینه صحیح";
                        worksheet.Cells[5, 1].Value = "تعداد گزینه اشتباه";
                        worksheet.Cells[6, 1].Value = "درصد گزینه اشتباه";
                        worksheet.Cells[7, 1].Value = "تعداد گزینه سفید";
                        worksheet.Cells[8, 1].Value = "درصد گزینه سفید";
                        worksheet.Cells[9, 1].Value = "درصد سوال";

                        worksheet.Column(1).Width = 15d;
                        worksheet.Row(1).Height = 150d;

                        worksheet.View.RightToLeft = true;
                    }
                    // Get the statics
                    int correct = 0, white = 0, wrong = 0;
                    foreach (var student in students)
                    {
                        if (student.IsMathStudent) // ignore non math students
                            continue;
                        if (student.Answers[qIndex] == 0) // question is white
                            white++;
                        else if (student.Answers[qIndex] == student.AnswersKey[qIndex]) // correct answer
                            correct++;
                        else // wrong answer
                            wrong++;
                    }
                    // Fill the style sheet
                    // Get the title
                    string questionTitle = subjectName;
                    foreach (var (name, ranges) in StudentHelpers.TajrobiSubjects)
                    {
                        // ReSharper disable once InvertIf Do not check other subjects
                        if (name == subjectName)
                        {
                            foreach (var (candidateTitle, range) in ranges)
                            {
                                // ReSharper disable once InvertIf We must not check other titles
                                if (range.Contains(qIndex + 1))
                                {
                                    questionTitle = candidateTitle;
                                    break;
                                }
                            }
                            break;
                        }
                    }
                    // Now fill the sheet
                    int column = worksheet.Dimension.Columns + 1;
                    worksheet.Cells[1, column].Value = questionTitle;
                    worksheet.Cells[1, column].Style.TextRotation = 180;
                    worksheet.Cells[2, column].Value = (qIndex + 1).ToString();
                    worksheet.Cells[3, column].Value = correct;
                    worksheet.Cells[4, column].Formula = ExcelFormula.QuestionsCorrectPercentage;
                    worksheet.Cells[5, column].Value = wrong;
                    worksheet.Cells[6, column].Formula = ExcelFormula.QuestionsWrongPercentage;
                    worksheet.Cells[7, column].Value = white;
                    worksheet.Cells[8, column].Formula = ExcelFormula.QuestionsWhitePercentage;
                    worksheet.Cells[9, column].Formula = ExcelFormula.QuestionsPercentage;
                }
                // Final styling for question specific sheets
                foreach (var (title, _) in StudentHelpers.TajrobiSubjects)
                {
                    var worksheet = excel.Workbook.Worksheets[title + " سوال به سوال"];
                    // Do coloring
                    worksheet.Cells[1, 1, 2, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, 2, worksheet.Dimension.Columns].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    // Do borders
                    for (int i = 1; i <= worksheet.Dimension.Rows; i++)
                        for (int j = 1; j <= worksheet.Dimension.Columns; j++)
                        {
                            worksheet.Cells[i, j].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            worksheet.Cells[i, j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        }

                    for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                    {
                        worksheet.Cells[2, i].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; // title rows
                        worksheet.Cells[worksheet.Dimension.Rows, i].Style.Border.Bottom.Style = ExcelBorderStyle.Medium; // last row
                    }
                    // Do text alignment
                    worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
                // Create overall sheet
                {
                    var worksheet = excel.Workbook.Worksheets.Add("نتایج کلی");
                    worksheet.View.RightToLeft = true;
                    worksheet.Column(1).Width = 27.5d; // names are big
                    // Add headers
                    worksheet.Cells[1, 1].Value = "نام و نام خانوادگی";
                    worksheet.Cells[1, 2].Value = "ادبیات";
                    worksheet.Cells[1, 3].Value = "عربی";
                    worksheet.Cells[1, 4].Value = "دینی";
                    worksheet.Cells[1, 5].Value = "زبان";
                    worksheet.Cells[1, 6].Value = "زمین شناسی";
                    worksheet.Cells[1, 7].Value = "ریاضی";
                    worksheet.Cells[1, 8].Value = "زیست شناسی";
                    worksheet.Cells[1, 9].Value = "فیزیک";
                    worksheet.Cells[1, 10].Value = "شیمی";
                    worksheet.Cells[1, 11].Value = "رتبه منطقه";
                    worksheet.Cells[1, 12].Value = "رتبه کشور";
                    worksheet.Cells[1, 13].Value = "تراز کل";
                    worksheet.Cells[1, 14].Value = "رتبه سهمیه";
                    // Add students
                    int rowCounter = 2;
                    foreach (var student in students)
                    {
                        if (student.IsMathStudent)
                            continue;
                        // Add data
                        worksheet.Cells[rowCounter, 1].Value = student.Name;
                        for (int i = 0; i < student.Grades.Length; i++)
                            worksheet.Cells[rowCounter, i + 2].Value = student.Grades[i];
                        worksheet.Cells[rowCounter, 11].Value = student.RankOverall;
                        worksheet.Cells[rowCounter, 12].Value = student.RankCountryOverall;
                        worksheet.Cells[rowCounter, 13].Value = student.NormalizedScoreOverall;
                        if (student.RankSpecialOverall == 0)
                            worksheet.Cells[rowCounter, 14].Value = "-";
                        else
                            worksheet.Cells[rowCounter, 14].Value = student.RankSpecialOverall;
                        // add to row counter
                        rowCounter++;
                    }
                    // Do styling
                    worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, worksheet.Dimension.Rows, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    // Draw some borders
                    for (int i = 1; i <= worksheet.Dimension.Rows; i++)
                    {
                        worksheet.Cells[i, 1].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[i, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    }
                    for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                    {
                        worksheet.Cells[1, i].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheet.Cells[1, i].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                    }
                    for (int i = 2; i <= worksheet.Dimension.Rows; i++)
                        for (int j = 2; j <= worksheet.Dimension.Columns; j++)
                        {
                            worksheet.Cells[i, j].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[i, j].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        }
                }
                excel.Save();
            }
            // Also output a json file for more details
            File.WriteAllText("students.json",JsonConvert.SerializeObject(students.ToArray()));
            Console.WriteLine("By Hirbod Behnam; AE High school");
        }
        /// <summary>
        /// Get the student's data from Sanjesh site
        /// </summary>
        /// <param name="parvande">شماره پرونده</param>
        /// <param name="rahgiri">کد رهگیری</param>
        /// <param name="idCode">کد ملی</param>
        /// <returns></returns>
        private static Student GetStudentData(string parvande, string rahgiri, string idCode)
        {
            var response = Client.PostAsync(MainResultUrl, new StringContent(string.Format(DataCompiled, parvande, rahgiri, idCode), Encoding.UTF8, "application/x-www-form-urlencoded")).Result;
            string content = Encoding.Default.GetString(response.Content.ReadAsByteArrayAsync().Result);
            // Get basic info of the student
            Student student = new Student
            {
                Parvande = parvande
            };
            var doc = new HtmlDocument();
            doc.LoadHtml(content);
            // Haha! string contains goes brrrrrr
            student.IsMathStudent = content.Contains("رياضي فيزيك");
            student.IsMale = content.Contains("مرد");
            // Get name
            try
            {
                var row = doc.DocumentNode.SelectNodes("//table[contains(@class, 'tt')]")[0]
                    .Descendants("tr").ToList()[1].Descendants("td").ToList();
                student.Name = row[row.Count - 2].InnerText;
            }
            catch (Exception) // incorrect username (probably?)
            {
                Console.WriteLine("\rError on user {0} {1} {2}", parvande, rahgiri, idCode);
                return new Student();
            }
            // Fix name
            {
                var s = student.Name.Split('-');
                student.Name = s[1].Trim() + " " + s[0].Trim();
            }
            // Get the main tables; The first one is used for the group ranks. The second is for "omoomi" lessons. The third is for "ekhtesasi" lessons. The last one is for overall status
            var masterTables = doc.GetElementbyId("table1").Descendants("tr").ToList()[1].Descendants("td").First().Descendants("table").ToArray();
            // Get group ranks
            {
                var rows = masterTables[0].Descendants("tr").ToList();
                bool specialCapacity = rows.Count == 8; // true if the student has سهمیه
                var localRanks = rows[specialCapacity ? 4 : 3].Descendants("td").ToList();
                student.Rank = new[]
                {
                    Helpers.ParsePersianNumber(localRanks[4].InnerText),
                    Helpers.ParsePersianNumber(localRanks[3].InnerText),
                    Helpers.ParsePersianNumber(localRanks[2].InnerText),
                };
                var countryRanks = rows[specialCapacity ? 5 : 4].Descendants("td").ToList();
                student.RankCountry = new[]
                {
                    Helpers.ParsePersianNumber(countryRanks[4].InnerText),
                    Helpers.ParsePersianNumber(countryRanks[3].InnerText),
                    Helpers.ParsePersianNumber(countryRanks[2].InnerText),
                };
                var normalizedScores = rows[specialCapacity ? 6 : 5].Descendants("td").ToList();
                student.NormalizedScore = new[]
                {
                    Helpers.ParsePersianNumber(normalizedScores[4].InnerText),
                    Helpers.ParsePersianNumber(normalizedScores[3].InnerText),
                    Helpers.ParsePersianNumber(normalizedScores[2].InnerText),
                };
                if (specialCapacity)
                {
                    var specialRank = rows[3].Descendants("td").ToList();
                    student.RankSpecial = new[]
                    {
                        Helpers.ParsePersianNumber(normalizedScores[4].InnerText),
                        Helpers.ParsePersianNumber(normalizedScores[3].InnerText),
                        Helpers.ParsePersianNumber(normalizedScores[2].InnerText),
                    };
                }
            }
            // Get "omoomi" results
            {
                var rowData = masterTables[1].Descendants("tr").ToList()[1].Descendants("td").ToList();
                student.Grades = new float[4]; // Note that this is hardcoded
                for (int i = 0; i < 4; i++)
                    student.Grades[i] = Convert.ToSingle(rowData[i].InnerText.Replace("/", ".")); // note that / must be replaced with .
            }
            // Get "ekhtesasi" results
            {
                var rowData = masterTables[2].Descendants("tr").ToList()[1].Descendants("td").ToList();
                var data = new List<float>(student.Grades);
                foreach (var subject in rowData)
                    data.Add(Convert.ToSingle(subject.InnerText.Replace("/", ".")));
                student.Grades = data.ToArray();
            }
            // Get final results
            {
                var rowData = masterTables[3].Descendants("tr").ToList()[1].Descendants("td").ToList();
                bool specialCapacity = rowData.Count == 5; // true if the student has سهمیه
                student.NormalizedScoreOverall = Helpers.ParsePersianNumber(rowData[specialCapacity ? 4 : 3].InnerHtml);
                student.RankOverall = Helpers.ParsePersianNumber(rowData[specialCapacity ? 3 : 2].InnerHtml);
                student.RankCountryOverall = Helpers.ParsePersianNumber(rowData[0].InnerHtml);
                if (specialCapacity)
                    student.RankSpecialOverall = Helpers.ParsePersianNumber(rowData[2].InnerHtml);
            }
            // Get the answer sheet
            {
                var inputs = doc.DocumentNode.Descendants("form").ToList()[1] // the first one is for diplom grades
                    .Descendants("input").ToList();
                var formValues = new Dictionary<string, string>();
                foreach (var input in inputs)
                {
                    // we must skip the button
                    if (input.Attributes["type"] == null || input.Attributes["type"].Value == "button")
                        continue;
                    formValues.Add(input.Attributes["name"].Value, input.Attributes["value"].Value);
                }
                var formContent = new FormUrlEncodedContent(formValues);
                response = Client.PostAsync(AnswerSheetUrl, formContent).Result;
                content = Encoding.Default.GetString(response.Content.ReadAsByteArrayAsync().Result);
                // Parse the answer sheet
                doc = new HtmlDocument();
                doc.LoadHtml(content);
                var tables = doc.DocumentNode.SelectNodes("//table[contains(@class, 'contacts')]").ToList();
                // Get the codes
                {
                    var row = tables[0].Descendants("table").First().Descendants("tr").ToList()[1].Descendants("td").ToList();
                    student.EkhtesasiCode = row[1].InnerText;
                    student.OmoomiCode = row[2].InnerText;
                }
                List<byte> answers = new List<byte>(320);
                List<byte> answersKey = new List<byte>(320);
                // Get omoomi sheet
                {
                    for (int tIndex = 1; tIndex < 5; tIndex++)
                    {
                        var rows = tables[tIndex].Descendants("tr").ToList(); // first table of omoomi exams
                        for (int i = 1; i < rows.Count; i++)
                        {
                            var rowData = rows[i].Descendants("td").ToList();
                            try
                            {
                                answers.Add((byte)Helpers.ParsePersianNumber(rowData[1].InnerText));
                            }
                            catch (Exception) // happens when the data is سفید
                            {
                                answers.Add(0);
                            }
                            // Check tajrobi changed fucking questions
                            try
                            {
                                answersKey.Add(Convert.ToByte(rowData[2].InnerText));
                            }
                            catch (Exception) // One of them is changed
                            {
                                var data = new string(rowData[2].InnerText.Where(char.IsDigit).ToArray());
                                answersKey.Add(Convert.ToByte(data));
                            }
                        }
                    }
                }
                // Get ekhtesasi sheet
                {
                    for (int tIndex = 6; tIndex < 12; tIndex++)
                    {
                        var rows = tables[tIndex].Descendants("tr").ToList();
                        for (int i = 1; i < rows.Count; i++)
                        {
                            var rowData = rows[i].Descendants("td").ToList();
                            try
                            {
                                answers.Add((byte)Helpers.ParsePersianNumber(rowData[1].InnerText));
                            }
                            catch (Exception) // happens when the data is سفید
                            {
                                answers.Add(0);
                            }
                            try
                            {
                                answersKey.Add(Convert.ToByte(rowData[2].InnerText));
                            }
                            catch (Exception) // happens when the data is سفید; note that in these tables the answers can be empty too; Also happens in tajrobi exam that a question is changed
                            {
                                var data = new string(rowData[2].InnerText.Where(char.IsDigit).ToArray());
                                answersKey.Add(data == "" ? (byte)0 : Convert.ToByte(data));
                            }
                        }
                    }
                }
                // Set the data
                student.Answers = answers.ToArray();
                student.AnswersKey = answersKey.ToArray();
            }
            return student;
        }
    }
}
