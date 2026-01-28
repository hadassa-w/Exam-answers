using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using Spire.Doc;
using ClosedXML.Excel;
using MyPerson = Project.Models.Person;

namespace Project.Services
{
    public class PersonService
    {
        // טעינת נתונים מקובץ CSV
        public static List<MyPerson> LoadData(string path)
        {
            using var reader = new StreamReader(path);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            return csv.GetRecords<MyPerson>().ToList();
        }

        // ניקוי כפילויות ושמות
        public static List<MyPerson> CleanData(List<MyPerson> people)
        {
            var cleaned = people
                .GroupBy(p => new { p.FirstName, p.LastName, p.Department })
                .Select(g => g.First())
                .ToList();

            foreach (var p in cleaned)
            {
                p.FirstName = (p.FirstName ?? "").Trim();
                p.LastName = (p.LastName ?? "").Trim();
                p.Department = (p.Department ?? "").Trim();
                p.Phone = (p.Phone ?? "").Trim();
                p.Email = (p.Email ?? "").Trim();

                p.FirstName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(p.FirstName.ToLower());
                p.LastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(p.LastName.ToLower());
                p.Department = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(p.Department.ToLower());
            }

            cleaned = cleaned.Where(p => !string.IsNullOrEmpty(p.FirstName) && !string.IsNullOrEmpty(p.LastName)).ToList();
            return cleaned;
        }

        // חישוב ציון סופי והכנת הודעה לכל משתמש
        public static void CalculateFinalScore(List<MyPerson> people)
        {
            foreach (var p in people)
            {
                p.FinalScore = 0.6 * p.PracticalScore + 0.4 * p.TheoryScore;

                if (p.FinalScore >= 90)
                {
                    p.Message = $"הרינו להודיעך כי עברת בהצלחה את ההכשרה. הציון הסופי שלך הינו {p.FinalScore:F2}{Environment.NewLine}" +
                                $"נמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית.";
                }
                else if (p.FinalScore >= 70)
                {
                    p.Message = $"הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.\n";
                }
            }
        }

        // שמירת הנתונים המעובדים ל-CSV חדש
        public static void SaveProcessedDataCsv(List<MyPerson> people, string outputPath)
        {
            using var writer = new StreamWriter(outputPath);
            using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
            csv.WriteHeader<MyPerson>();
            csv.NextRecord();

            foreach (var p in people)
            {
                csv.WriteRecord(p);
                csv.NextRecord();
            }
        }

        // שמירת הנתונים המעובדים ל-Excel חדש
        public static void SaveProcessedDataExcel(List<MyPerson> people, string outputPath)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("ProcessedData");

            worksheet.Cell(1, 1).Value = "FirstName";
            worksheet.Cell(1, 2).Value = "LastName";
            worksheet.Cell(1, 3).Value = "Department";
            worksheet.Cell(1, 4).Value = "Phone";
            worksheet.Cell(1, 5).Value = "Email";
            worksheet.Cell(1, 6).Value = "PracticalScore";
            worksheet.Cell(1, 7).Value = "TheoryScore";
            worksheet.Cell(1, 8).Value = "FinalScore";
            worksheet.Cell(1, 9).Value = "Message";

            for (int i = 0; i < people.Count; i++)
            {
                var p = people[i];
                worksheet.Cell(i + 2, 1).Value = p.FirstName;
                worksheet.Cell(i + 2, 2).Value = p.LastName;
                worksheet.Cell(i + 2, 3).Value = p.Department;
                worksheet.Cell(i + 2, 4).Value = p.Phone;
                worksheet.Cell(i + 2, 5).Value = p.Email;
                worksheet.Cell(i + 2, 6).Value = p.PracticalScore;
                worksheet.Cell(i + 2, 7).Value = p.TheoryScore;
                worksheet.Cell(i + 2, 8).Value = p.FinalScore;
                worksheet.Cell(i + 2, 9).Value = p.Message;
            }

            workbook.SaveAs(outputPath);
        }

        // הפקת מכתב PDF לפי Mail Merge
        public static void GenerateLetter(MyPerson p, string templatePath, string outputFolder)
        {
            var doc = new Document();
            doc.LoadFromFile(templatePath);

            doc.MailMerge.Execute(
                new string[] { "FullName", "FirstName", "LastName", "Department", "Phone", "Email", "FinalScore", "Message" },
                new string[] { $"{p.FirstName} {p.LastName}", p.FirstName, p.LastName, p.Department, p.Phone, p.Email, p.FinalScore.ToString("F2"), p.Message }
            );

            string safeFileName = $"{p.FirstName}_{p.LastName}"
                                  .Replace(" ", "_")
                                  .Replace("\"", "")
                                  .Replace("'", "")
                                  .Replace("\\", "")
                                  .Replace("/", "");

            string pdfFile = Path.Combine(outputFolder, $"{safeFileName}_Certificate.pdf");
            doc.SaveToFile(pdfFile, FileFormat.PDF);
        }

        // הפקת מכתבים לכל הרשימה (רק למי שעבר את ההכשרה)
        public static void GenerateAllLetters(List<MyPerson> people, string templatePath, string outputFolder)
        {
            foreach (var p in people)
            {
                if (p.FinalScore >= 70)
                    GenerateLetter(p, templatePath, outputFolder);
            }
        }
    }
}
