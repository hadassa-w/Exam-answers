using System;
using System.Collections.Generic;
using System.IO;
using Project.Models;
using Project.Services;

try
{
    string inputCsv = "data.csv";
    string projectRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
    string templatePath = Path.Combine(projectRoot, "Templates", "LetterTemplate.docx");
    string outputFolder = Path.Combine(projectRoot, "Output", "PDFs");
    string processedCsv = Path.Combine(projectRoot, "Output", "ProcessedData.xlsx");

    Directory.CreateDirectory(Path.GetDirectoryName(outputFolder)!);

    // יצירת תיקיית פלט אם אינה קיימת
    if (!Directory.Exists(outputFolder))
        Directory.CreateDirectory(outputFolder);

    // טעינת נתונים
    List<Person> people = PersonService.LoadData(inputCsv);

    // ניקוי כפילויות ושמות
    people = PersonService.CleanData(people);

    // חישוב ציון סופי והכנת הודעה
    PersonService.CalculateFinalScore(people);

    // שמירת הנתונים המעובדים ל-CSV חדש
    PersonService.SaveProcessedDataExcel(people, processedCsv);
    Console.WriteLine($"Processed CSV saved to: {processedCsv}");

    // הפקת מכתבים לכל מי שעבר את ההכשרה
    PersonService.GenerateAllLetters(people, templatePath, outputFolder);
    Console.WriteLine($"PDF letters generated in folder: {outputFolder}");

    Console.WriteLine("All operations completed successfully!");
}
catch (Exception ex)
{
    Console.WriteLine("Error: " + ex.Message);
}
