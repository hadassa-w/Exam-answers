# Person Certification Generator

## תאור כללי

תוכנית זו נבנתה ב-C# ומטרתה לייצר מכתבי הסמכה (PDF) לנבחנים על בסיס נתונים מקובץ CSV. התוכנית מבצעת את כל השלבים מהטעינה של הנתונים, ניקוי כפילויות, חישוב ציון סופי ועד הפקת מכתבים מותאמים אישית באמצעות **Mail Merge** בתבנית Word.

---

## פעולות התוכנית

### 1. טעינת נתונים

* קובץ מקור: CSV עם מידע על הנבחנים.
* כל הנתונים נטענים למבנה `Person` המכיל: `FirstName`, `LastName`, `Department`, `Phone`, `Email`, `PracticalScore`, `TheoryScore` (בהמשך נוסיף גם את שדות `FinalScore`, `Message`)

### 2. ניקוי נתונים

* הסרת רשומות כפולות על בסיס `FirstName + LastName + Department`.
* תיקון פורמט שמות (אות ראשונה גדולה, שאר האותיות קטנות).
* חיתוך רווחים מיותרים מכל השדות.

### 3. חישוב ציון סופי וסינון

* ציון סופי מחושב כך: `FinalScore = 0.6 * PracticalScore + 0.4 * TheoryScore`
* יצירת הודעה מותאמת לכל משתמש (`Message`) בהתאם לציון:

  * מעל 90: הצלחה עם המלצה לתפקיד מוביל טכנולוגי מחלקתי.
  * בין 70 ל-90: הצלחה אך לא עבר את ההכשרה (תפקיד לא נמצא).
* הודעות כוללות **מעבר שורה** (`Environment.NewLine`) שמאפשר הצגה נכונה בקובץ Excel וב-PDF.

### 4. שמירת הנתונים

* ניתן לשמור את הנתונים המעובדים:

  * כ-**CSV** (לצורך בדיקות).
  * כ-**Excel** (`.xlsx`) עם תא Message שמציג את ההודעה במעבר שורה תקין.

### 5. הפקת מכתבים (PDF)

* מבוצע באמצעות **Mail Merge** לתבנית Word אחת.
* השדות בתבנית: `FirstName, LastName, Department, Phone, Email, FinalScore, Message`
* כל מכתב נשמר כ-PDF עם שם קובץ תקני: `FirstName_LastName_Certificate.pdf`
* מכתבים נוצרים רק למי שעבר את ההכשרה (ציון סופי ≥ 70).

---

## שימוש בתוכנית

### דוגמת Main

```csharp
    string inputCsv = "data.csv";
    string projectRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
    string templatePath = Path.Combine(projectRoot, "Templates", "LetterTemplate.docx");
    string outputFolder = Path.Combine(projectRoot, "Output", "PDFs");
    string processedCsv = Path.Combine(projectRoot, "Output", "ProcessedData.xlsx");

    Directory.CreateDirectory(Path.GetDirectoryName(outputFolder)!);

    if (!Directory.Exists(outputFolder))
        Directory.CreateDirectory(outputFolder);

    List<Person> people = PersonService.LoadData(inputCsv);

    people = PersonService.CleanData(people);

    PersonService.CalculateFinalScore(people);

    PersonService.SaveProcessedDataExcel(people, processedCsv);
    Console.WriteLine($"Processed CSV saved to: {processedCsv}");

    PersonService.GenerateAllLetters(people, templatePath, outputFolder);
    Console.WriteLine($"PDF letters generated in folder: {outputFolder}");

    Console.WriteLine("All operations completed successfully!");
```

### דרישות

* תבנית Word אחת (**Template/LetterTemplate.docx**) עם **Mail Merge Fields**.
* ספריות:

  * `CsvHelper` – לקריאת CSV.
  * `ClosedXML` – ליצירת Excel.
  * `Spire.Doc` – ליצירת PDF מה-Mail Merge.

---

## שיפורים מומלצים

1. בדיקות שמירה על תקינות קבצים: אם קובץ PDF קיים – ליצור שם ייחודי או לשאול את המשתמש.
2. בדיקה שהשדות בתבנית Mail Merge קיימים לפני הפקת המכתב.
3. בדיקות על תוכן Excel – לוודא שההודעות ארוכות מספיק ומעברי השורה נשמרים.
4. לוג של ההודעות שנוצרו כדי למנוע הפקה של מסמך שגוי.
5. תמיכה בשפות שונות (UTF-8) בקבצי CSV ו-Excel.

---

## קישור ל-GitHub

[GitHub Repository - Person Certification Generator](https://github.com/hadassa-w/Exam-answers)

---

## סיכום

* קלט: CSV עם מידע נבחנים.
* פלט: Excel מעובד + PDF מכתבים מותאמים אישית.
* השימוש ב-Mail Merge מאפשר לנבחנים לא טכניים לערוך את התבנית בקלות.
* המערכת שומרת על פורמט אחיד ומאפשרת הפקה מהירה ומדויקת של מסמכים.
