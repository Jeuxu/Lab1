using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Domain;

namespace Data.Reports
{
    public class DocxReportService
    {
        public void GenerateReport(List<CityAirQuality> data, string filePath)
        {
            if (data == null || data.Count == 0)
                throw new ArgumentException("Дані відсутні для генерації звіту");

            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                Body body = mainPart.Document.Body;

                body.Append(CreateParagraph("Мілецький Ілля Валерійович", 28, true, JustificationValues.Center));
                body.Append(CreateParagraph("КН-212", 24, false, JustificationValues.Center));
                body.Append(CreateParagraph("37", 24, false, JustificationValues.Center));
                body.Append(CreateParagraph(
                    "Джерело датасету: https://www.kaggle.com/datasets/dnkumars/air-quality-index",
                    20, false, JustificationValues.Center));
                body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                int rowCount = data.Count;
                int colCount = 3 + 12;
                bool hasMissing = data.Any(c =>
                    string.IsNullOrEmpty(c.CityCountry) ||
                    c.MonthlyData.Any(m => m.Value == null));

                body.Append(CreateParagraph("Опис датасету", 26, true, JustificationValues.Center));
                body.Append(CreateParagraph($"Кількість рядків: {rowCount}", 22));
                body.Append(CreateParagraph($"Кількість стовпців: {colCount}", 22));
                body.Append(CreateParagraph("Поля: Rank(int), CityCountry(string), AverageAQI(int), MonthlyData(int?)", 22));
                body.Append(CreateParagraph($"Пропуски: {(hasMissing ? "є" : "немає")}", 22));

                var preview = data.Take(5).Select(c =>
                    $"{c.Rank}, {c.CityCountry}, {c.AverageAQI}, " +
                    string.Join(", ", c.MonthlyData.Select(m => m.Value.ToString() ?? "-"))
                );

                body.Append(CreateParagraph("Попередній перегляд (перші 5 рядків):", 22, true));
                foreach (var line in preview)
                {
                    body.Append(CreateParagraph(line, 20));
                }

                mainPart.Document.Save();
            }
        }

        private Paragraph CreateParagraph(string text, int fontSize = 20,
            bool bold = false, JustificationValues? align = null)
        {
            JustificationValues actualAlign = align ?? JustificationValues.Left;
            RunProperties runProps = new RunProperties(
                new FontSize() { Val = (fontSize * 2).ToString() }
            );
            if (bold) runProps.Append(new Bold());

            Run run = new Run(runProps, new Text(text));

            ParagraphProperties paraProps = new ParagraphProperties(
                new Justification() { Val = align }
            );

            return new Paragraph(paraProps, run);
        }
    }
}
