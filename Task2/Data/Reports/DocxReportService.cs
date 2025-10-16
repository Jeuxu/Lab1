using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Domain;
using System.Diagnostics;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Data.Reports
{
    public class DocxReportService
    {
        public void GenerateReport(string outputPath, List<CityAirQuality> data, string chartPath)
        {
            using (var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body());
                var body = main.Document.Body;

                var titleStyle = new RunProperties(
                    new RunFonts { Ascii = "Segoe UI" },
                    new Bold(),
                    new FontSize { Val = "30" }
                );

                var subTitleStyle = new RunProperties(
                    new RunFonts { Ascii = "Segoe UI" },
                    new FontSize { Val = "26" }
                );

                var textStyle = new RunProperties(
                    new RunFonts { Ascii = "Segoe UI" },
                    new FontSize { Val = "22" }
                );

                body.Append(new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new SpacingBetweenLines { After = "200" }
                    ),
                    new Run(titleStyle, new Text("ЗВІТ ПРО ЯКІСТЬ ПОВІТРЯ"))
                ));

                var info = new[]
                {
                    "Студент: Мілецький Ілля Валерійович",
                    "Варіант: 37",
                    "Джерело: https://www.kaggle.com/datasets/dnkumars/air-quality-index",
                    $"Дата: {DateTime.Now:dd.MM.yyyy}"
                };

                foreach (var line in info)
                {
                    body.Append(new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Left },
                            new SpacingBetweenLines { After = "100" }
                        ),
                        new Run(textStyle.CloneNode(true), new Text(line))
                    ));
                }
                body.Append(new Paragraph(new Run(new Break())));

                body.Append(new Paragraph(textStyle.CloneNode(true), new Run(subTitleStyle.CloneNode(true), new Text("Таблиця з даними за перші 20 міст."))));
                var table = new Table(new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 6 },
                        new BottomBorder { Val = BorderValues.Single, Size = 6 },
                        new LeftBorder { Val = BorderValues.Single, Size = 6 },
                        new RightBorder { Val = BorderValues.Single, Size = 6 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
                    )));

                var header = new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("Rank")))),
                    new TableCell(new Paragraph(new Run(new Text("City / Country")))),
                    new TableCell(new Paragraph(new Run(new Text("Average AQI"))))
                );
                table.Append(header);

                foreach (var city in data.Take(20))
                {
                    if (city == null) continue;

                    table.Append(new TableRow(
                        new TableCell(new Paragraph(new Run(new Text(city.Rank.ToString())))),
                        new TableCell(new Paragraph(new Run(new Text(city.CityCountry ?? "")))),
                        new TableCell(new Paragraph(new Run(new Text(city.AverageAQI.ToString()))))
                    ));
                }

                body.Append(table);

                body.Append(new Paragraph(new Run(subTitleStyle.CloneNode(true), new Text("Кругова діаграма"))));
                if (File.Exists(chartPath))
                {
                    var imgPart = main.AddImagePart(ImagePartType.Png);
                    using (var fs = new FileStream(chartPath, FileMode.Open))
                        imgPart.FeedData(fs);

                    string relId = main.GetIdOfPart(imgPart);

                    var drawing = new Drawing(
                        new DW.Inline(
                            new DW.Extent() { Cx = 6000000L, Cy = 3500000L },
                            new DW.EffectExtent(),
                            new DW.DocProperties() { Id = 1, Name = $"Chart" },
                            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                            new A.Graphic(
                                new A.GraphicData(
                                    new PIC.Picture(
                                        new PIC.NonVisualPictureProperties(
                                            new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "ChartImage.png" },
                                            new PIC.NonVisualPictureDrawingProperties()),
                                        new PIC.BlipFill(
                                            new A.Blip() { Embed = relId },
                                            new A.Stretch(new A.FillRectangle())),
                                        new PIC.ShapeProperties(
                                            new A.Transform2D(
                                                new A.Offset() { X = 0L, Y = 0L },
                                                new A.Extents() { Cx = 6000000L, Cy = 3500000L }),
                                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                                    )
                                )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                        )
                    );

                    body.Append(new Paragraph(new Run(drawing)));
                }
                body.Append(new Paragraph(new Run(new Break())));

                int totalCities = data.Count;
                var bestCity = data.OrderBy(c => c.AverageAQI).FirstOrDefault();
                var worstCity = data.OrderByDescending(c => c.AverageAQI).FirstOrDefault();
                double avgAqi = data.Average(c => c.AverageAQI);
                double medianAqi = data.Select(c => c.AverageAQI).OrderBy(v => v)
                    .ElementAt(totalCities / 2);
                double minAqi = data.Min(c => c.AverageAQI);
                double maxAqi = data.Max(c => c.AverageAQI);

                int goodCount = data.Count(c => c.AverageAQI < 50);
                int moderateCount = data.Count(c => c.AverageAQI is >= 50 and < 100);
                int badCount = data.Count(c => c.AverageAQI >= 100);

                var topBest = string.Join(", ", data.OrderBy(c => c.AverageAQI)
                    .Take(5)
                    .Select(c => $"{c.CityCountry} ({c.AverageAQI:F1})"));

                var topWorst = string.Join(", ", data.OrderByDescending(c => c.AverageAQI)
                    .Take(5)
                    .Select(c => $"{c.CityCountry} ({c.AverageAQI:F1})"));


                body.Append(new Paragraph(new Run(subTitleStyle.CloneNode(true), new Text("Висновки"))));
                body.Append(new Paragraph(new Run(textStyle.CloneNode(true), new Text($"У наборі даних міститься інформація про {totalCities} міст. " +
    $"Середній індекс якості повітря (AQI) становить {avgAqi:F1}, медіана — {medianAqi:F1}. " +
    $"Діапазон значень AQI: від {minAqi:F1} до {maxAqi:F1}. " +
    $"Найкращу якість повітря має {bestCity?.CityCountry ?? "—"} (AQI {bestCity?.AverageAQI:F1}), " +
    $"а найгіршу — {worstCity?.CityCountry ?? "—"} (AQI {worstCity?.AverageAQI:F1}). " +
    $"До категорії міст з хорошим повітрям належить {goodCount} міст, з помірним забрудненням — {moderateCount}, " +
    $"з поганим повітрям — {badCount}. " +
    $"Найчистіші міста: {topBest}. Найзабрудненіші: {topWorst}. "))));

                main.Document.Save();
            }
        }
    }
}