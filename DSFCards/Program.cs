using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace DSFCards
{
    class ScoreCardEntry
    {
        public int Index { get; set; }
        public int PersonId { get; set; }
        public string EventName { get; set; }
        public string GroupNo { get; set; }
        public int StationNo { get; set; }
    }

    class CompCardEntry
    {
        public int Index { get; set; }
        public int PersonId { get; set; }
        public string[] EventList { get; set; }
    }

    class Program
    {
        static string AppVersion = "0.1";

        static void Main(string[] args)
        {
            Console.WriteLine($"DSF Cards - version {AppVersion}");
            Console.WriteLine("(C) 2020 Thor Muto Asmund");
            Console.WriteLine();

            if (args.Length < 2)
            {
                Console.WriteLine("USAGE: DSFCards.exe <SCORE_CARD_PDF> <COMP_CARD_PDF>");
                return;
            }
            var scoreCardInputFileName = args[0];
            var compCardInputFileName = args[1];
            var isDebug = args.Length > 2 && args[2] == "DEBUG";

            var scoreCardOutputFileName = scoreCardInputFileName.Replace(".pdf", "_out.pdf");
            var dataOutputFileName = scoreCardInputFileName.Replace(".pdf", "_data.csv");
            var compCardOutputFileName = compCardInputFileName.Replace(".pdf", "_out.pdf");

            var scoreCardText = PdfToText(scoreCardInputFileName);
            if (isDebug)
            {
                File.WriteAllText("scorecards_raw.txt", scoreCardText);
            }
            var compCardText = PdfToText(compCardInputFileName);
            if (isDebug)
            {
                File.WriteAllText("compcards_raw.txt", compCardText);
            }

            // Read data
            var scoreCardEntries = new List<ScoreCardEntry>();
            var compCardEntries = new List<CompCardEntry>();
            ReadScoreCardData(scoreCardText, scoreCardEntries, isDebug);
            ReadCompCardData(compCardText, compCardEntries, isDebug);
            
            // Create pdfs
            CreateScoreCards(scoreCardInputFileName, scoreCardOutputFileName, scoreCardEntries);
            CreateCompCards(compCardInputFileName, compCardOutputFileName, scoreCardEntries, compCardEntries);

            Console.WriteLine("Finished!");

            if (isDebug)
            {
                Console.ReadKey();
            }
        }

        static string PdfToText(string inputFileName)
        {
            var sw = new StringWriter();
            using (PdfReader reader = new PdfReader(inputFileName))
            {
                using (PdfDocument pdf = new PdfDocument(reader))
                {
                    for (int pageNo = 1; pageNo <= pdf.GetNumberOfPages(); pageNo++)
                    {
                        var page = pdf.GetPage(pageNo);
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string currentText = PdfTextExtractor.GetTextFromPage(page, strategy);
                        sw.WriteLine(currentText);
                    }
                }
            }

            return sw.ToString();
        }

        static void CreateScoreCards(string inputFileName, string outputFileName, List<ScoreCardEntry> entries)
        { 
            var font = PdfFontFactory.CreateFont(iText.IO.Font.FontConstants.HELVETICA);

            var p = 0;
            using (var reader = new PdfReader(inputFileName))
            {
                using (var inputPdf = new PdfDocument(reader))
                {
                    using (var writer = new PdfWriter(outputFileName))
                    {
                        using (var outputPdf = new PdfDocument(writer))
                        {
                            for (int pageNo = 1; pageNo <= inputPdf.GetNumberOfPages(); pageNo++)
                            {
                                var inputPage = inputPdf.GetPage(pageNo);
                                var outputPage = inputPage.CopyTo(outputPdf);

                                var mediaBox = outputPage.GetMediaBox();
                                var canvas = new Canvas(new PdfCanvas(outputPage, true), mediaBox);

                                for (int i = 0; i < 4; ++i)
                                {
                                    var entry = entries.FirstOrDefault(e => e.Index == p);

                                    if (entry != null)
                                    {
                                        var x = mediaBox.GetWidth() - 25f;
                                        if (i % 2 == 0)
                                        {
                                            x -= (mediaBox.GetWidth() / 2f);
                                        }
                                        var y = mediaBox.GetHeight() - 30f;
                                        if (i / 2 >= 1)
                                        {
                                            y -= (mediaBox.GetHeight() / 2f);
                                        }
                                        var text = $"S{entry.StationNo}";
                                        canvas.ShowTextAligned(new Paragraph(text).SetFont(font).SetFontSize(11), x, y, TextAlignment.RIGHT);
                                    }
                                    p++;
                                }
                                canvas.Close();

                                outputPdf.AddPage(outputPage);
                            }
                        }
                    }
                }
            }
        }

        static void CreateCompCards(string inputFileName, string outputFileName, List<ScoreCardEntry> scoreCardEntries, List<CompCardEntry> compCardEntries)
        {
            var font = PdfFontFactory.CreateFont(iText.IO.Font.FontConstants.HELVETICA);

            var p = 0;
            using (var reader = new PdfReader(inputFileName))
            {
                using (var inputPdf = new PdfDocument(reader))
                {
                    using (var writer = new PdfWriter(outputFileName))
                    {
                        using (var outputPdf = new PdfDocument(writer))
                        {
                            for (int pageNo = 1; pageNo <= inputPdf.GetNumberOfPages(); pageNo++)
                            {
                                var inputPage = inputPdf.GetPage(pageNo);
                                var outputPage = inputPage.CopyTo(outputPdf);

                                var mediaBox = outputPage.GetMediaBox();
                                var canvas = new Canvas(new PdfCanvas(outputPage, true), mediaBox);

                                for (int i = 0; i < 12; ++i)
                                {
                                    var compCardEntry = compCardEntries.FirstOrDefault(e => e.Index == p);

                                    if (compCardEntry != null)
                                    {
                                        var x = mediaBox.GetWidth() + 49f;
                                        x -= (mediaBox.GetWidth() / 3f - 3.5f) * (3 - (i % 3));
                                        var y = mediaBox.GetHeight() - 55f;
                                        y -= (mediaBox.GetHeight() / 4f - 38.5f) * (i / 3);

                                        foreach (var eventName in compCardEntry.EventList)
                                        {
                                            var scoreCardEntry = scoreCardEntries.FirstOrDefault(e => e.PersonId == compCardEntry.PersonId && e.EventName == eventName);
                                            if (scoreCardEntry != null)
                                            {
                                                var text = $"S{scoreCardEntry.StationNo}";
                                                canvas.ShowTextAligned(new Paragraph(text).SetFont(font).SetFontSize(8), x, y, TextAlignment.LEFT);
                                            }
                                            y -= 12.5f;
                                        }
                                    }
                                    p++;
                                }
                                canvas.Close();

                                outputPdf.AddPage(outputPage);
                            }
                        }
                    }
                }
            }
        }

        static void ReadScoreCardData(string text, List<ScoreCardEntry> entries, bool isDebug)
        {
            var output = new StringWriter();
            output.WriteLine("index\tpersonId\teventName\tgroup\tstationNo");

            var lines = Regex.Split(text, "\r\n|\r|\n");
            var lineNo = -1;
            var stationNo = 0;
            var index = -1;
            string currentGroupNo = "";
            string currentEventName = "";
            do
            {
                int serialId = 0;
                for (; ; )
                {
                    lineNo++;
                    if (lineNo >= lines.Length)
                    {
                        break;
                    }
                    if (Int32.TryParse(lines[lineNo].Trim(), out serialId))
                    {
                        break;
                    }
                }

                lineNo += 3;
                if (lineNo >= lines.Length)
                {
                    break;
                }

                // Event line
                var eventLine = lines[lineNo].Trim();
                var eventData = Regex.Split(eventLine, " ").ToArray();
                var eventName = string.Join(" ", eventData.Take(eventData.Length - 2));
                var groupNo = eventData[eventData.Length - 1];
                if (groupNo != currentGroupNo || eventName != currentEventName)
                {
                    stationNo = 0;
                    var pageBreakOffset = 4 - ((index + 1) % 4);
                    if (pageBreakOffset < 4)
                    {
                        index += pageBreakOffset;
                    }
                }
                currentGroupNo = groupNo;
                currentEventName = eventName;

                // Person id line
                lineNo+=2;
                var personLine = lines[lineNo].Trim();
                var personData = Regex.Split(personLine, " ").ToArray();
                var personId = Int32.Parse(personData[0]);

                // Add entry
                stationNo++;
                index++;

                output.WriteLine($"{index}\t{personId}\t{eventName}\t{groupNo}\t{stationNo}");
                entries.Add(new ScoreCardEntry()
                {
                    Index = index,
                    PersonId = personId,
                    EventName = eventName,
                    GroupNo = groupNo,
                    StationNo = stationNo
                });

                for (; ; )
                {
                    lineNo++;
                    if (lineNo >= lines.Length || lines[lineNo].Trim() == "_")
                    {
                        break;
                    }
                }
            }
            while (lineNo < lines.Length);

            if (isDebug)
            {
                File.WriteAllText("scorecards_data.txt", output.ToString());
            }
        }

        static void ReadCompCardData(string text, List<CompCardEntry> entries, bool isDebug)
        {
            var output = new StringWriter();
            output.WriteLine("index\tpersonId\tevents");

            var lines = Regex.Split(text, "\r\n|\r|\n");
            var lineNo = -1;
            var index = -1;
            do
            {
                int personId = 0;
                for (; ; )
                {
                    lineNo++;
                    if (lineNo >= lines.Length)
                    {
                        break;
                    }
                    if (lines[lineNo].StartsWith("ID"))
                    {
                        var idData = Regex.Split(lines[lineNo].Trim(), " ").ToArray();
                        personId = Int32.Parse(idData[1]);
                        break;
                    }
                }

                lineNo += 2;

                var events = new List<string>();
                while (lineNo < lines.Length)
                {
                    var eventLine = lines[lineNo].Trim();
                    if (eventLine.Length == 0)
                    {
                        break;
                    }
                    var eventData = Regex.Split(eventLine, " ").Where(e =>
                    {
                        if (e.EndsWith(","))
                        {
                            e = e.Substring(0, e.Length - 1);
                        }

                        return !Int32.TryParse(e, out _);
                    }).ToArray();

                    events.Add(string.Join(" ", eventData));
                    lineNo++;
                }

                index++;

                output.WriteLine($"{index}\t{personId}\t{string.Join("|",events)}");
                entries.Add(new CompCardEntry()
                {
                    Index = index,
                    PersonId = personId,
                    EventList = events.ToArray()
                });

                while (lineNo < lines.Length && lines[lineNo].Trim().Length == 0)
                {
                    lineNo++;
                }
            }
            while (lineNo < lines.Length);

            if (isDebug)
            {
                File.WriteAllText("compcards_data.txt", output.ToString());
            }
        }
    }
}
