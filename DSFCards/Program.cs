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
using System.Globalization;
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
        static string AppVersion = "0.2";

        static void Main(string[] args)
        {
            var scoreCardInputFileName = "";
            var compCardInputFileName = "";
            var compCardRowsPerPage = 4;
            var compCardColumnsPerPage = 3;
            float? compCardLeftMargin = null;
            float? compCardTopMargin = null;
            string overrideScoreCardOutputFileName = null;
            string overrideCompCardOutputFileName = null;

            Console.WriteLine($"DSF Cards - version {AppVersion}");
            Console.WriteLine("(C) 2020 Thor Muto Asmund");
            Console.WriteLine();

            if (args.Length < 2 || (args.Length > 0 && args[0] == "/?"))
            {
                Console.WriteLine("DSFCards.exe scorecard_pdf compcard_pdf [/CR] [/CC] [/CLM] [/CTM] [/SO] [/CO]");
                Console.WriteLine();
                Console.WriteLine($"  scorecard_pdf  PDF file with score cards");
                Console.WriteLine($"  compcard_pdf   PDF file with comp cards");
                Console.WriteLine($"  /CR            Comp card rows per page (default {compCardRowsPerPage})");
                Console.WriteLine($"  /CC            Comp card columns per page (default {compCardColumnsPerPage})");
                Console.WriteLine($"  /CLM           Comp card left margin");
                Console.WriteLine($"  /CTM           Comp card top margin");
                Console.WriteLine($"  /SO            Score card output file");
                Console.WriteLine($"  /CO            Comp card output file");
                Environment.Exit(0);
            }

            var arg = 0;
            var noarg = 0;
            while (arg < args.Length)
            {
                switch (args[arg].ToUpper())
                {
                    case "/CR":
                        arg++;
                        if (!Int32.TryParse(args[arg], out compCardRowsPerPage))
                        {
                            Console.WriteLine("Error parsing CR argument");
                            Environment.Exit(1);
                        }
                        arg++;
                        break;
                    case "/CC":
                        arg++;
                        if (!Int32.TryParse(args[arg], out compCardColumnsPerPage))
                        {
                            Console.WriteLine("Error parsing CC argument");
                            Environment.Exit(1);
                        }
                        arg++;
                        break;
                    case "/CLM":
                        arg++;
                        if (!float.TryParse(args[arg], System.Globalization.NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var readLeftMargin))
                        {
                            Console.WriteLine("Error parsing CLM argument");
                            Environment.Exit(1);
                        }
                        compCardLeftMargin = readLeftMargin;
                        arg++;
                        break;
                    case "/CTM":
                        arg++;
                        if (!float.TryParse(args[arg], System.Globalization.NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var readTopMargin))
                        {
                            Console.WriteLine("Error parsing CTM argument");
                            Environment.Exit(1);
                        }
                        compCardTopMargin = readTopMargin;
                        arg++;
                        break;
                    case "/SO":
                        arg++;
                        overrideScoreCardOutputFileName = args[arg];
                        arg++;
                        break;
                    case "/CO":
                        arg++;
                        overrideCompCardOutputFileName = args[arg];
                        arg++;
                        break;
                    default:
                        switch (noarg)
                        {
                            case 0:
                                scoreCardInputFileName = args[arg];
                                noarg++;
                                break;
                            case 1:
                                compCardInputFileName = args[arg];
                                noarg++;
                                break;
                            default:
                                Console.WriteLine("Too many arguments");
                                Environment.Exit(1);
                                break;
                        }
                        arg++;
                        break;
                }
            }

            if (noarg < 2)
            {
                Console.WriteLine("Missing arguments");
                Environment.Exit(1);
            }

            if (!File.Exists(scoreCardInputFileName))
            {
                Console.WriteLine($"File not found: {scoreCardInputFileName}");
                Environment.Exit(1);
            }
            if (!File.Exists(compCardInputFileName))
            {
                Console.WriteLine($"File not found: {compCardInputFileName}");
                Environment.Exit(1);
            }

            var isDebug = args.Length > 2 && args[2] == "DEBUG";

            var scoreCardOutputFileName = overrideScoreCardOutputFileName ?? scoreCardInputFileName.Replace(".pdf", "_out.pdf");
            var compCardOutputFileName = overrideCompCardOutputFileName ?? compCardInputFileName.Replace(".pdf", "_out.pdf");
            
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
            CreateCompCards(compCardInputFileName, compCardOutputFileName, scoreCardEntries, compCardEntries, 
                compCardRowsPerPage, compCardColumnsPerPage, compCardLeftMargin, compCardTopMargin);

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

        static void CreateCompCards(string inputFileName, string outputFileName, List<ScoreCardEntry> scoreCardEntries, List<CompCardEntry> compCardEntries,
            int compCardRowsPerPage, int compCardColumnsPerPage, float? compCardLeftMargin, float? compCardTopMargin)
        {
            // 59f 55f
            var font = PdfFontFactory.CreateFont(iText.IO.Font.FontConstants.HELVETICA);
            var originalFontSize = 7.75f;
            var fontSize = 5f;

            var p = 0;
            using (var reader = new PdfReader(inputFileName))
            {
                using (var inputPdf = new PdfDocument(reader))
                {
                    using (var writer = new PdfWriter(outputFileName))
                    {
                        using (var outputPdf = new PdfDocument(writer))
                        {
                            // Calculate top margin
                            if (!compCardTopMargin.HasValue)
                            {
                                compCardTopMargin = 55f;
                            }

                            // Calculate left margin
                            if (!compCardLeftMargin.HasValue)
                            {
                                var calculatedLeftMargin = 0f;
                                var eventNames = scoreCardEntries.Select(s => s.EventName).Distinct();
                                foreach (var eventName in eventNames)
                                {
                                    var textWidth = font.GetWidth(eventName, originalFontSize);
                                    if (textWidth > calculatedLeftMargin)
                                    {
                                        calculatedLeftMargin = textWidth;
                                    }
                                }
                                compCardLeftMargin = calculatedLeftMargin;
                            }

                            // Process pages
                            for (int pageNo = 1; pageNo <= inputPdf.GetNumberOfPages(); pageNo++)
                            {
                                var inputPage = inputPdf.GetPage(pageNo);
                                var outputPage = inputPage.CopyTo(outputPdf);

                                var mediaBox = outputPage.GetMediaBox();
                                var canvas = new Canvas(new PdfCanvas(outputPage, true), mediaBox);

                                for (int i = 0; i < compCardRowsPerPage* compCardColumnsPerPage; ++i)
                                {
                                    var compCardEntry = compCardEntries.FirstOrDefault(e => e.Index == p);

                                    if (compCardEntry != null)
                                    {
                                        var x = 17.5f + compCardLeftMargin.Value;
                                        x += (mediaBox.GetWidth() / compCardColumnsPerPage - 3.5f) * (i % compCardColumnsPerPage);
                                        var y = mediaBox.GetHeight() - compCardTopMargin.Value;
                                        y -= (mediaBox.GetHeight() / compCardRowsPerPage - 38.5f) * (i / compCardColumnsPerPage);

                                        foreach (var eventName in compCardEntry.EventList)
                                        {
                                            var scoreCardEntry = scoreCardEntries.FirstOrDefault(e => e.PersonId == compCardEntry.PersonId && e.EventName == eventName);
                                            if (scoreCardEntry != null)
                                            {
                                                var text = $"S{scoreCardEntry.StationNo}";
                                                canvas.ShowTextAligned(new Paragraph(text).SetFont(font).SetFontSize(fontSize), x, y, TextAlignment.LEFT);
                                                //canvas.ShowTextAligned(new Paragraph("BOB").SetFont(font).SetFontSize(fontSize), x+ calculatedLeftMargin, y, TextAlignment.LEFT);
                                                
                                            }
                                            y -= 12.4f;
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
                if (eventData.Length < 3)
                {
                    Console.WriteLine($"Error in event data line: {eventLine}");
                    break;
                }

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
                lineNo += 2;
                if (lineNo >= lines.Length)
                {
                    break;
                }

                var personLine = lines[lineNo].Trim();
                var personData = Regex.Split(personLine, " ").ToArray();
                if (personData.Length < 1)
                {
                    Console.WriteLine($"Error in person data line: {personLine}");
                    break;
                }

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
                        var idLine = lines[lineNo].Trim();
                        var idData = Regex.Split(idLine, " ").ToArray();
                        if (idData.Length < 2)
                        {
                            Console.WriteLine($"Error in id data line: {idLine}");
                            break;
                        }

                        if (!Int32.TryParse(idData[1], out personId))
                        {
                            Console.WriteLine($"Error parsing id: {idData[1]}");
                            break;
                        }
                        break;
                    }
                }

                if (lineNo >= lines.Length)
                {
                    break;
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
                    else if(eventLine.StartsWith("ID"))
                    {
                        events.RemoveAt(events.Count - 1);
                        lineNo--;
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
            }
            while (lineNo < lines.Length);

            if (isDebug)
            {
                File.WriteAllText("compcards_data.txt", output.ToString());
            }
        }
    }
}
