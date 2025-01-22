using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using System.IO;

namespace ExcelToWord
{
    class Replacer
    {
        private object wordFilePath;
        private object excelFilePath;
        private object outfilepathfolder;
        public Replacer(object w, object e, object p) { wordFilePath = w; excelFilePath = e; outfilepathfolder = p; }
        public bool FindAndReplace()
        {
            byte[] byteArray = File.ReadAllBytes((string)wordFilePath);
            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    using (SpreadsheetDocument excDoc = SpreadsheetDocument.Open((string)excelFilePath, true))
                    {
                        var wordBody = wordDoc.MainDocumentPart.Document.Body;
                        var paragraphs = wordBody.Elements<Paragraph>();
                        Regex markerRegEx = new Regex(@"<#\d+#[A-Z]+\d+>");

                        foreach (var paragraph in paragraphs)
                        {
                            foreach (var run in paragraph.Elements<Run>())
                            {
                                foreach (var text in run.Elements<Text>())
                                {
                                    MatchCollection markerMatches = markerRegEx.Matches(text.Text);

                                    foreach (Match match in markerMatches)
                                    {
                                        Regex sheetRegEx = new Regex(@"#\d+#");
                                        Regex cellRegEx = new Regex(@"#[A-Z]+\d+>");
                                        int sheetIndex = Int32.Parse(sheetRegEx.Match(match.Value).Value.Trim('#'));
                                        string cellIndex = cellRegEx.Match(match.Value).Value.Trim('#', '>');
                                        WorkbookPart wbPart = excDoc.WorkbookPart;
                                        Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.SheetId == sheetIndex);
                                        WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                                        Cell cell = wsPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellIndex);

                                        var value = cell.InnerText;

                                        if (!(cell.DataType is null))
                                        {
                                            if (cell.DataType.Value == CellValues.SharedString)
                                            {
                                                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;

                                                text.Text = text.Text.Replace(match.Value, value);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    wordDoc.Save();
                }

                stream.Position = 0;
                File.WriteAllBytes($@"{outfilepathfolder}\out_file.docx", stream.ToArray());

                return true;
            }
            return false;
        }
    }
}
