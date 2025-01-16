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
            using (WordprocessingDocument wordDoc = WordprocessingDocument.CreateFromTemplate((string)wordFilePath))
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
                                    int sheet = Int32.Parse(sheetRegEx.Match(match.Value).Value.Trim('#'));
                                    string cell = cellRegEx.Match(match.Value).Value.Trim('#', '>');
                                    string relationshipId = excDoc.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.SheetId.Equals(sheet))?.Id;

                                    text.Text.Replace(match.Value, cell);
                                }
                            }
                        }
                    }
                }
                return true;
            }
            return false;
        }
    }
}
