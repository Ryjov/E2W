using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Windows;

namespace ExcelToWord
{
    class FindAndReplaceObject
    {
        private object wordFilePath;
        private object excelFilePath;
        private object outfilepathfolder;
        public FindAndReplaceObject (object w, object e, object p) { wordFilePath = w; excelFilePath = e; outfilepathfolder = p; }
        public bool FindAndReplace()
        {
            Word.Application wordApp = new Word.Application { Visible = false };
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            Word.Document wordDoc = wordApp.Documents.Open(wordFilePath, ReadOnly: false, Visible: false);
            //wordDoc.Activate();
            Regex markerRegEx = new Regex(@"<#\d+#[A-Z]+\d+>");
            string rangeText = wordDoc.Range().Text;
            MatchCollection markerMatches = markerRegEx.Matches(rangeText);
            Excel.Application excApp = new Excel.Application { Visible = false };
            excApp.DisplayAlerts = false;
            Excel.Workbook excBook = excApp.Workbooks.Add(excelFilePath);
            object missing = Type.Missing;
            object read_only = false;
            object visible = true;
            object wrap = 1;
            Word.Range rng = wordDoc.Content;
            rng.Find.ClearFormatting();
            try
            {
                foreach (Match match in markerMatches)
                {
                    Regex sheetRegEx = new Regex(@"#\d+#");
                    Regex cellRegEx = new Regex(@"#[A-Z]+\d+>");
                    int sheet = Int32.Parse(sheetRegEx.Match(match.Value).Value.Trim('#'));
                    Excel.Worksheet excSheet = (excBook.Worksheets[sheet]);
                    string cell = cellRegEx.Match(match.Value).Value.Trim('#', '>');
                    Excel.Range excRng;
                    excRng = excSheet.get_Range(cell);
                    rng.Find.Execute(match.Value, missing, missing, missing, missing, missing, missing, missing, missing, excRng.Value2, 2);
                }
                //aDoc.SaveAs2(@"ExcelToWordfile.docx");
                //wordDoc.SaveAs2("C:\\Users\\Егор\\Desktop\\диплом\\test\\doc1.doc");
                wordDoc.SaveAs2($@"{outfilepathfolder}\out_file.doc");
                wordDoc.Close();
                wordDoc = null;
                wordApp.Quit();
                wordApp = null;
                excBook.Close(0);
                excApp.Quit();
                return true;
            }
            catch
            {
                //wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                //excApp.DisplayAlerts = false;
                wordDoc.Close();
                wordDoc = null;
                wordApp.Quit();
                wordApp = null;
                excBook.Close(0);
                excApp.Quit();
                return false;
            }
        }
    }
}
