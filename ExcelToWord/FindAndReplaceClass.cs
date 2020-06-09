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
    class FindAndReplaceClass
    {
        public static void FindAndReplace(Word.Document doc, object findText, object replaceWithText)
        {
            Word.Range rng = doc.Content;
            rng.Find.ClearFormatting();
            //options
            /*object matchCase = false;
            object matchWholeWord = false;
            object matchWildCards = true;
            object matchSoundsLike = false;                                     // с этими опциями не заменялось
            object matchAllWordForms = false;
            object forward = true;
            object format = false;*/
            object missing = Type.Missing;
            object read_only = false;
            object visible = true;
            object wrap = 1;
            //execute find and replace
            //object smth = doc.Selection.Find;
            rng.Find.Execute(ref findText, missing, missing, missing, missing, missing, missing, missing, missing, ref replaceWithText, 2);
        }

        /*public static string Find(Word.Application doc, object findText)
        {
            doc.Selection.Find(findText);
            return;
        }*/
    }
}
