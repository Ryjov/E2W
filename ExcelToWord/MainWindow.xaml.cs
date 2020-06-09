using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Microsoft.Win32;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Word=Microsoft.Office.Interop.Word;
using Excel=Microsoft.Office.Interop.Excel;

namespace ExcelToWord
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        int availability = 0;
        OpenFileDialog ofd = new OpenFileDialog();
        String markertosearch = ("([\\<])#([0-9]*)#([A-Z]*[0-9]*)([\\>])");
        public string wordpathfolder = " ";//лучше сохранять в папку экселя
        private void wordpathbutton_Click(object sender, RoutedEventArgs e)
        {
            ofd.Filter = "Word Documents|* .doc; *.docx";
            if (ofd.ShowDialog() == true)
            {
                wordpath.Text = (ofd.FileName);
                wordpathfolder = System.IO.Path.GetDirectoryName(ofd.FileName);
            }
        }

        private void excelpathbutton_Click(object sender, RoutedEventArgs e)
        {
            ofd.Filter = "Excel Worksheets| *.xls; *.xlsx";
            if (ofd.ShowDialog() == true)
                excelpath.Text = (ofd.FileName);
        }

        public void executebutton_Click(object sender, RoutedEventArgs e)
        {
            object fileName = System.IO.Path.Combine(@wordpath.Text);
            Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            Word.Document wordDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
            wordDoc.Activate();
            Regex markerRegEx = new Regex(@"<#\d+#[A-Z]+\d+>");
            string rangeText = wordDoc.Range().Text;
            MatchCollection markerMatches = markerRegEx.Matches(rangeText);
            Excel.Application excApp = new Microsoft.Office.Interop.Excel.Application { Visible = false };
            excApp.DisplayAlerts = false;
            Excel.Workbook excBook = excApp.Workbooks.Add(@excelpath.Text);
            try
            {
                foreach (Match match in markerMatches)
                {
                    Regex sheetRegEx = new Regex(@"#\d+#");
                    Regex cellRegEx = new Regex(@"#[A-Z]+\d+>");
                    int sheet = Int32.Parse(sheetRegEx.Match(match.Value).Value.Trim('#'));
                    Excel.Worksheet excSheet = (excBook.Worksheets[sheet]);
                    string cell = cellRegEx.Match(match.Value).Value.Trim('#','>');
                    Excel.Range excRng;
                    excRng = excSheet.get_Range(cell);
                    FindAndReplaceClass.FindAndReplace(wordDoc, match.Value, excRng.Value2);
                }
                //aDoc.SaveAs2(@"ExcelToWordfile.docx");
                wordDoc.SaveAs2("C:\\Users\\Егор\\Desktop\\диплом\\test\\doc1.doc");
                wordDoc.Close();
                wordDoc = null;
                wordApp.Quit();
                wordApp = null;
                excBook.Close(0);
                excApp.Quit();
                eventlog.Text = eventlog.Text + "\n Завершено успешно";
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
                eventlog.Text = eventlog.Text + "\n Произошла ошибка";
            }
        }

        private void leavewindow_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
