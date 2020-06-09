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
        public string wordpathfolder = " ";//лучше сохранять в папку экселя
        public string excelpathfolder = " ";
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
            {
                excelpath.Text = (ofd.FileName);
                excelpathfolder = System.IO.Path.GetDirectoryName(ofd.FileName);
            }
        }

        public void executebutton_Click(object sender, RoutedEventArgs e)
        {
            FindAndReplaceObject file = new FindAndReplaceObject(System.IO.Path.Combine(@wordpath.Text), System.IO.Path.Combine(@excelpath.Text), excelpathfolder);
            if (file.FindAndReplace())
                MessageBox.Show("Обработка успешно завершена");
            else
                MessageBox.Show("Во время работы программы произошла ошибка. Файлы не были обработаны");
        }

        private void leavewindow_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
