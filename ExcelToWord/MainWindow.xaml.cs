using System;
using System.Windows;
using Microsoft.Win32;
using Forms=System.Windows.Forms;

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
        bool availabilityTop = false;
        bool availabilityBottom = false;
        bool availabilityTotal = false;
        OpenFileDialog ofd = new OpenFileDialog();
        Forms.FolderBrowserDialog fbd = new Forms.FolderBrowserDialog();
        
        public string wordpathfolder = " ";//лучше сохранять в папку экселя
        public string excelpathfolder = " ";
        private void wordpathbutton_Click(object sender, RoutedEventArgs e)
        {
            ofd.Filter = "Word Documents|* .doc; *.docx";
            if (ofd.ShowDialog() == true)
            {
                wordpath.Text = (ofd.FileName);
                wordpathfolder = System.IO.Path.GetDirectoryName(ofd.FileName);
                availabilityTop = true;
                checkTextbox(availabilityBottom);
            }
        }

        private void excelpathbutton_Click(object sender, RoutedEventArgs e)
        {
            ofd.Filter = "Excel Worksheets| *.xls; *.xlsx";
            if (ofd.ShowDialog() == true)
            {
                excelpath.Text = (ofd.FileName);
                excelpathfolder = System.IO.Path.GetDirectoryName(ofd.FileName);
                availabilityBottom = true;
                checkTextbox(availabilityTop);
            }
        }

        public void executebutton_Click(object sender, RoutedEventArgs e)
        {
            if (ExcelRadio.IsChecked==true)
            {
                Replacer file = new Replacer(System.IO.Path.Combine(@wordpath.Text), System.IO.Path.Combine(@excelpath.Text), excelpathfolder);
                if (file.FindAndReplace())
                    MessageBox.Show("Обработка успешно завершена");
                else
                    MessageBox.Show("Во время работы программы произошла ошибка. Файлы не были обработаны");
            }
            else if (WordRadio.IsChecked==true)
            {
                Replacer file = new Replacer(System.IO.Path.Combine(@wordpath.Text), System.IO.Path.Combine(@excelpath.Text), wordpathfolder);
                if (file.FindAndReplace())
                    MessageBox.Show("Обработка успешно завершена");
                else
                    MessageBox.Show("Во время работы программы произошла ошибка. Файлы не были обработаны");
            }
            else if ((PathRadio.IsChecked==true)&&(!(String.IsNullOrEmpty(OutfilePathText.Text))))
            {
                Replacer file = new Replacer(System.IO.Path.Combine(@wordpath.Text), System.IO.Path.Combine(@excelpath.Text), OutfilePathText.Text);
                if (file.FindAndReplace())
                    MessageBox.Show("Обработка успешно завершена");
                else
                    MessageBox.Show("Во время работы программы произошла ошибка. Файлы не были обработаны");
            }
        }

        public void checkTextbox (bool checkVariable)
        {
            if (checkVariable)
            {
                executebutton.IsEnabled = true;
            }
            else
            {
                executebutton.IsEnabled = false;
            }
        }

        public void PathRadioChecked(object sender, RoutedEventArgs e)
        {
            OutfilePathText.IsEnabled = true;
        }

        public void PathRadioUnchecked(object sender, RoutedEventArgs e)
        {
            OutfilePathText.IsEnabled = false;
        }

        private void OutfilePathButton_Click(object sender, RoutedEventArgs e)
        {
            PathRadio.IsChecked = true;
            OutfilePathText.IsEnabled = true;
            if (fbd.ShowDialog()==Forms.DialogResult.OK)
            {
                OutfilePathText.Text = fbd.SelectedPath;
            }
        }
    }
}
