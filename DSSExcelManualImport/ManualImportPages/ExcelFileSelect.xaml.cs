using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Shapes;

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for ExcelFileSelect.xaml
    /// </summary>
    public partial class ExcelFileSelect : Window
    {
        public string FileName { get; set; }
        public ExcelFileSelect()
        {
            InitializeComponent();
        }

        public event RoutedEventHandler NextClick;
        public event RoutedEventHandler BackClick;

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            NextClick?.Invoke(this, e);
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            BackClick?.Invoke(this, e);
        }

        private void FileSelectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select or Create Excel File";
            //dialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx|CSV Files (*.csv)|*.csv";
            dialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (dialog.ShowDialog() == true)
            {
                FileName = dialog.FileName;
                FileNameTextBox.Text = FileName;
            }
        }
    }
}
