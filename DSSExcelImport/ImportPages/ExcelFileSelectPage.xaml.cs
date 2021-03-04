using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for ExcelFileSelectPage.xaml
    /// </summary>
    public partial class ExcelFileSelectPage : UserControl
    {
        public object NextPage { get; set; }
        public string FileName { get; set; }
        public ExcelFileSelectPage()
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
                NextButton.IsEnabled = true;
            }
        }
    }
}
