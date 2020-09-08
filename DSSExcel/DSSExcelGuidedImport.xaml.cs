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
using System.Windows.Shapes;

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for SelectDataType.xaml
    /// </summary>
    public partial class DSSExcelGuidedImport : Window
    {
        public DSSExcelGuidedImport()
        {
            InitializeComponent();
        }

        private void SelectDataTypePage_Click(object sender, RoutedEventArgs e)
        {
            DataTypePage.Visibility = Visibility.Collapsed;
            DatePage.Visibility = Visibility.Visible;
        }
    }
}
