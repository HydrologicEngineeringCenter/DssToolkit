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

        private void DataTypePage_PairedDataNextClick(object sender, RoutedEventArgs e)
        {
            DataTypePage.Visibility = Visibility.Collapsed;
            OrdinatePage.Visibility = Visibility.Visible;
        }

        private void DataTypePage_TimeSeriesNextClick(object sender, RoutedEventArgs e)
        {
            DataTypePage.Visibility = Visibility.Collapsed;
            DatePage.Visibility = Visibility.Visible;
        }

        private void DatePage_NextClick(object sender, RoutedEventArgs e)
        {
            DatePage.Visibility = Visibility.Collapsed;
            TimeSeriesValuePage.Visibility = Visibility.Visible;
        }

        private void DatePage_BackClick(object sender, RoutedEventArgs e)
        {
            DatePage.Visibility = Visibility.Collapsed;
            DataTypePage.Visibility = Visibility.Visible;
        }

        private void OrdinatePage_NextClick(object sender, RoutedEventArgs e)
        {
            OrdinatePage.Visibility = Visibility.Collapsed;
            PairedDataValuePage.Visibility = Visibility.Visible;
        }

        private void OrdinatePage_BackClick(object sender, RoutedEventArgs e)
        {
            OrdinatePage.Visibility = Visibility.Collapsed;
            DataTypePage.Visibility = Visibility.Visible;
        }

        private void TimeSeriesValuePage_ImportClick(object sender, RoutedEventArgs e)
        {

        }

        private void TimeSeriesValuePage_BackClick(object sender, RoutedEventArgs e)
        {
            TimeSeriesValuePage.Visibility = Visibility.Collapsed;
            DatePage.Visibility = Visibility.Visible;
        }

        private void PairedDataValuePage_ImportClick(object sender, RoutedEventArgs e)
        {

        }

        private void PairedDataValuePage_BackClick(object sender, RoutedEventArgs e)
        {
            PairedDataValuePage.Visibility = Visibility.Collapsed;
            OrdinatePage.Visibility = Visibility.Visible;
        }
    }
}
