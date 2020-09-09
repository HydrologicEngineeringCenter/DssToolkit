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

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for SelectDataType.xaml
    /// </summary>
    public partial class SelectDataTypePage : UserControl
    {
        public SelectDataTypePage()
        {
            InitializeComponent();
        }

        private void TimeSeriesOption_Selected(object sender, RoutedEventArgs e)
        {
            PairedDataImage.Visibility = Visibility.Collapsed;
            PairedDataSummary.Visibility = Visibility.Collapsed;
            TimeSeriesImage.Visibility = Visibility.Visible;
            TimeSeriesSummary.Visibility = Visibility.Visible;
            TimeSeriesNextButton.Visibility = Visibility.Visible;
            PairedDataNextButton.Visibility = Visibility.Collapsed;
        }

        private void PairedDataOption_Selected(object sender, RoutedEventArgs e)
        {
            TimeSeriesImage.Visibility = Visibility.Collapsed;
            TimeSeriesSummary.Visibility = Visibility.Collapsed;
            PairedDataImage.Visibility = Visibility.Visible;
            PairedDataSummary.Visibility = Visibility.Visible;
            PairedDataNextButton.Visibility = Visibility.Visible;
            TimeSeriesNextButton.Visibility = Visibility.Collapsed;
        }

        public event RoutedEventHandler TimeSeriesNextClick;
        public event RoutedEventHandler PairedDataNextClick;

        private void TimeSeriesNextButton_Click(object sender, RoutedEventArgs e)
        {
            this.TimeSeriesNextClick?.Invoke(this, e);
        }

        private void PairedDataNextButton_Click(object sender, RoutedEventArgs e)
        {
            this.PairedDataNextClick?.Invoke(this, e);
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            TimeSeriesOption.IsSelected = true;
        }
    }
}
