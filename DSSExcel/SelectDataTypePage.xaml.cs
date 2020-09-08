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
            DataTypeSelectNextButton.IsEnabled = true;
        }

        private void PairedDataOption_Selected(object sender, RoutedEventArgs e)
        {
            TimeSeriesImage.Visibility = Visibility.Collapsed;
            TimeSeriesSummary.Visibility = Visibility.Collapsed;
            PairedDataImage.Visibility = Visibility.Visible;
            PairedDataSummary.Visibility = Visibility.Visible;
            DataTypeSelectNextButton.IsEnabled = true;
        }

        public event RoutedEventHandler Click;
        void DataTypeSelectButton_Click(object sender, RoutedEventArgs e)
        {
            this.Click?.Invoke(this, e);
        }
    }
}
