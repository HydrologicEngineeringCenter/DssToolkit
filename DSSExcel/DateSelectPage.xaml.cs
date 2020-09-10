using SpreadsheetGear;
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
    /// Interaction logic for DateSelect.xaml
    /// </summary>
    public partial class DateSelectPage : UserControl
    {
        public IRange Dates;
        public DateSelectPage()
        {
            InitializeComponent();
        }

        public event RoutedEventHandler NextClick;
        public event RoutedEventHandler BackClick;

        private void DateSelectNextButton_Click(object sender, RoutedEventArgs e)
        {
            this.NextClick?.Invoke(this, e);
        }

        private void DateSelectBackButton_Click(object sender, RoutedEventArgs e)
        {
            this.BackClick?.Invoke(this, e);
        }

        private void ExcelView_RangeSelectionChanged(object sender, SpreadsheetGear.Windows.Controls.RangeSelectionChangedEventArgs e)
        {
            Dates = ExcelView.RangeSelection;
        }
    }
}
