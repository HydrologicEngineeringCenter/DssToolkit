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
using System.Windows.Shapes;
using Hec.Dss.Excel;

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for SelectDataType.xaml
    /// </summary>
    public partial class DSSExcelGuidedImport : Window
    {
        ExcelReader r;
        public DSSExcelGuidedImport(string filename)
        {
            InitializeComponent();
            r = new ExcelReader(filename);
            DatePage.ExcelView.ActiveWorkbook = r.workbook;
            OrdinatePage.ExcelView.ActiveWorkbook = r.workbook;
            TimeSeriesValuePage.ExcelView.ActiveWorkbook = r.workbook;
            PairedDataValuePage.ExcelView.ActiveWorkbook = r.workbook;
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
            if (!CheckDates(DatePage.Dates))
                return;
            DatePage.Visibility = Visibility.Collapsed;
            TimeSeriesValuePage.Visibility = Visibility.Visible;
        }

        private bool CheckDates(IRange dates)
        {
            if (dates.ColumnCount != 1)
            {
                MessageBox.Show("The selection for Date/Time should only have one column of data.", "Date/Time Selection Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            DatePage.ExcelView.ActiveWorkbookSet.GetLock();
            for (int i = 0; i < dates.RowCount; i++)
            {
                if (dates[i, 0].NumberFormatType != NumberFormatType.DateTime &&
                    dates[i, 0].NumberFormatType != NumberFormatType.Date)
                {
                    MessageBox.Show("All values selected for Date/Time don't follow the date and time format.", "Date/Time Selection Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            DatePage.ExcelView.ActiveWorkbookSet.ReleaseLock();
            return true;
        }

        private void DatePage_BackClick(object sender, RoutedEventArgs e)
        {
            DatePage.Visibility = Visibility.Collapsed;
            DataTypePage.Visibility = Visibility.Visible;
        }

        private void OrdinatePage_NextClick(object sender, RoutedEventArgs e)
        {
            if (!CheckOrdinates(OrdinatePage.Ordinates))
                return;
            OrdinatePage.Visibility = Visibility.Collapsed;
            PairedDataValuePage.Visibility = Visibility.Visible;
        }

        private bool CheckOrdinates(IRange ordinates)
        {
            if (ordinates.ColumnCount != 1)
            {
                MessageBox.Show("The selection for ordinates should only have one column of data.", "Ordinate Selection Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            OrdinatePage.ExcelView.ActiveWorkbookSet.GetLock();
            for (int i = 0; i < ordinates.RowCount; i++)
            {
                if (ordinates[i, 0].NumberFormatType != NumberFormatType.Number)
                {
                    MessageBox.Show("All values selected for ordinates are not numbers.", "Ordinate Selection Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            OrdinatePage.ExcelView.ActiveWorkbookSet.ReleaseLock();
            return true;
        }

        private void OrdinatePage_BackClick(object sender, RoutedEventArgs e)
        {
            OrdinatePage.Visibility = Visibility.Collapsed;
            DataTypePage.Visibility = Visibility.Visible;
        }

        private void TimeSeriesValuePage_NextClick(object sender, RoutedEventArgs e)
        {
            if (!CheckTimeSeriesValues(TimeSeriesValuePage.Values))
                return;
            TimeSeriesValuePage.Visibility = Visibility.Collapsed;
            PathPage.Visibility = Visibility.Visible;
        }

        private bool CheckTimeSeriesValues(IRange values)
        {
            if (values.ColumnCount != 1)
            {
                MessageBox.Show("The selection for values should only have one column of data.", "Value Selection Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            TimeSeriesValuePage.ExcelView.ActiveWorkbookSet.GetLock();
            for (int i = 0; i < values.RowCount; i++)
            {
                if (values[i, 0].NumberFormatType != NumberFormatType.Number)
                {
                    MessageBox.Show("All selected values are not numbers.", "Value Selection Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            TimeSeriesValuePage.ExcelView.ActiveWorkbookSet.ReleaseLock();
            return true;
        }

        private void TimeSeriesValuePage_BackClick(object sender, RoutedEventArgs e)
        {
            TimeSeriesValuePage.Visibility = Visibility.Collapsed;
            DatePage.Visibility = Visibility.Visible;
        }

        private void PairedDataValuePage_NextClick(object sender, RoutedEventArgs e)
        {
            if (!CheckPairedDataValues(PairedDataValuePage.Values))
                return;
            PairedDataValuePage.Visibility = Visibility.Collapsed;
            PathPage.Visibility = Visibility.Visible;
        }

        private bool CheckPairedDataValues(IRange values)
        {
            PairedDataValuePage.ExcelView.ActiveWorkbookSet.GetLock();
            for (int i = 0; i < values.RowCount; i++)
            {
                for (int j = 0; j < values.ColumnCount; j++)
                {
                    if (values[i, j].NumberFormatType != NumberFormatType.Number)
                    {
                        MessageBox.Show("All selected values are not numbers.", "Value Selection Error",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        return false;
                    }
                }
            }
            PairedDataValuePage.ExcelView.ActiveWorkbookSet.ReleaseLock();
            return true;
        }

        private void PairedDataValuePage_BackClick(object sender, RoutedEventArgs e)
        {
            PairedDataValuePage.Visibility = Visibility.Collapsed;
            OrdinatePage.Visibility = Visibility.Visible;
        }

        private void PathPage_ImportClick(object sender, RoutedEventArgs e)
        {

        }

        private void PathPage_BackClick(object sender, RoutedEventArgs e)
        {

        }
    }
}
