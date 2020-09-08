using Hec.Dss.Excel;
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

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for DSSExcelVisualEditor.xaml
    /// </summary>
    public partial class DSSExcelVisualEditor : Window
    {
        public IRange Dates;
        public IRange Ordinates;
        public IRange Values;

        public string ComponentSelection
        {
            get
            {
                return (RecordComponent.SelectedItem as ListBoxItem).Content.ToString();
            }
        }
        public DSSExcelVisualEditor(IWorkbook workbook)
        {
            InitializeComponent();
            ExcelWorkbook.ActiveWorkbook = workbook;
            Dates = ExcelWorkbook.RangeSelection;
            Ordinates = ExcelWorkbook.RangeSelection;
            Values = ExcelWorkbook.RangeSelection;
        }

        private void RecordComponent_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComponentSelection == "Date/Time")
            {
                ExcelWorkbook.RangeSelection = Dates == null ? ExcelWorkbook.RangeSelection : Dates;
                ExcelWorkbook.RangeSelectionBorderBrush = Brushes.Aqua;
            }
            else if (ComponentSelection == "Ordinates")
            {
                ExcelWorkbook.RangeSelection = Ordinates == null ? ExcelWorkbook.RangeSelection : Ordinates;
                ExcelWorkbook.RangeSelectionBorderBrush = Brushes.Green;
            }
            else if (ComponentSelection == "Values")
            {
                ExcelWorkbook.RangeSelection = Values == null ? ExcelWorkbook.RangeSelection : Values;
                ExcelWorkbook.RangeSelectionBorderBrush = Brushes.Red;
            }
        }

        private void ExcelWorkbook_RangeSelectionChanged(object sender, SpreadsheetGear.Windows.Controls.RangeSelectionChangedEventArgs e)
        {
            if (ComponentSelection == "Date/Time")
            {
                Dates = ExcelWorkbook.RangeSelection;
            }
            else if (ComponentSelection == "Ordinates")
            {
                Ordinates = ExcelWorkbook.RangeSelection;
            }
            else if (ComponentSelection == "Values")
            {
                Values = ExcelWorkbook.RangeSelection;
            }
        }

        private bool CheckRecordComponents()
        {

            if (!CheckDates(Dates))
                return false;

            if (!AreSelectionsCompatible())
                return false;

            if (!CheckOrdinates())
                return false;

            if (!CheckValues())
                return false;

            return true;
        }

        private bool CheckValues()
        {
            throw new NotImplementedException();
        }

        private bool CheckOrdinates()
        {
            throw new NotImplementedException();
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            if (!CheckRecordComponents())
                return;

            for (int i = 0; i < Dates.RowCount; i++)
            {
                
            }
            
            for (int i = 0; i < Ordinates.RowCount; i++)
            {

            }
            
            for (int i = 0; i < Values.RowCount; i++)
            {

            }
        }

        private DateTime GetDateTime(double value)
        {
            DateTime dt;
            var b = DateTime.TryParse(ExcelWorkbook.ActiveWorkbook.NumberToDateTime(value).ToString(), out dt);
            return b ? dt : new DateTime();
        }

        private bool CheckDates(IRange selection)
        {
            if (selection.ColumnCount != 1)
            {
                MessageBox.Show("The selection for Date/Time should only have one column of data.", "Date/Time Selection Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            ExcelWorkbook.ActiveWorkbookSet.GetLock();
            for (int i = 0; i < selection.RowCount; i++)
            {
                if (selection[i, 0].NumberFormatType != NumberFormatType.DateTime &&
                    selection[i, 0].NumberFormatType != NumberFormatType.Date)
                {
                    MessageBox.Show("All values selected for Date/Time don't follow the date and time format.", "Date/Time Selection Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            ExcelWorkbook.ActiveWorkbookSet.ReleaseLock();
            return true;
        }

        private bool AreSelectionsCompatible()
        {
            throw new NotImplementedException();
        }
    }
}
