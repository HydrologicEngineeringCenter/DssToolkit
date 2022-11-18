using SpreadsheetGear;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DSSExcel
{
  /// <summary>
  /// Interaction logic for DateSelect.xaml
  /// </summary>
  public partial class DateSelectPage : UserControl
  {
    public IRange Dates
    {
      get
      {
        return ExcelView.RangeSelection;
      }
    }
    public DateSelectPage()
    {
      InitializeComponent();
    }

    public event RoutedEventHandler NextClick;
    public event RoutedEventHandler BackClick;
    public event EventHandler TabSelectionChanged;

    private void DateSelectNextButton_Click(object sender, RoutedEventArgs e)
    {
      if (Dates.ColumnCount != 1)
      {
        MessageBox.Show("The selection for Date/Time should only have one column of data.", "Date/Time Selection Error",
                MessageBoxButton.OK, MessageBoxImage.Error);
        return;
      }

      this.NextClick?.Invoke(this, e);
    }

    private void DateSelectBackButton_Click(object sender, RoutedEventArgs e)
    {
      this.BackClick?.Invoke(this, e);
    }

    private void ExcelView_ActiveTabChanged(object sender, SpreadsheetGear.Windows.Controls.ActiveTabChangedEventArgs e)
    {
      this.TabSelectionChanged?.Invoke(this, e);
    }
  }
}
