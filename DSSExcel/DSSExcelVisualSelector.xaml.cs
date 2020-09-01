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
        public DSSExcelVisualEditor(IWorkbook workbook)
        {
            InitializeComponent();
            ExcelWorkbook.ActiveWorkbook = workbook;
        }
    }
}
