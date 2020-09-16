using Hec.Dss;
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
    /// Interaction logic for DSSPathPage.xaml
    /// </summary>
    public partial class DSSPathPage : UserControl
    {
        public UserControl PreviousPage;
        public RecordType currentRecordType;
        private bool tsPathGenerated = false;
        private bool pdPathGenerated = false;
        private DssPath tsPath = new DssPath();
        private DssPath pdPath = new DssPath();

        public string Apart 
        {
            get
            {
                return ApartTextBox.Text;
            }
        }
        public string Bpart 
        {
            get
            {
                return BpartTextBox.Text;
            }
        }
        public string Cpart 
        {
            get
            {
                return CpartTextBox.Text;
            }
        }
        public string Dpart 
        {
            get
            {
                return DpartTextBox.Text;
            }
        }
        public string Epart 
        {
            get
            {
                return EpartTextBox.Text;
            }
        }
        public string Fpart 
        {
            get
            {
                return FpartTextBox.Text;
            }
        }


        public string GetPath 
        { 
            get
            {
                return "/" + Apart +
                    "/" + Bpart +
                    "/" + Cpart +
                    "/" + Dpart +
                    "/" + Epart +
                    "/" + Fpart + "/";
            }
        }
        public DSSPathPage()
        {
            InitializeComponent();
        }

        public event RoutedEventHandler ImportClick;
        public event RoutedEventHandler BackClick;
        private void DSSPathImportButton_Click(object sender, RoutedEventArgs e)
        {
            this.ImportClick?.Invoke(this, e);
        }

        private void DSSPathBackButton_Click(object sender, RoutedEventArgs e)
        {
            this.BackClick?.Invoke(this, e);
        }

        private void ShowTimeSeriesPath()
        {
            if (!tsPathGenerated)
                GenerateTimeSeriesPath();
            DataContext = tsPath;
        }

        private void GenerateTimeSeriesPath()
        {
            tsPath.Apart = "a" + Hec.Dss.Excel.ExcelTools.RandomString(3);
            tsPath.Bpart = "b" + Hec.Dss.Excel.ExcelTools.RandomString(3);
            tsPath.Cpart = "c" + Hec.Dss.Excel.ExcelTools.RandomString(3);
            tsPath.Dpart = "";
            tsPath.Epart = "";
            tsPath.Fpart = "TimeSeries" + Hec.Dss.Excel.ExcelTools.RandomString(3);
            tsPathGenerated = true;
        }

        private void ShowPairedDataPath()
        {
            if (!pdPathGenerated)
                GeneratePairedDataPath();
            DataContext = pdPath;
        }

        private void GeneratePairedDataPath()
        {
            pdPath.Apart = "a" + Hec.Dss.Excel.ExcelTools.RandomString(3);
            pdPath.Bpart = "b" + Hec.Dss.Excel.ExcelTools.RandomString(3);
            pdPath.Cpart = "c" + Hec.Dss.Excel.ExcelTools.RandomString(3);
            pdPath.Dpart = "";
            pdPath.Epart = "e" + Hec.Dss.Excel.ExcelTools.RandomString(3);
            pdPath.Fpart = "PairedData" + Hec.Dss.Excel.ExcelTools.RandomString(3);
            pdPathGenerated = true;
        }

        public void ShowPath(object record)
        {
            if (record is RecordType.IrregularTimeSeries || record is RecordType.RegularTimeSeries)
            {
                currentRecordType = RecordType.RegularTimeSeries;
                ShowTimeSeriesPath();
            }
            else if (record is RecordType.PairedData)
            {
                currentRecordType = RecordType.PairedData;
                ShowPairedDataPath();
            }
        }
    }
}
