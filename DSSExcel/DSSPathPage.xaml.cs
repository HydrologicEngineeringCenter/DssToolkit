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
        public string Apart { get; set; }
        public string Bpart { get; set; }
        public string Cpart { get; set; }
        public string Dpart { get; set; }
        public string Epart { get; set; }
        public string Fpart { get; set; }

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
    }
}
