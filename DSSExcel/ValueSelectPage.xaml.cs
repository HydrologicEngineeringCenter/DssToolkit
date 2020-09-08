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
    /// Interaction logic for ValueSelectPage.xaml
    /// </summary>
    public partial class ValueSelectPage : UserControl
    {
        public ValueSelectPage()
        {
            InitializeComponent();
        }

        public event RoutedEventHandler NextClick;
        public event RoutedEventHandler BackClick;

        private void ValueSelectNextButton_Click(object sender, RoutedEventArgs e)
        {
            this.NextClick?.Invoke(this, e);
        }

        private void ValueSelectBackButton_Click(object sender, RoutedEventArgs e)
        {
            this.BackClick?.Invoke(this, e);
        }
    }
}
