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


namespace ElevationDesignation
{
    /// <summary>
    /// Interaction logic for Window.xaml
    /// </summary>
    public partial class frmReplaceElevation : Window
    {
        public frmReplaceElevation()
        {
            InitializeComponent();

            List<string> listElevations = new List<string> { "A", "B", "C", "D", "S", "T" };

            foreach (string elevation in listElevations)
            {
                cmbCurElev.Items.Add(elevation);
                cmbNewElev.Items.Add(elevation);
            }

            cmbCurElev.SelectedIndex = 0;
            cmbNewElev.SelectedIndex = 0;
        }

        public string GetComboBoxCurElevSelectedItem()
        {
            return cmbCurElev.SelectedItem.ToString();
        }

        public string GetComboBoxNewElevSelectedItem()
        {
            return cmbNewElev.SelectedItem.ToString();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
