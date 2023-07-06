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

            List<string> listMasonry = new List<string> { "0", "25", "50", "75", "100" };

            foreach (string masonry in listMasonry)
            {
                cmbCodeMasonry.Items.Add(masonry);
            }

            cmbCodeMasonry.SelectedIndex = 0;
        }

        public string GetComboBoxCurElevSelectedItem()
        {
            return cmbCurElev.SelectedItem.ToString();
        }

        public string GetComboBoxNewElevSelectedItem()
        {
            return cmbNewElev.SelectedItem.ToString();
        }

        public string GetComboBoxCodeMasonrySelectedItem()
        {
            string content = "";

            ComboBoxItem selectedItem = cmbCodeMasonry.SelectedItem as ComboBoxItem;

            if (selectedItem == null)
            {
                content = cmbCodeMasonry.Text;
            }
            else
            {
                content = selectedItem.Content.ToString();
            }

            return content;
        }

        public string ManualTextEnter = "";

        private void cmbCodeMasonry_TextInput(object sender, TextChangedEventArgs e)
        {
            ManualTextEnter = cmbCodeMasonry.Text;
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

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
