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
using VedomostPropuskovPGEK.Data;

namespace VedomostPropuskovPGEK
{
    /// <summary>
    /// Логика взаимодействия для Gruppa.xaml
    /// </summary>
    public partial class Gruppa : Window
    {
        string cn;
        public Gruppa(string cn_C)
        {
            InitializeComponent();
            cbGruppa.ItemsSource = DataService.GetCurator(cn_C);
            cn = cn_C;
        }

        private void bGruppa_Click(object sender, RoutedEventArgs e)
        {

            MainWindow mainWindow = new MainWindow(cbGruppa.SelectedValue.ToString(), cn);
            mainWindow.Show();
            Close();
        }
    }
}
