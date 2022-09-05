using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using VedomostPropuskovPGEK.Data;
using System.DirectoryServices.AccountManagement;

namespace VedomostPropuskovPGEK
{
    public partial class Authorization : Window
    {
        public Authorization()
        {
            InitializeComponent();
            LoginBox.Focus();
        }

        private string Login { get; set; }
        private string Password { get; set; }

        private void AuthorizationButton_Click(object sender, RoutedEventArgs e)
        {
            Login = "fik"; /*LoginBox.Text;*/
            Password = "291268"; /*PasswordBox.Password;*/
            try
            {
                //using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, "college.local", "DC=college,DC=local", Login, Password))
                //{
                //    if (DataService.GetCurator(Login).Count != 0 && pc.ValidateCredentials(Login, Password))
                //    {
                        int n = DataService.GetCurator(Login).Count;
                        if (n > 1)
                        {
                            Gruppa gruppa = new Gruppa(Login);
                            gruppa.Show();
                            Close();
                        }
                        else
                        {
                            string gr = DataService.GetCurator(Login)[0].cn_G;
                            MainWindow window = new MainWindow(gr, DataService.GetCurator(Login)[0].cn_C);
                            window.Show();
                            Close();
                        }
                //    }
                //    else
                //    {
                //        MessageBox.Show("Куратора с данным логином и паролем не существует!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
                //    }
                //    pc.Dispose();
                //}
            }
            catch 
            {
               MessageBox.Show("Ошибка при подключении к базе данных!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AuthorizationForms_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter) 
                    NextVoiti.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
