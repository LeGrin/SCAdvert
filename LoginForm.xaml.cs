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

namespace SCAdvert
{
    /// <summary>
    /// Interaction logic for LoginForm.xaml
    /// </summary>
    public partial class LoginForm : Window
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private AdminForm AdmForm;

        private void BtnEnter_Click(object sender, RoutedEventArgs e)
        {

            if (TxtLogin.Text == "SCAdvert" & txtPassword.Password == "111111")
            {
                AdmForm = new AdminForm();
                AdmForm.Show();
                Close();
            }
            else
            {
                MessageBox.Show("Неверный логин или пароль!");
            }


            

        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
            
        }
    }
}
