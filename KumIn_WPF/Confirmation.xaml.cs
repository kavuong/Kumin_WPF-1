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

namespace KumIn_WPF
{
    /// <summary>
    /// Interaction logic for Confirmation.xaml
    /// </summary>
    public partial class Confirmation : Window
    {
        public Confirmation(string number)
        {
            InitializeComponent();
            Number = number;
        }

        private void btnSignIn_Click(object sender, RoutedEventArgs e)
        {

            this.DialogResult = true;
            this.Close();
            

        }

        public string FirstName
        {
            get { return txtFirstName.Text; }
        }

        public string LastName
        {
            get { return txtLastName.Text; }
        }

        public string Number
        {
            get { return txtNumber.Text; }
            set { txtNumber.Text = value.ToString(); }
        }
    }
}
