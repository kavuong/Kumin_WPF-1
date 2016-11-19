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
    /// Interaction logic for DeleteStudent.xaml
    /// </summary>
    public partial class DeleteStudent : Window
    {
        public DeleteStudent()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxButton deleteButton = MessageBoxButton.YesNo;

            if (MessageBox.Show("Are you sure you want to delete this student?",
                "Confirm Delete", deleteButton) == MessageBoxResult.Yes)
            {
                // Delete student from database
                MessageBox.Show("Student delete OK", "Student Deleted");
            }
            
        }
    }
}
