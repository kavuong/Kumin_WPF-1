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
using System.Windows.Threading;
using System.Data;
using System.Globalization;

namespace KumIn_WPF
{
    /// <summary>
    /// Interaction logic for Primary.xaml
    /// </summary>
    public partial class Primary : Window
    {
        DateTime timeNow = DateTime.Now;
        DataTable dummyTable = new DataTable();

        public Primary()
        {
            InitializeComponent();
            lblTime.Content = timeNow.ToString("f");

            DispatcherTimer myTimer = new DispatcherTimer();
            myTimer.Interval = new TimeSpan(0, 0, 15);
            myTimer.Tick += new EventHandler(myTimer_Tick);
            myTimer.Start();
        }

        private void myTimer_Tick(object sender, object e)
        {
            timeNow = DateTime.Now;
            lblTime.Content = timeNow.ToString("f");
            foreach (DataRow row in dummyTable.Rows)
            {
                DateTime inTime = Convert.ToDateTime(row["InTime"]);
                TimeSpan t = TimeSpan.FromMinutes((timeNow - inTime).Minutes);
                int h = t.Hours;
                int mm = t.Minutes;
                row["Duration"] = t.ToString(@"h\:mm");                       
            }
            
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void btnAddNewStudent_Click(object sender, RoutedEventArgs e)
        {
            AddStudent myAddStudent = new AddStudent();
            myAddStudent.Show();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            DeleteStudent myDeleteStudent = new DeleteStudent();
            myDeleteStudent.Show();
        }

        private void btnSignIn_Click(object sender, RoutedEventArgs e)
        {
            DataRow dummyRow = dummyTable.NewRow();

            if (!dummyTable.Columns.Contains("FirstName"))
            {
                dummyTable.Columns.Add("FirstName");
            }
            if (!dummyTable.Columns.Contains("LastName"))
            {
                dummyTable.Columns.Add("LastName");
            }
            if (!dummyTable.Columns.Contains("InTime"))
            {
                dummyTable.Columns.Add("InTime");
            }
            if (!dummyTable.Columns.Contains("Duration"))
            {
                dummyTable.Columns.Add("Duration");
            }
            DataClasses1DataContext db = new DataClasses1DataContext();
            var user = (from u in db.FStudentTables
                        where u.Barcode == txtSignIn.Text
                        select u).FirstOrDefault();
            dummyRow["FirstName"] = user.FirstName;
            dummyRow["LastName"] = user.LastName;
            dummyRow["InTime"] = DateTime.Now.ToString("t");
            dummyRow["Duration"] = "00:00:00";


            dgdListing.ItemsSource = dummyTable.DefaultView;
            dummyTable.Rows.Add(dummyRow);

            txtSignIn.Clear();
        }

    }
}
