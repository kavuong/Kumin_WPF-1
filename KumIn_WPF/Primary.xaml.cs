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
using System.Net;
using System.Net.Mail;

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

            List<CheckBox> checkBoxlist = new List<CheckBox>();
            // Find all elements
            FindChildGroup<CheckBox>(dgdListing, "checkboxinstance", ref checkBoxlist);

            foreach (CheckBox c in checkBoxlist)
            {
                if (c.IsChecked.Value)
                {
                    MessageBox.Show("works");
                }
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
        
        private void dgdListing_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
        }
        public static void FindChildGroup<T>(DependencyObject parent, string childName, ref List<T> list) where T : DependencyObject
        {
            // Checks should be made, but preferably one time before calling.
            // And here it is assumed that the programmer has taken into
            // account all of these conditions and checks are not needed.
            //if ((parent == null) || (childName == null) || (<Type T is not inheritable from FrameworkElement>))
            //{
            //    return;
            //}

            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);

            for (int i = 0; i < childrenCount; i++)
            {
                // Get the child
                var child = VisualTreeHelper.GetChild(parent, i);

                // Compare on conformity the type
                T child_Test = child as T;

                // Not compare - go next
                if (child_Test == null)
                {
                    // Go the deep
                    FindChildGroup<T>(child, childName, ref list);
                }
                else
                {
                    // If match, then check the name of the item
                    FrameworkElement child_Element = child_Test as FrameworkElement;

                    if (child_Element.Name == childName)
                    {
                        // Found
                        list.Add(child_Test);
                    }

                    // We are looking for further, perhaps there are
                    // children with the same name
                    FindChildGroup<T>(child, childName, ref list);
                }
            }

            return;
        }
    }
}
