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

        private void onHomeworkChecked(object sender, RoutedEventArgs e)
        {
            sendHomeworkEmail();
        }

        private void onOutChecked(object sender, RoutedEventArgs e)
        {
            sendOutEmail();
        }

        private void sendHomeworkEmail()
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                DataClasses1DataContext db = new DataClasses1DataContext();

                DataRowView drv = (DataRowView) dgdListing.SelectedItem;
                String firstName = (drv["FirstName"]).ToString();
                String lastName = (drv["LastName"]).ToString();
                var user = (from u in db.FStudentTables
                            where u.FirstName == firstName && u.LastName == lastName
                            select u).FirstOrDefault();

                mail.From = new MailAddress("anthonyluukumon@gmail.com");
                mail.To.Add(user.RealEmail); // reg email
                mail.Subject = "Kumon HW notification";
                mail.Body = "Your student has not completed all of the homework assigned since last Kumon session.";

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("anthonyluukumon@gmail.com"
                    , "letmeout");
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
                MessageBox.Show("HW Email Sent");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void sendOutEmail()
        {
            try
            {
                MailMessage mail1 = new MailMessage();
                MailMessage mail2 = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                DataClasses1DataContext db = new DataClasses1DataContext();
                String carrierString1 = "";
                String carrierString2 = "";
                DataRowView drv = (DataRowView)dgdListing.SelectedItem;
                String firstName = (drv["FirstName"]).ToString();
                String lastName = (drv["LastName"]).ToString();
                var user = (from u in db.FStudentTables
                            where u.FirstName == firstName && u.LastName == lastName
                            select u).FirstOrDefault();

                if (user.Carrier1 == "Verizon")
                {
                    carrierString1 = "@vtext.com";
                }
                else if (user.Carrier1 == "AT&T")
                {
                    carrierString1 = "@txt.att.net";
                }
                else if (user.Carrier1 == "Sprint")
                {
                    carrierString1 = "@messaging.sprintpcs.com";
                }
                else if (user.Carrier1 == "T-Mobile")
                {
                    carrierString1 = "@tmomail.net";
                }                
                mail1.From = new MailAddress("anthonyluukumon@gmail.com");
                mail1.To.Add(user.Phone1 + carrierString1); //phone
                mail1.Subject = "Kumon Reminder";
                mail1.Body = "Your student is done.";
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("anthonyluukumon@gmail.com"
                       , "letmeout");
                SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail1);

                if (user.Carrier2 == "Verizon")
                {
                    carrierString2 = "@vtext.com";
                }
                else if (user.Carrier2 == "AT&T")
                {
                    carrierString2 = "@txt.att.net";
                }
                else if (user.Carrier2 == "Sprint")
                {
                    carrierString2 = "@messaging.sprintpcs.com";
                }
                else if (user.Carrier2 == "T-Mobile")
                {
                    carrierString2 = "@tmomail.net";
                }
                mail2.From = new MailAddress("anthonyluukumon@gmail.com");
                mail2.To.Add(user.Phone2 + carrierString2); //phone
                mail2.Subject = "Kumon Reminder";
                mail2.Body = "Your student is done.";
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("anthonyluukumon@gmail.com"
                       , "letmeout");
                SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail2);

                MessageBox.Show("Text(s) sent");
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        
    }
}
