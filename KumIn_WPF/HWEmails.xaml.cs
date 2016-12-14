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
using System.Net.Mail;
using System.Data;

namespace KumIn_WPF
{
    /// <summary>
    /// Interaction logic for HWEmails.xaml
    /// </summary>
    public partial class HWEmails : Window
    {
        SpreadsheetConnection emailConnection = new SpreadsheetConnection();
        DataTable emailTable = new DataTable();

        public const string ATTENDANCE_SHEET = "14j-XmVSs87CnsLX-TteOeIaAPak2G6_UTX6nU06kNWk";
        public const string ATTENDANCE_SHEET_RECORD = "Record";

        public HWEmails()
        {
            InitializeComponent();

            dpkDate.DisplayDate = DateTime.Now;

            emailTable.Columns.Add("FirstName");
            emailTable.Columns.Add("LastName");
            emailTable.Columns.Add("#Completed");
            emailTable.Columns.Add("#Missing");
            emailTable.Columns.Add("Email");

        }




        private void dpkDate_SelectedDateChanged(object sender, CalendarDateChangedEventArgs e)
        {
            int rowNum = emailConnection.getRowNum(ATTENDANCE_SHEET, ATTENDANCE_SHEET_RECORD + "!A1:A"
                , dpkDate.SelectedDate.Value.ToString("mm/dd/yyyy"));
            IList<IList<Object>> desiredDate = emailConnection.get(ATTENDANCE_SHEET, ATTENDANCE_SHEET_RECORD
                + "!A" + rowNum.ToString() + ":B" + rowNum.ToString());

            if (desiredDate[0].Count == 1)
            {
                populate(rowNum + 1);
            }
            else if (desiredDate[0].Count == 2)
            {
                MessageBox.Show("Emails for this date have already been processed. Select a new date.");
            }
            else if (desiredDate[0] == null)
            {
                MessageBox.Show("No students in record for selected date. Try again.");
            }
        }






        private void btnSendEmails_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < emailTable.Rows.Count; i++)
            {
                CheckBox outBox = dgdEmails.Columns[4].GetCellContent(dgdEmails.Items[i]) as CheckBox;

                if (outBox.IsChecked.Value)
                {
                    DataRowView drv = (DataRowView)dgdEmails.Items[i];
                    sendHomeworkEmail(drv);
                }

            }
        }





        private void sendHomeworkEmail(DataRowView drv)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

                string firstName = (drv["FirstName"]).ToString();
                string lastName = (drv["LastName"]).ToString();
                string email = (drv["Email"]).ToString();


                mail.From = new MailAddress("kumonsrn@gmail.com");
                mail.To.Add(email); // reg email
                mail.Subject = "Kumon HW notification";
                mail.Body = "Dear KUMON Parents,\n Your child, " + drv["FirstName"].ToString() + " " + drv["LastName"].ToString() +
                "attended center session today and turned in " + drv["#Completed"].ToString() +
                " assignment(s). We are still missing " + drv["#Missing"].ToString() +
                " of his/her assignment(s).\n Per " +
                "your request, this automated message is sent to notify you of the " +
                "missing homework. Although it's common that students will miss an assignment " +
                "from time to time due to various activities, we hope these notifications will " +
                "help you identify whether your child is chronically missing homework. \n\n Regards, \n KUMON San Ramon North \n 925-318-1628";
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("kumonsrn@gmail.com"
                    , "letmeout");
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
                MessageBox.Show("HW Email Sent", "Success");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        } 







        private void populate(int rowNum)
        {
            IList<IList<Object>> studentRecord = emailConnection.get(ATTENDANCE_SHEET
                , ATTENDANCE_SHEET_RECORD + "!A" + rowNum.ToString() + ":AAA" + rowNum.ToString());

            if (studentRecord != null)
            {
                if (studentRecord[0].Count > 2)
                {
                    string numMissing = studentRecord[0][10].ToString();

                    if (numMissing != "" && int.Parse(numMissing) > 0)
                    {
                        DataRow newStudent = emailTable.NewRow();

                        newStudent["FirstName"] = studentRecord[0][1].ToString();
                        newStudent["LastName"] = studentRecord[0][0].ToString();
                        newStudent["#Completed"] = studentRecord[0][9].ToString();
                        newStudent["#Missing"] = studentRecord[0][10].ToString();
                        newStudent["Email"] = studentRecord[0][3].ToString();

                        emailTable.Rows.Add(newStudent);
                    }

                    populate(rowNum + 1);
                }
            }
        }
    
    }
}
