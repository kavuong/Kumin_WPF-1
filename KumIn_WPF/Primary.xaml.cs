using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
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
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API KumIn";
        public Primary()
        {
            InitializeComponent();

            dummyTable.Columns.Add("FirstName");
            dummyTable.Columns.Add("LastName");
            dummyTable.Columns.Add("InTime");
            dummyTable.Columns.Add("Duration");
            dummyTable.Columns.Add("LastDay");

            lblTime.Content = timeNow.ToString("f");

            DispatcherTimer myTimer = new DispatcherTimer();
            myTimer.Interval = new TimeSpan(0, 0, 15);
            myTimer.Tick += new EventHandler(myTimer_Tick);
            myTimer.Start();

            string spreadsheetId = "14j-XmVSs87CnsLX-TteOeIaAPak2G6_UTX6nU06kNWk";
            string range = "Record!A1:A";
            IList<IList<Object>> checkValues = getSpreadsheetInfo(spreadsheetId, range);

            bool flag = false;
            foreach (var row in checkValues)
            {
                if (row[0].ToString() == DateTime.Now.ToString("MM/dd/yyyy"))
                {
                    flag = true;
                    break;
                }
            }

            if (!flag)
            {
                List<Object> date = new List<object>() { DateTime.Now.ToString("MM/dd/yyyy") };
                appendSpreadsheetInfo(date, spreadsheetId, range);
            }

        }

        private void myTimer_Tick(object sender, object e)
        {
            timeNow = DateTime.Now;
            lblTime.Content = timeNow.ToString("f");
            /*
            foreach (DataRow row in dummyTable.Rows)
            {
                DateTime inTime = Convert.ToDateTime(row["InTime"]);
                TimeSpan t = TimeSpan.FromMinutes((timeNow - inTime).Minutes);
                int h = t.Hours;
                int mm = t.Minutes;
                row["Duration"] = t.ToString(@"h\:mm");
                                     
            }
            */

            IList<IList<Object>> values = getSpreadsheetInfo("14j-XmVSs87CnsLX-TteOeIaAPak2G6_UTX6nU06kNWk", "Sheet1!A1:Z");

            if (values != null && values.Count > 0)
            {
                dummyTable.Clear();
                foreach (var row in values)
                {                    
                    if (row.Count != 1)
                    {
                        if (row != values[0] && row.Count != 0)
                        {
                            DataRow dummyRow = dummyTable.NewRow();
                            char[] delimiterChars = { ':', ' ' };
                            dummyRow["FirstName"] = row[1];
                            dummyRow["LastName"] = row[0];
                            /*
                            string[] inTime = (row[9].ToString()).Split(delimiterChars);
                            MessageBox.Show(Convert.ToString(inTime.Length));
                            dummyRow["InTime"] = Convert.ToString(inTime[0]) + Convert.ToString(inTime[1]);
                            DateTime duration = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, Convert.ToInt32(inTime[0]), Convert.ToInt32(inTime[1]), 0, 0);
                            TimeSpan result = DateTime.Now - duration;   
                            */
                            dummyRow["InTime"] = row[9];

                            TimeSpan duration = DateTime.Parse(Convert.ToString(DateTime.Now)).Subtract(DateTime.Parse(Convert.ToString(row[9])));
                            dummyRow["Duration"] = duration.ToString(@"hh\:mm");
                            dummyRow["LastDay"] = row[8];

                            dgdListing.ItemsSource = dummyTable.DefaultView;
                            dummyTable.DefaultView.Sort = "Duration DESC";
                            dummyTable.Rows.Add(dummyRow);
                        }
                    }
                    

                }
            }

        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            // If scanner is not working ==> dialog confirm
            if (!char.IsLetter(txtUpdate.Text[0]))
            {
                Confirmation myConfirm = new Confirmation(txtUpdate.Text);
                if (myConfirm.ShowDialog() == true)
                {
                    // Populate display after checking firstname, lastname, number
                    if (!isSignedIn(new string[3] { myConfirm.Number, myConfirm.FirstName, myConfirm.LastName }))
                    {
                        populateDataGrid(new string[3] { myConfirm.Number, myConfirm.FirstName, myConfirm.LastName });
                    }
                    else // confirm signout then do it
                    {
                        int rowNum = 0;
                        foreach (DataRow row in dummyTable.Rows)
                        {
                            if (txtUpdate.Text.Substring(1) == row["Barcode"].ToString())
                                break;
                            else
                                rowNum++;
                        }
                        DateTime timeIn = Convert.ToDateTime(dummyTable.Rows[rowNum]["TimeIn"].ToString());
                        if (DateTime.Now - timeIn > new TimeSpan(0, 1, 0))
                        {
                            signOut(new string[3] { myConfirm.Number, myConfirm.FirstName, myConfirm.LastName });
                        }
                        else if (MessageBox.Show("Student was recently signed in, sign out already?"
                            , "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            signOut(new string[3] { myConfirm.Number, myConfirm.FirstName, myConfirm.LastName });
                        }
                    }
                    // Throw error if not found.
                }
            }
            // Scanner is working so barcode = 'a' + number
            else
            {
                // Populate display after checking barcode
                if (!isSignedIn(new string[1] { txtUpdate.Text.Substring(1) }))
                {
                    populateDataGrid(new string[1] { txtUpdate.Text.Substring(1) });
                }
                else // confirm signout then do it
                {
                    int rowNum = 0;
                    foreach (DataRow row in dummyTable.Rows)
                    {
                        if (txtUpdate.Text.Substring(1) == row["Barcode"].ToString())
                            break;
                        else
                            rowNum++;
                    }
                    DateTime timeIn = Convert.ToDateTime(dummyTable.Rows[rowNum]["TimeIn"].ToString());
                    if (DateTime.Now - timeIn > new TimeSpan(0,1,0))
                    {
                        signOut(new string[1] { txtUpdate.Text.Substring(1) });
                    }
                    else if (MessageBox.Show("Student was recently signed in, sign out already?"
                        , "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        signOut(new string[1] { txtUpdate.Text.Substring(1) });
                    }
                }
                
            }

            
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

        private bool isSignedIn(string[] checkValues)
        {
            // Scanner works? checkvalues.count == 1

            foreach (DataRow row in dummyTable.Rows)
            {
                if (row["Barcode"].ToString() == checkValues[0])
                {
                    // Check if manual entry
                    if (checkValues.Count() == 3)
                        return row["FirstName"].ToString().ToUpper() == checkValues[1].ToUpper()
                            && row["LastName"].ToString().ToUpper() == checkValues[2].ToUpper();
                    else
                        return true;
                }

                // We need to store the barcode inside dataTable but not display in datagrid.
                // AKA: new column in datatable, but dont bind that column to datagrid.
                // Use dateTime.Now - signInTime < 1 ==> messagebox confirm 
                // Also, remove the sign out textbox and button.


                // Good idea to also add column with number of subjects in our dataTable too
                // Storing this info which we use a lot can help shorten number of spreadsheet calls
                // All of this would also help when we record information into the record as we can
                // use the info in our internal datatable row rather than  
            }

            return false;
        }

        private void signOut(string[] checkValues)
        {
            // Scanner works? checkvalues.count == 1.

            string tempSheet = "14j-XmVSs87CnsLX-TteOeIaAPak2G6_UTX6nU06kNWk";
            string range = "Sheet1!A1:J";

            IList<IList<Object>> getStudents = getSpreadsheetInfo(tempSheet, range);
            int rowNum = 1;
            foreach (var row in getStudents)
            {
                if (row.Count != 0)
                {
                    if (checkValues[0] == row[2].ToString())
                    {
                        if (checkValues.Count() == 3)
                        {
                            if (row[1].ToString().ToUpper() == checkValues[1].ToUpper()
                            && row[0].ToString().ToUpper() == checkValues[2].ToUpper())
                                break; // additional autherntiacation
                        }
                        else
                            break; // Scan by scanner
                    }
                    else
                    {
                        rowNum++;
                    }
                }
                else
                    rowNum++;
            }

            List<Object> pasteRange = new List<object>() { };
            for (int i = 0; i < 10; i++)
                pasteRange.Add(getStudents[rowNum - 1][i]);
            TimeSpan duration = DateTime.Now.Subtract(Convert.ToDateTime(getStudents[rowNum - 1][9]));

            pasteRange.Add((duration).ToString(@"hh\:mm"));

            List<Object> deleteRow = new List<object>();

            for (int i = 0; i < 10; i++)
                deleteRow.Add("");

            range = "Sheet1!A" + rowNum.ToString() + ":J" + rowNum.ToString();

            updateSpreadsheetInfo(deleteRow, tempSheet, range);

            // Transfer row to permanent record

            range = "Record!A1:A";

            appendSpreadsheetInfo(pasteRange, tempSheet, range);

            txtUpdate.Text = "";
            txtUpdate.Focus();
        }

        private void populateDataGrid(string[] checkValues)
        {
            // Scanner works? checkvalues.count == 1.


            DataRow dummyRow = dummyTable.NewRow();
            string firstName = "";
            string lastName = "";
            string lastDayIn = "";
            string range = "";
            string spreadsheetId = "1Lxn9qUxUbNWt3cI70CuTEIxCfgpxjAlZPd6ARph4oCM";  //STUDENT-database
            string spreadsheetId2 = "14j-XmVSs87CnsLX-TteOeIaAPak2G6_UTX6nU06kNWk"; //attendance record
            string barcode = "";


            IList<IList<Object>> values = getSpreadsheetInfo(spreadsheetId, "Sheet1!A1:I");
            int rowNum = 1;

            if (values != null && values.Count > 0)
            {


                foreach (var row in values)
                {
                    if (row[4].ToString() == checkValues[0])
                    {
                        if (checkValues.Count() == 3)
                        {
                            if (row[1].ToString().ToUpper() == checkValues[1].ToUpper()
                            && row[0].ToString().ToUpper() == checkValues[2].ToUpper())
                            {
                                range = "Sheet1!A" + rowNum.ToString() + ":" + "AAA" + rowNum.ToString();
                                break; // additional autherntiacation
                            }
                        }
                        else
                        {
                            range = "Sheet1!A" + rowNum.ToString() + ":" + "AAA" + rowNum.ToString();
                            break;
                        }
                    }
                    else
                        rowNum++;
                }
            }

            values = getSpreadsheetInfo(spreadsheetId, range);


            foreach (var row in values)
            {
                lastName = row[7].ToString();
                firstName = row[5].ToString();
                lastDayIn = row[8].ToString(); //check
                barcode = row[4].ToString(); // check             
            }

            dummyRow["FirstName"] = firstName;
            dummyRow["LastName"] = lastName;
            dummyRow["InTime"] = DateTime.Now.ToString("t");
            dummyRow["Duration"] = "00:00:00";
            dummyRow["LastDay"] = lastDayIn;
            dummyRow["Barcode"] = barcode;


            string range2 = "Sheet1!A1:Z1000";
            values = getSpreadsheetInfo(spreadsheetId2, range2);

            int pasteRowNum = 1;
            foreach (var row in values)
            {
                if (row.Count == 0)
                {
                    break;
                }
                pasteRowNum++;
            }

            // edit this to pull appropriate valued from database into temp and record sheet columns
            if (lastName != null && firstName != null)
            {
                range = "Sheet1!A" + rowNum.ToString() + ":Z" + rowNum.ToString();
                range2 = "Sheet1!A" + pasteRowNum.ToString() + ":Z" + pasteRowNum.ToString();

                IList<IList<Object>> getValues = getSpreadsheetInfo(spreadsheetId, range);

                var oblist = new List<Object>();

                foreach (var row in getValues)
                {
                    for (int i = 0; i < row.Count; i++)
                        oblist.Add(row[i]);
                }

                updateSpreadsheetInfo(oblist, spreadsheetId2, range2);


                // inputs timeIn into temp spreadsheet
                String range3 = "Sheet1!J" + pasteRowNum.ToString();
                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "COLUMNS";

                var oblist2 = new List<object>() { DateTime.Now.ToString("t") };
                valueRange.Values = new List<IList<object>> { oblist2 };

                updateSpreadsheetInfo(oblist2, spreadsheetId2, range3);
            }
            dgdListing.ItemsSource = dummyTable.DefaultView;
            dummyTable.Rows.Add(dummyRow);
            dummyTable.DefaultView.Sort = "Duration DESC";
            txtUpdate.Clear();
            txtUpdate.Focus();

            List<Object> today = new List<object>() { DateTime.Now.ToString("MM/dd") };
            updateSpreadsheetInfo(today, spreadsheetId, "I" + rowNum);
        }


        private void sendHomeworkEmail()
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                //DataClasses1DataContext db = new DataClasses1DataContext();

                DataRowView drv = (DataRowView) dgdListing.SelectedItem;
                string firstName = (drv["FirstName"]).ToString();
                string lastName = (drv["LastName"]).ToString();
                string email = "";
                string range = "";
                string spreadsheetId = "1KmeqPF07jnjZJ_FwWgZ4taj9N3c-3HK7qvOJmG1TpWs";
                /*
                var user = (from u in db.FStudentTables
                            where u.FirstName == firstName && u.LastName == lastName
                            select u).FirstOrDefault();
                            */

                IList<IList<Object>> values = getSpreadsheetInfo(spreadsheetId, "Sheet1!A1:J");
                int rowNum = 1;
                if (values != null && values.Count > 0)
                {
                    

                    foreach (var row in values)
                    {
                        if (drv["FirstName"].ToString() == row[1].ToString()
                        && drv["LastName"].ToString() == row[0].ToString())
                        {
                            break;
                        }
                        else
                            rowNum++;
                    }
                }


                var myRow = values[rowNum - 1];

                email = myRow[3].ToString();

                mail.From = new MailAddress("anthonyluukumon@gmail.com");
                mail.To.Add(email); // reg email
                mail.Subject = "Kumon HW notification";
                mail.Body = "Your student has not completed all of the homework assigned since last Kumon session.";

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("anthonyluukumon@gmail.com"
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

        private void sendOutEmail()
        {
            try
            {
                MailMessage mail1 = new MailMessage();
                MailMessage mail2 = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                string phone1 = "";
                string phone2 = "";
                string carrier1 = "";
                string carrier2 = "";
                string carrierString1 = "";
                string carrierString2 = "";
                /*
                DataClasses1DataContext db = new DataClasses1DataContext(); */
                DataRowView drv = (DataRowView)dgdListing.SelectedItem;
                string firstName = (drv["FirstName"]).ToString();
                string lastName = (drv["LastName"]).ToString();
                /*
                var user = (from u in db.FStudentTables
                            where u.FirstName == firstName && u.LastName == lastName
                            select u).FirstOrDefault();                                                          
                */
                
                IList<IList<Object>>values = getSpreadsheetInfo("1KmeqPF07jnjZJ_FwWgZ4taj9N3c-3HK7qvOJmG1TpWs", "Sheet1!A1:J");

                int rowNum = 1;
                foreach (var row in values)
                {
                    if (drv["FirstName"].ToString() == row[1].ToString()
                        && drv["LastName"].ToString() == row[0].ToString())
                        break;
                    else
                        rowNum++;
                }

                var myRow = values[rowNum - 1];

                phone1 = myRow[4].ToString();
                carrier1 = myRow[5].ToString();
                phone2 = myRow[6].ToString();
                carrier2 = myRow[7].ToString();

                if (carrier1 == "Verizon")
                {
                    carrierString1 = "@vtext.com";
                }
                else if (carrier1 == "AT&T")
                {
                    carrierString1 = "@txt.att.net";
                }
                else if (carrier1 == "Sprint")
                {
                    carrierString1 = "@messaging.sprintpcs.com";
                }
                else if (carrier1 == "T-Mobile")
                {
                    carrierString1 = "@tmomail.net";
                }

                if (carrier2 == "Verizon")
                {
                    carrierString2 = "@vtext.com";
                }
                else if (carrier2 == "AT&T")
                {
                    carrierString2 = "@txt.att.net";
                }
                else if (carrier2 == "Sprint")
                {
                    carrierString2 = "@messaging.sprintpcs.com";
                }
                else if (carrier2 == "T-Mobile")
                {
                    carrierString2 = "@tmomail.net";
                }
                if (phone1 != null && carrierString1 != null)
                {
                    mail1.From = new MailAddress("anthonyluukumon@gmail.com");
                    mail1.To.Add(phone1 + carrierString1); //phone
                    mail1.Subject = "Kumon Reminder";
                    mail1.Body = "Your child - " + firstName + " " + lastName + " - is ready to be picked up";
                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("anthonyluukumon@gmail.com"
                           , "letmeout");
                    SmtpServer.EnableSsl = true;
                    SmtpServer.Send(mail1);
                }
                /*

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
                */
                if (phone2 != "" && carrierString2 != "")
                {
                    mail2.From = new MailAddress("anthonyluukumon@gmail.com");
                    mail2.To.Add(phone2 + carrierString2); //phone
                    mail2.Subject = "Kumon Reminder";
                    mail2.Body = "Your student is done.";
                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("anthonyluukumon@gmail.com"
                           , "letmeout");
                    SmtpServer.EnableSsl = true;
                    SmtpServer.Send(mail2);
                }
                

                MessageBox.Show("Text(s) sent","Success");
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

                

        private void btnAddNewStudent_Click(object sender, RoutedEventArgs e)
        {
            AssignWork myAssignWork = new AssignWork();
            myAssignWork.Show();
        }


        private IList<IList<Object>> getSpreadsheetInfo (string spreadsheetId, string range)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = System.IO.Path.Combine(credPath, ".credentials/sheets.googleapis.com-kumin-assignment.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            SpreadsheetsResource.ValuesResource.GetRequest request =
                        service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            return values;
        }
        private void updateSpreadsheetInfo(List<Object> oblist, string spreadsheetId, string range)
        {
            List<IList<Object>> values = new List<IList<object>> { oblist };

            ValueRange valueRange = new ValueRange();
            valueRange.Values = values;
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = System.IO.Path.Combine(credPath, ".credentials/sheets.googleapis.com-kumin-assignment.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            SpreadsheetsResource.ValuesResource.UpdateRequest request =
                        service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            request.Execute();
        }

        private void appendSpreadsheetInfo(List<Object> oblist, string spreadsheetId, string range)
        {
            List<IList<Object>> values = new List<IList<object>> { oblist };

            ValueRange valueRange = new ValueRange();
            valueRange.Values = values;
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = System.IO.Path.Combine(credPath, ".credentials/sheets.googleapis.com-kumin-assignment.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            SpreadsheetsResource.ValuesResource.AppendRequest request =
                        service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            request.Execute();
        }

        private void txtUpdate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnUpdate_Click((object)sender, (RoutedEventArgs)e);
            }
        }
    }
}
