/* 
When the window is constructed, all the columns are added to the DataTable. 
The myTimerTick function is configured to run every 15 seconds. Checks if
the current date is in the Attendance Record spreadsheet already, if not, 
the current date is appended to serve as delimiter for record of signins.
*/
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

// TRY CATCH STATEMENTS
// IF CONDITIONAL YES (notifcations)
// Change font # completed column
namespace KumIn_WPF
{
    /// <summary>
    /// Interaction logic for Primary.xaml
    /// </summary>
    public partial class Primary : Window
    {
        DateTime timeNow = DateTime.Now;
        DataTable dummyTable = new DataTable();
        SpreadsheetConnection kuminConnection = new SpreadsheetConnection();

        // Constants
        public const int TIMER_CYCLE = 15;                                      // seconds
        public const string ATTENDANCE_SHEET = "14j-XmVSs87CnsLX-TteOeIaAPak2G6_UTX6nU06kNWk";

        public const string ATTENDANCE_SHEET_PERM_RECORD = "Record";
        public const string ATTENDANCE_SHEET_TEMP_RECORD = "Sheet1";


        // Column Indexes
        public const int DATAGRID_FIRSTNAME = 0;
        public const int DATAGRID_LASTNAME = 1;
        public const int DATAGRID_COMPLETED = 5;
        public const int DATAGRID_MISSING = 6;

        public const int TEMPSHEET_FIRSTNAME = 1;
        public const int TEMPSHEET_LASTNAME = 0;
        public const int TEMPSHEET_BARCODE = 2;
        public const int TEMPSHEET_INTIME = 11;
        public const int TEMPSHEET_SUBJECTS = 12;
        public const int TEMPSHEET_LASTDAY = 8;
        public const int TEMPSHEET_COMPLETED = 9;
        public const int TEMPSHEET_MISSING = 10;

        public Primary()
        {
            InitializeComponent();

            dummyTable.Columns.Add("FirstName");
            dummyTable.Columns.Add("LastName");
            dummyTable.Columns.Add("InTime");
            dummyTable.Columns.Add("Duration");
            dummyTable.Columns.Add("LastDay");
            dummyTable.Columns.Add("#Completed");
            dummyTable.Columns.Add("#Missing");
            dummyTable.Columns.Add("Barcode");
            dummyTable.Columns.Add("#Subjects");

            TimeLabel = timeNow.ToString("f");

            
            DispatcherTimer myTimer = new DispatcherTimer();
            myTimer.Interval = new TimeSpan(0, 0, TIMER_CYCLE);
            myTimer.Tick += new EventHandler(myTimer_Tick);
            myTimer.Start();

            string centerDates = ATTENDANCE_SHEET_PERM_RECORD + "!A1:A";

            if (!kuminConnection.isValuePresent(ATTENDANCE_SHEET, centerDates, DateTime.Now.ToString("MM/dd/yyyy"))) ;
            {
                List<Object> date = new List<object>() { DateTime.Now.ToString("MM/dd/yyyy") };
                kuminConnection.append(date, ATTENDANCE_SHEET, centerDates);
            }
        }








        //*****************************************************************************************************
        // Definition of myTimer_Tick()
        // Transfers the DataGrid data to a new DataTable object. Checks if the data table is not null and 
        // if the completed and missing column values are filled out for a given row in the data table. 
        // Then, it gets the values in the completed and missing homework columns of that DataTable row and 
        // updates the temporary attendance sheet at the temporary spreadsheet row number matching the record 
        // of the student in the temporary sheet.
        // Populates the columns of the DataTable from the temporary sheet values. 
        private void myTimer_Tick(object sender, object e)
        {
            timeNow = DateTime.Now;
            TimeLabel = timeNow.ToString("f");
            string tempSheet = ATTENDANCE_SHEET_TEMP_RECORD + "!A1:Z";
            IList<IList<Object>> tempSheetValues = kuminConnection.get(ATTENDANCE_SHEET, tempSheet);



            if (dummyTable.Rows.Count != 0)
            {
                DataTable dt = new DataTable();
                dt = ((DataView)dgdListing.ItemsSource).ToTable();


                foreach (DataRow row in dt.Rows)
                {
                    if (row[DATAGRID_COMPLETED].ToString() != "" && row[DATAGRID_MISSING].ToString() != "")
                    {
                        int rowNum;
                        string firstNames = ATTENDANCE_SHEET_TEMP_RECORD + "!B1:B";
                        string lastNames = ATTENDANCE_SHEET_TEMP_RECORD + "!A1:A";
                        string studentRow;
                        List<Object> studentAssignmentEval = new List<object>
                            { row[DATAGRID_COMPLETED].ToString(), row[DATAGRID_MISSING].ToString() };


                        rowNum = kuminConnection.getRowNum(ATTENDANCE_SHEET, lastNames
                            , row[DATAGRID_LASTNAME].ToString(), firstNames, row[DATAGRID_FIRSTNAME].ToString());

                        studentRow = ATTENDANCE_SHEET_TEMP_RECORD + "!J" + rowNum.ToString() 
                            + ":K" + rowNum.ToString();

                        kuminConnection.update(studentAssignmentEval, ATTENDANCE_SHEET, studentRow);
                    }
                }
            }


            

            if (tempSheetValues != null && tempSheetValues.Count > 0)
            {
                dummyTable.Clear();
                for (int i = 1; i < tempSheetValues.Count; i++)
                {
                    IList<Object> studentIn = tempSheetValues[i];
                    DataRow dummyRow = dummyTable.NewRow();
                    TimeSpan duration = DateTime.Parse(Convert.ToString(DateTime.Now))
                        .Subtract(DateTime.Parse(Convert.ToString(studentIn[TEMPSHEET_INTIME])));

                    dummyRow["FirstName"] = studentIn[TEMPSHEET_FIRSTNAME];
                    dummyRow["LastName"] = studentIn[TEMPSHEET_LASTNAME];
                    dummyRow["Barcode"] = studentIn[TEMPSHEET_BARCODE];

                    dummyRow["InTime"] = studentIn[TEMPSHEET_INTIME];
                    dummyRow["#Subjects"] = studentIn[TEMPSHEET_SUBJECTS];

                    
                    dummyRow["Duration"] = duration.ToString(@"hh\:mm");
                    dummyRow["LastDay"] = studentIn[TEMPSHEET_LASTDAY];

                    dummyRow["#Completed"] = studentIn[TEMPSHEET_COMPLETED];
                    dummyRow["#Missing"] = studentIn[TEMPSHEET_MISSING];

                    dgdListing.ItemsSource = dummyTable.DefaultView;
                    dummyTable.Rows.Add(dummyRow);
                    dummyTable.DefaultView.Sort = "Duration DESC";
                }
            }

        }
        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        private void txtUpdate_TextChanged(object sender, TextChangedEventArgs e)
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
                    if (!isSignedIn(new string[3] { "A" + myConfirm.Number, myConfirm.FirstName, myConfirm.LastName }))
                    {
                        populateDataGrid(new string[3] { "A" + myConfirm.Number, myConfirm.FirstName, myConfirm.LastName });
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
                            signOut(new string[3] { "A" + myConfirm.Number, myConfirm.FirstName, myConfirm.LastName });
                        }
                        else if (MessageBox.Show("Student was recently signed in, sign out already?"
                            , "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            signOut(new string[3] { "A" + myConfirm.Number, myConfirm.FirstName, myConfirm.LastName });
                        }
                    }
                    // Throw error if not found.
                }
            }
            // Scanner is working so barcode = 'a' + number
            else
            {
                // Populate display after checking barcode
                if (!isSignedIn(new string[1] { txtUpdate.Text }))
                {
                    populateDataGrid(new string[1] { txtUpdate.Text });
                }
                else // confirm signout then do it
                {
                    int rowNum = 0;
                    foreach (DataRow row in dummyTable.Rows)
                    {
                        if (txtUpdate.Text == row["Barcode"].ToString())
                            break;
                        else
                            rowNum++;
                    }
                    DateTime timeIn = Convert.ToDateTime(dummyTable.Rows[rowNum]["InTime"].ToString());
                    if (DateTime.Now - timeIn > new TimeSpan(0,1,0))
                    {
                        signOut(new string[1] { txtUpdate.Text });
                    }
                    else if (MessageBox.Show("Student was recently signed in, sign out already?"
                        , "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        signOut(new string[1] { txtUpdate.Text });
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
            string range = "Sheet1!A1:M";

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
            for (int i = 0; i < 13; i++)
                pasteRange.Add(getStudents[rowNum - 1][i]);
            TimeSpan duration = DateTime.Now.Subtract(Convert.ToDateTime(getStudents[rowNum - 1][11]));

            pasteRange.Add((duration).ToString(@"hh\:mm"));

            List<Object> deleteRow = new List<object>();

            for (int i = 0; i < 13; i++)
                deleteRow.Add("");

            range = "Sheet1!A" + rowNum.ToString() + ":M" + rowNum.ToString();
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
            string tempRange = "Sheet1!A1:Z";
            string spreadsheetId = "1Lxn9qUxUbNWt3cI70CuTEIxCfgpxjAlZPd6ARph4oCM";  //STUDENT-database
            string spreadsheetId2 = "14j-XmVSs87CnsLX-TteOeIaAPak2G6_UTX6nU06kNWk"; //attendance record
            string barcode = "";
            string realEmail;
            string phone1;
            string carrier1;
            string phone2;
            string carrier2;
            int subjects = 0;
            string[] subjectsArray;

            IList<IList<Object>> values = getSpreadsheetInfo(spreadsheetId, "DB-Master!A1:AI");
            IList<IList<Object>> pasteValues = getSpreadsheetInfo(spreadsheetId2, tempRange);
            int rowNum = 1;

            // Get rowNum and set range
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    if (row[3].ToString() == checkValues[0])
                    {
                        if (checkValues.Length == 3)
                        {
                            if (row[5].ToString().ToUpper() == checkValues[1].ToUpper()
                            && row[7].ToString().ToUpper() == checkValues[2].ToUpper())
                            {
                                range = "DB-Master!A" + rowNum.ToString() + ":" + "AAA" + rowNum.ToString();
                                break; // additional autherntiacation
                            }
                        }
                        else
                        {
                            range = "DB-Master!A" + rowNum.ToString() + ":" + "AAA" + rowNum.ToString();
                            break;
                        }
                    }
                    else
                        rowNum++;
                }
            }

            // Get appropriate row
            values = getSpreadsheetInfo(spreadsheetId, range);
            var oblist = new List<Object> { };
            foreach (var row in values)
            {
                lastName = row[7].ToString();
                firstName = row[5].ToString();
                barcode = row[3].ToString();
                realEmail = row[13].ToString();
                phone1 = row[9].ToString();
                carrier1 = row[10].ToString();
                phone2 = row[15].ToString();
                carrier2 = row[16].ToString();
                subjectsArray = (row[2].ToString()).Split(',');
                if (subjectsArray.Length == 1)
                    subjects = 1;
                else if (subjectsArray.Length == 2)
                    subjects = 2;


                oblist.Add(lastName);
                oblist.Add(firstName);
                oblist.Add(barcode);
                oblist.Add(realEmail);
                oblist.Add(phone1);
                oblist.Add(carrier1);
                oblist.Add(phone2);
                oblist.Add(carrier2);
                oblist.Add("");                     // Updating Last-Day-In subsequently                
                oblist.Add(" ");                     // Completed
                oblist.Add(" ");                     // Missing
                oblist.Add(DateTime.Now.ToString("t"));
                oblist.Add(subjects.ToString());
            }

            // Find paste row num
            int pasteRowNum = 1;
            foreach (var row in pasteValues)
            {
                if (row.Count == 0)
                {
                    break;
                }
                pasteRowNum++;
            }

            tempRange = "Sheet1!A" + pasteRowNum.ToString() + ":Z" + pasteRowNum.ToString();
            updateSpreadsheetInfo(oblist, spreadsheetId2, tempRange);

            int rowNum2 = 1;
            // edit this to pull appropriate valued from database into temp and record sheet columns
            if (lastName != null && firstName != null)
            {                
                //get rowNum from Assignment Record spreadsheet
                IList<IList<Object>> values2 = getSpreadsheetInfo("1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s", "Test!B1:C");
                string assignmentRecordRange = "";
                foreach (var row in values2)
                {
                    if (row[1].ToString() == firstName && row[0].ToString() == lastName)
                    {
                        assignmentRecordRange = "Test!A" + rowNum2.ToString() + ":" + "AAA" + rowNum2.ToString();
                        break;
                    }
                    else
                    {
                        rowNum2++;
                    }
                }
                values2 = getSpreadsheetInfo("1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s", assignmentRecordRange);

                foreach (var row in values2)
                {
                    lastDayIn = row[7].ToString();
                }

                // Paste Last Day In to temp sheet
                List<Object> lastDay = new List<object>() { lastDayIn };
                updateSpreadsheetInfo(lastDay, spreadsheetId2, "Sheet1!I" + pasteRowNum.ToString());

                //get number of subjects
                dummyRow["FirstName"] = firstName;
                dummyRow["LastName"] = lastName;
                dummyRow["InTime"] = DateTime.Now.ToString("t");
                dummyRow["Duration"] = "00:00:00";
                dummyRow["LastDay"] = lastDayIn;
                dummyRow["Barcode"] = barcode;
                dummyRow["#Subjects"] = subjects.ToString();
            }
            dgdListing.ItemsSource = dummyTable.DefaultView;
            dummyTable.Rows.Add(dummyRow);
            dummyTable.DefaultView.Sort = "Duration DESC";
            txtUpdate.Clear();
            txtUpdate.Focus();

            // Set new last day in to today in assignmentrecord
            List<Object> today = new List<object>() { DateTime.Now.ToString("MM/dd") };
            updateSpreadsheetInfo(today, "1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s", "H" + rowNum2);
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
                string spreadsheetId = "1Lxn9qUxUbNWt3cI70CuTEIxCfgpxjAlZPd6ARph4oCM";
                /*
                var user = (from u in db.FStudentTables
                            where u.FirstName == firstName && u.LastName == lastName
                            select u).FirstOrDefault();
                            */

                IList<IList<Object>> values = getSpreadsheetInfo(spreadsheetId, "DB-Master!A1:AI");
                int rowNum = 1;
                if (values != null && values.Count > 0)
                {                  
                    foreach (var row in values)
                    {
                        if (drv["FirstName"].ToString() == row[5].ToString()
                        && drv["LastName"].ToString() == row[7].ToString())
                        {
                            break;
                        }
                        else
                            rowNum++;
                    }
                }

                var myRow = values[rowNum - 1];

                email = myRow[15].ToString();

                mail.From = new MailAddress("anthonyluukumon@gmail.com");
                mail.To.Add(email); // reg email
                mail.Subject = "Kumon HW notification";
                mail.Body = "Dear KUMON Parents,\n Your child, " + drv["FirstName"].ToString() + " " + drv["LastName"].ToString() +
                "attended center session today and turned in " + drv["#Completed"].ToString() +
                " assignment(s). We are still missing " + drv["#Missing"].ToString() + 
                " of his/her assignment(s).\n Per " +
                "your request, this automated message is sent to notify you of the " +
                "missing homework. Although it's common that students will miss an assignment " +
                "from time to time due to various activities, we hope these notifications will " +
                "help you identify whether your child is chronically missing homework. \n Regards, \n KUMON San Ramon North \n 925-318-1628";
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
                
                IList<IList<Object>>values = getSpreadsheetInfo("1Lxn9qUxUbNWt3cI70CuTEIxCfgpxjAlZPd6ARph4oCM", "DB-Master!A1:AI");

                int rowNum = 1;
                foreach (var row in values)
                {
                    if (drv["FirstName"].ToString() == row[5].ToString()
                        && drv["LastName"].ToString() == row[7].ToString())
                        break;
                    else
                        rowNum++;
                }

                var myRow = values[rowNum - 1];

                phone1 = myRow[11].ToString();
                carrier1 = myRow[12].ToString();
                phone2 = myRow[17].ToString();
                carrier2 = myRow[18].ToString();

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


        public string TimeLabel
        {
            get { return lblTime.Content.ToString(); }
            set { lblTime.Content = value; }
        }
    }
}
