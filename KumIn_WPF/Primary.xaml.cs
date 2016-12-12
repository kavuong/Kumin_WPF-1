/* 
When the window is constructed, all the columns are added to the DataTable. 
The myTimerTick function is configured to run every 15 seconds. Checks if
the current date is in the Attendance Record spreadsheet already, if not, 
the current date is appended to serve as delimiter for record of signins.
*/
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
        public const string DATABASE_SHEET = "1Lxn9qUxUbNWt3cI70CuTEIxCfgpxjAlZPd6ARph4oCM";
        public const string ASSIGNMENT_RECORD_SHEET = "1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s";

        public const string ATTENDANCE_SHEET_PERM_RECORD = "Record";
        public const string ATTENDANCE_SHEET_TEMP_RECORD = "Sheet1";
        public const string DATABASE_SHEET_RECORD = "DB-Master";
        public const string ASSIGNMENT_RECORD = "Test";


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

        public const int CONFIRMATION_INPUT = 3;
        public const int TEMPSHEET_COLUMNS = 13;

        public const int DATABASE_SUBJECTS = 2;
        public const int DATABASE_BARCODE = 3;
        public const int DATABASE_FIRST_NAME = 5;
        public const int DATABASE_LAST_NAME = 7;
        public const int DATABASE_PHONE1 = 9;
        public const int DATABASE_CARRIER1 = 10;
        public const int DATABASE_PICKUP1_VERIF = 11;
        public const int DATABASE_HW1_VERIF = 12;
        public const int DATABASE_EMAIL = 13;
        public const int DATABASE_PHONE2 = 15;
        public const int DATABASE_CARRIER2 = 16;
        public const int DATABASE_PICKUP2_VERIF = 17;
        public const int DATABASE_HW2_VERIF = 18;

        public const int ASSIGNMENT_RECORD_LAST_DAY = 7;
        public const int ASSIGNMENT_RECORD_DURATION = 12;

        public const int ONE_SUBJECT_LENGTH = 1;
        public const int TWO_SUBJECT_LENGTH = 2;


        //CheckValues array indices
        public const int barcodeIndex = 0;
        public const int firstNameIndex = 1;
        public const int lastNameIndex = 2;

        public const int FIRST_CHARACTER_INDEX = 0;
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


            if (!kuminConnection.isValuePresent(ATTENDANCE_SHEET, centerDates, DateTime.Now.ToString("MM/dd/yyyy"))) 
            {
                List<Object> date = new List<object>() { DateTime.Now.ToString("MM/dd/yyyy") };
                kuminConnection.append(date, ATTENDANCE_SHEET, centerDates);
            }

            this.dgdListing.CellEditEnding += new EventHandler<DataGridCellEditEndingEventArgs>(dgdListing_CellEditEnding);
        }







        //
        private void dgdListing_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int rowIndex = ((DataGrid)sender).ItemContainerGenerator.IndexFromContainer(e.Row);
            string text = ((TextBox)e.EditingElement).Text;
            int attRowNum = kuminConnection.getRowNum(ATTENDANCE_SHEET, ATTENDANCE_SHEET_TEMP_RECORD 
                + "!C1:C", dummyTable.Rows[rowIndex]["Barcode"].ToString());
            int assRowNum = kuminConnection.getRowNum(ASSIGNMENT_RECORD_SHEET, ASSIGNMENT_RECORD
                + "!D1:D", dummyTable.Rows[rowIndex]["Barcode"].ToString());
            

            if (e.Column.SortMemberPath.Equals("#Completed"))
            {
                List<Object> completed = new List<object>() { text };
                kuminConnection.update(completed, ATTENDANCE_SHEET, ATTENDANCE_SHEET_TEMP_RECORD + "!J" + attRowNum.ToString());
                kuminConnection.update(completed, ASSIGNMENT_RECORD_SHEET, ASSIGNMENT_RECORD + "!F" + assRowNum.ToString());
            }
            else if (e.Column.SortMemberPath.Equals("#Missing"))
            {
                List<Object> missing = new List<object>() { text };
                kuminConnection.update(missing, ATTENDANCE_SHEET, ATTENDANCE_SHEET_TEMP_RECORD + "!K" + attRowNum.ToString());
                kuminConnection.update(missing, ASSIGNMENT_RECORD_SHEET, ASSIGNMENT_RECORD + "!G" + assRowNum.ToString());
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


          
            if (tempSheetValues != null && tempSheetValues.Count > 0)
            {
                dummyTable.Clear();
                for (int i = 1; i < tempSheetValues.Count; i++)
                {
                    if (tempSheetValues[i].Count != 0)                      // Holes are created on signout causing unhandled exception
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
            if (!char.IsLetter(txtUpdate.Text[FIRST_CHARACTER_INDEX]))
            {
                Confirmation myConfirm = new Confirmation(txtUpdate.Text);
                if (myConfirm.ShowDialog() == true)
                {
                    // Populate display after checking firstname, lastname, number
                    if (!isSignedIn(new string[CONFIRMATION_INPUT] {
                        "A" + myConfirm.Number, myConfirm.FirstName, myConfirm.LastName }))
                    {
                        populateDataGrid(new string[CONFIRMATION_INPUT] { "A" + myConfirm.Number, myConfirm.FirstName, myConfirm.LastName });
                    }
                    else // confirm signout then do it
                    {
                        int studentRowNum = 0;
                        foreach (DataRow row in dummyTable.Rows)
                        {
                            string barcode = row["Barcode"].ToString();
                            if ("A" + txtUpdate.Text == barcode)
                                break;
                            else
                                studentRowNum++;
                        }
                        DateTime timeIn = Convert.ToDateTime(dummyTable.Rows[studentRowNum]["InTime"].ToString());
                        if (DateTime.Now - timeIn > new TimeSpan(0, 1, 0))
                        {
                            signOut(new string[CONFIRMATION_INPUT] { "A" + myConfirm.Number, myConfirm.FirstName, myConfirm.LastName });
                        }
                        else if (MessageBox.Show("Student was recently signed in, sign out already?"
                            , "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            signOut(new string[CONFIRMATION_INPUT] { "A" + myConfirm.Number, myConfirm.FirstName, myConfirm.LastName });
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
                    int studentRowNum = 0;
                    foreach (DataRow row in dummyTable.Rows)
                    {
                        if (txtUpdate.Text == row["Barcode"].ToString())
                            break;
                        else
                            studentRowNum++;
                    }
                    DateTime timeIn = Convert.ToDateTime(dummyTable.Rows[studentRowNum]["InTime"].ToString());
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







        private bool isSignedIn(string[] studentInput)
        {
            // Scanner works? checkvalues.count == 1
            string studentBarcode = studentInput[barcodeIndex];


            foreach (DataRow row in dummyTable.Rows)
            {
                string firstName = row["FirstName"].ToString();
                string lastName = row["LastName"].ToString();
                string barcode = row["Barcode"].ToString();
                if (barcode == studentBarcode)
                {
                    // Check if manual entry
                    if (studentInput.Count() == CONFIRMATION_INPUT)
                        return firstName.ToUpper() == studentInput[firstNameIndex].ToUpper()
                            && lastName.ToUpper() == studentInput[lastNameIndex].ToUpper();
                    else
                        return true;
                }
            }

            return false;
        }







        //*********************************************************************
        // Definition of signOut()                                            *
        // Takes three columns corresponding to first name, last name and     *
        // barcode from the temporary record spreadsheet. Then compares the   *
        // input of the checkValues array with the values in the three columns*
        // to ensure that the person being signed out actually is in the      *
        // temporary record spreadsheet. Once that verification is done, then *
        // the row is deleted from the temporary spreadsheet and moved to     *
        // permanent record.                                                  *
        //*********************************************************************
        private void signOut(string[] checkValues)
        {
            // Scanner works? checkvalues.count == 1.

            string tempColumn1 = ATTENDANCE_SHEET_TEMP_RECORD + "!A1:A"; // last name     
            string tempColumn2 = ATTENDANCE_SHEET_TEMP_RECORD + "!B1:B"; // first name    
            string tempColumn3 = ATTENDANCE_SHEET_TEMP_RECORD + "!C1:C"; // barcode


            int rowNum = 0;
            List<Object> pasteRange = new List<object>() { };

            if (checkValues.Count() == 3)
            {
                rowNum = kuminConnection.getRowNum(ATTENDANCE_SHEET, tempColumn1, checkValues[lastNameIndex].ToUpper(), tempColumn2
                , checkValues[firstNameIndex].ToUpper(), tempColumn3, checkValues[barcodeIndex].ToUpper());
            }
            else if (checkValues.Count() == 1)
            {
                rowNum = kuminConnection.getRowNum(ATTENDANCE_SHEET, tempColumn3, checkValues[barcodeIndex].ToUpper());
            }
                           
            
            string signOutRange = ATTENDANCE_SHEET_TEMP_RECORD + "!A" + rowNum.ToString() + ":AAA" + rowNum.ToString();
            IList<IList<Object>> signOutStudent = kuminConnection.get(ATTENDANCE_SHEET, signOutRange);
            TimeSpan duration = DateTime.Now.Subtract(Convert.ToDateTime(signOutStudent[0][TEMPSHEET_INTIME]));

            for (int i = 0; i < TEMPSHEET_COLUMNS; i++)
            {
                pasteRange.Add(signOutStudent[0][i]);
                
            }

            pasteRange.Add((duration).ToString(@"hh\:mm"));

            List<Object> deleteRow = new List<object>();

            for (int i = 0; i < TEMPSHEET_COLUMNS; i++)
                deleteRow.Add("");

            string deleteStudentRange = ATTENDANCE_SHEET_TEMP_RECORD + "!A" + rowNum.ToString() + ":M" + rowNum.ToString();
            kuminConnection.update(deleteRow, ATTENDANCE_SHEET, deleteStudentRange);

            // Transfer row to permanent record
            string addStudentRange = ATTENDANCE_SHEET_PERM_RECORD + "!A1:A";

            kuminConnection.append(pasteRange, ATTENDANCE_SHEET, addStudentRange);

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
            string databaseRange = DATABASE_SHEET_RECORD + "!A1:AI";
            string barcodeRange = DATABASE_SHEET_RECORD + "!D1:D";
            string firstNameRange = DATABASE_SHEET_RECORD + "!F1:F";
            string lastNameRange = DATABASE_SHEET_RECORD + "!H1:H";
            
            string tempSheetColumn = ATTENDANCE_SHEET_TEMP_RECORD + "!A1:A";
            string barcode = "";
            string realEmail;
            string phone1;
            string carrier1;
            string phone2;
            string carrier2;
            int subjects = 0;
            string[] subjectsArray;
            string studentRowRange;
            int rowNum = 0;
            int assignmentRecordRowNum = 0;
            IList<IList<Object>> databaseValues = kuminConnection.get(DATABASE_SHEET, databaseRange);
            IList<IList<Object>> assignmentRecordValues = kuminConnection.get(ASSIGNMENT_RECORD_SHEET, ASSIGNMENT_RECORD + "!B1:C");

            IList<IList<Object>> studentRow;
            IList<IList<Object>> assignmentRecordRow;



            if (databaseValues != null && databaseValues.Count > 0)
            {
                if (checkValues.Length == 1)
                    rowNum = kuminConnection.getRowNum(DATABASE_SHEET, barcodeRange, checkValues[barcodeIndex]);
                else if (checkValues.Length == CONFIRMATION_INPUT)
                    rowNum = kuminConnection.getRowNum(DATABASE_SHEET, barcodeRange, checkValues[barcodeIndex]
                        , firstNameRange, checkValues[firstNameIndex], lastNameRange, checkValues[lastNameIndex]);
            }

            studentRowRange = DATABASE_SHEET_RECORD + "!A" + rowNum.ToString() + ":AAA" + rowNum.ToString();


            // Get appropriate row
            studentRow = kuminConnection.get(DATABASE_SHEET, studentRowRange);
            var studentRowList = new List<Object> { };
            foreach (var row in studentRow)
            {
                lastName = row[DATABASE_LAST_NAME].ToString();
                firstName = row[DATABASE_FIRST_NAME].ToString();
                barcode = row[DATABASE_BARCODE].ToString();
                realEmail = row[DATABASE_EMAIL].ToString();
                phone1 = row[DATABASE_PHONE1].ToString();
                carrier1 = row[DATABASE_CARRIER1].ToString();
                phone2 = row[DATABASE_PHONE2].ToString();
                carrier2 = row[DATABASE_CARRIER2].ToString();
                subjectsArray = (row[DATABASE_SUBJECTS].ToString()).Split(',');
                if (subjectsArray.Length == ONE_SUBJECT_LENGTH)
                    subjects = ONE_SUBJECT_LENGTH;
                else if (subjectsArray.Length == TWO_SUBJECT_LENGTH)
                    subjects = TWO_SUBJECT_LENGTH;


                studentRowList.Add(lastName);
                studentRowList.Add(firstName);
                studentRowList.Add(barcode);
                studentRowList.Add(realEmail);
                studentRowList.Add(phone1);
                studentRowList.Add(carrier1);
                studentRowList.Add(phone2);
                studentRowList.Add(carrier2);
                studentRowList.Add("");                     // Updating Last-Day-In subsequently                
                studentRowList.Add("");                     // Completed
                studentRowList.Add("");                     // Missing
                studentRowList.Add(DateTime.Now.ToString("t"));
                studentRowList.Add(subjects.ToString());
            }

            // Find paste row num
            int pasteRowNum = kuminConnection.getRowNum(ATTENDANCE_SHEET, tempSheetColumn, "");

            string pasteRowRange = ATTENDANCE_SHEET_TEMP_RECORD + "!A" + pasteRowNum.ToString() + ":Z" + pasteRowNum.ToString();
            kuminConnection.update(studentRowList, ATTENDANCE_SHEET, pasteRowRange);

            
            // edit this to pull appropriate valued from database into temp and record sheet columns
            if (lastName != null && firstName != null)
            {
                //get rowNum from Assignment Record spreadsheet
                string assignmentRecordRange;
                assignmentRecordRowNum = kuminConnection.getRowNum(ASSIGNMENT_RECORD_SHEET, ASSIGNMENT_RECORD + "!B1:B", lastName
                    , ASSIGNMENT_RECORD + "!C1:C", firstName);

                assignmentRecordRange = ASSIGNMENT_RECORD + "!A" + assignmentRecordRowNum.ToString() 
                    + ":AAA" + assignmentRecordRowNum.ToString();
                
                assignmentRecordRow = kuminConnection.get(ASSIGNMENT_RECORD_SHEET, assignmentRecordRange);

                foreach (var row in assignmentRecordRow)
                {
                    lastDayIn = row[ASSIGNMENT_RECORD_LAST_DAY].ToString();
                }

                // Paste Last Day In to temp sheet
                List<Object> lastDay = new List<object>() { lastDayIn };
                kuminConnection.update(lastDay, ATTENDANCE_SHEET, ATTENDANCE_SHEET_TEMP_RECORD + "!I" + pasteRowNum.ToString());

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
            kuminConnection.update(today, ASSIGNMENT_RECORD_SHEET, ASSIGNMENT_RECORD + "!H" + assignmentRecordRowNum);
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
                string hwEmailVerif = "";
                int rowNum = 1;
                /*
                var user = (from u in db.FStudentTables
                            where u.FirstName == firstName && u.LastName == lastName
                            select u).FirstOrDefault();
                            */

                IList<IList<Object>> databaseValues = kuminConnection.get(DATABASE_SHEET, DATABASE_SHEET_RECORD + "!A1:Z");
                
                if (databaseValues != null && databaseValues.Count > 0)
                {
                    rowNum = kuminConnection.getRowNum(DATABASE_SHEET, DATABASE_SHEET_RECORD + 
                        "!F1:F", firstName, DATABASE_SHEET_RECORD + "H1:H", lastName);                    
                }

                var myRow = databaseValues[rowNum - 1];

                email = myRow[DATABASE_EMAIL].ToString();
                hwEmailVerif = myRow[DATABASE_HW1_VERIF].ToString();

                if (hwEmailVerif.ToUpper() == "YES")
                {
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
                    "help you identify whether your child is chronically missing homework. \n Regards, \n KUMON San Ramon North \n 925-318-1628";
                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("anthonyluukumon@gmail.com"
                        , "letmeout");
                    SmtpServer.EnableSsl = true;

                    SmtpServer.Send(mail);
                    MessageBox.Show("HW Email Sent", "Success");
                }
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
                string pickUpVerif1 = "";
                string pickUpVerif2 = "";
                string carrier1 = "";
                string carrier2 = "";
                string carrierString1 = "";
                string carrierString2 = "";
                int rowNum = 1;

                DataRowView drv = (DataRowView)dgdListing.SelectedItem;
                string firstName = (drv["FirstName"]).ToString();
                string lastName = (drv["LastName"]).ToString();

                
                IList<IList<Object>>values = kuminConnection.get(DATABASE_SHEET, DATABASE_SHEET_RECORD + "!A1:Z");
                rowNum = kuminConnection.getRowNum(DATABASE_SHEET, DATABASE_SHEET_RECORD +
                "!F1:F", firstName, DATABASE_SHEET_RECORD + "H1:H", lastName);

                var studentRow = values[rowNum - 1];

                
                phone1 = studentRow[DATABASE_PHONE1].ToString();
                carrier1 = studentRow[DATABASE_CARRIER1].ToString();
                pickUpVerif1 = studentRow[DATABASE_PICKUP1_VERIF].ToString();
                pickUpVerif2 = studentRow[DATABASE_PICKUP2_VERIF].ToString();
                phone2 = studentRow[DATABASE_PHONE2].ToString();
                carrier2 = studentRow[DATABASE_CARRIER2].ToString();

                carrierString1 = returnCarrierString(carrier1);
                carrierString2 = returnCarrierString(carrier2);

                if (phone1 != null && carrierString1 != null && pickUpVerif1.ToUpper() == "YES")
                {
                    mail1.From = new MailAddress("kumonsrn@gmail.com");
                    mail1.To.Add(phone1 + carrierString1); //phone
                    mail1.Subject = "Kumon Reminder";
                    mail1.Body = "Your child - " + firstName + " " + lastName + " - is ready to be picked up";
                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("anthonyluukumon@gmail.com"
                           , "letmeout");
                    SmtpServer.EnableSsl = true;
                    SmtpServer.Send(mail1);
                }

                if (phone2 != "" && carrierString2 != "" && pickUpVerif2.ToUpper() == "YES")
                {
                    mail2.From = new MailAddress("anthonyluukumon@gmail.com");
                    mail2.To.Add(phone2 + carrierString2); //phone
                    mail2.Subject = "Kumon Reminder";
                    mail2.Body = "Your child - " + firstName + " " + lastName + " - is ready to be picked up";
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
        /*
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
        */

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

        public string returnCarrierString(string carrier)
        {
            if (carrier == "Verizon")
                return "@vtext.com";

            else if (carrier == "AT&T")
                return "@txt.att.net";

            else if (carrier == "Sprint")
                return "@messaging.sprintpcs.com";

            else if (carrier == "T-Mobile")
                return "@tmomail.net";

            else if (carrier == "Cricket")
                return "@sms.mycricket.com";

            else if (carrier == "Ultra")
                return "@mailmymobile.net";

            return null;

        }
    }
}
