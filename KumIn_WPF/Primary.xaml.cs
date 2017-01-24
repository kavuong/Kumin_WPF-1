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
        public const string ATTENDANCE_SHEET = "1PQzK4uaL5k7V-P9W7eUNxpriCz_WHSoEVuEp-wek5ys";
        public const string DATABASE_SHEET = "1Gav1wmBzJ9xwIwlLxRkQHy32d2pzOuJ6CCO_3jZMJDY";
        public const string ASSIGNMENT_RECORD_SHEET = "1_i3YFC0DT44WbIqPMfBFv74p1WXi5NBMDwetqUPYoqY";

        public const string ATTENDANCE_SHEET_PERM_RECORD = "DayAttendance";
        public const string ATTENDANCE_SHEET_TEMP_RECORD = "StudentsIn";
        public const string DATABASE_SHEET_RECORD = "Database";
        public const string ASSIGNMENT_RECORD = "Record";


        // Column Indexes
        public const int DATAGRID_FIRSTNAME = 0;
        public const int DATAGRID_LASTNAME = 1;
        public const int DATAGRID_COMPLETED = 5;
        public const int DATAGRID_MISSING = 6;

        public const int TEMPSHEET_FIRSTNAME = 1;
        public const int TEMPSHEET_LASTNAME = 0;
        public const int TEMPSHEET_BARCODE = 2;
        public const int TEMPSHEET_PHONE1 = 4;
        public const int TEMPSHEET_CARRIER1 = 5;
        public const int TEMPSHEET_PHONE2 = 6;
        public const int TEMPSHEET_CARRIER2 = 7;
        public const int TEMPSHEET_INTIME = 11;
        public const int TEMPSHEET_SUBJECTS = 12;
        public const int TEMPSHEET_LASTDAY = 8;
        public const int TEMPSHEET_COMPLETED = 9;
        public const int TEMPSHEET_MISSING = 10;
        public const int TEMPSHEET_OUT1_VERIF = 13;
        public const int TEMPSHEET_OUT2_VERIF = 14;

        public const int CONFIRMATION_INPUT = 3;
        public const int TEMPSHEET_COLUMNS = 15;

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
            try
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
                dummyTable.Columns.Add("OutEmail", typeof(bool));

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
                this.Show();
                myTimer_Tick(new int(), new int());
            }
            catch(System.Net.Http.HttpRequestException ex)
            {
                MessageBox.Show("Error: Not connected to the internet, please connect and restart.");
                this.Close();
            }
        }







        //
        private void dgdListing_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditingElement is TextBox)
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

                txtUpdate.Focus();
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

                        if (studentIn[TEMPSHEET_OUT1_VERIF].ToString() == "YES"
                            || studentIn[TEMPSHEET_OUT2_VERIF].ToString() == "YES")
                            dummyRow["OutEmail"] = true;
                        else
                            dummyRow["OutEmail"] = false;

                        dgdListing.ItemsSource = dummyTable.DefaultView;
                        dummyTable.Rows.Add(dummyRow);
                        dummyTable.DefaultView.Sort = "Duration DESC";
                    }
                }
            }
        }







        private void txtUpdate_TextChanged(object sender, TextChangedEventArgs e)
        {

        }






        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // If scanner is not working ==> dialog confirm
                if (!char.IsLetter(txtUpdate.Text[FIRST_CHARACTER_INDEX]))
                {
                    Confirmation myConfirm = new Confirmation(txtUpdate.Text);
                    if (myConfirm.ShowDialog() == true)
                    {
                        // Populate display after checking firstname, lastname, number
                        if (!isSignedIn(new string[CONFIRMATION_INPUT] {
                        "A" + (int.Parse(myConfirm.Number)).ToString("D4"), myConfirm.FirstName, myConfirm.LastName }))
                        {
                            populateDataGrid(new string[CONFIRMATION_INPUT] { "A" + (int.Parse(myConfirm.Number)).ToString("D4")
                            , myConfirm.FirstName, myConfirm.LastName });
                        }
                        else // confirm signout then do it
                        {
                            MessageBox.Show("Student is already signed in. Click button to sign out");
                            txtUpdate.Clear();
                            txtUpdate.Focus();
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
                        MessageBox.Show("Student is already signed in. Click button to sign out");
                        txtUpdate.Clear();
                        txtUpdate.Focus();
                    }
                }
                txtUpdate.Clear();
            }
            catch (IndexOutOfRangeException iEx)
            {
                MessageBox.Show("Please enter a barcode.");
            }
            catch (NullReferenceException nEx)
            {
                MessageBox.Show("For automatic sign-in, please enter a valid barcode. \n\nFor manual sign-in, " +
                    "please enter your first name, last name and barcode spelled correctly.");
            }
        }




        
        private void dgdListing_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }








        private void onOutChecked(object sender, RoutedEventArgs e)
        {
            if (dgdListing.SelectedIndex != -1)
            {
                int rowIndex = dgdListing.SelectedIndex;
                int tempRowNum = kuminConnection.getRowNum(ATTENDANCE_SHEET, ATTENDANCE_SHEET_TEMP_RECORD + "!C1:C"
                    , dummyTable.Rows[rowIndex]["Barcode"].ToString());

                kuminConnection.update(new List<object> { "YES", "YES" }, ATTENDANCE_SHEET, ATTENDANCE_SHEET_TEMP_RECORD
                    + "!N" + tempRowNum.ToString() + ":O" + tempRowNum.ToString());
            }

        }






        private void onOutUnchecked(object sender, RoutedEventArgs e)
        {
            if (dgdListing.SelectedIndex != -1)
            {
                int rowIndex = dgdListing.SelectedIndex;
                int tempRowNum = kuminConnection.getRowNum(ATTENDANCE_SHEET, ATTENDANCE_SHEET_TEMP_RECORD + "!C1:C"
                    , dummyTable.Rows[rowIndex]["Barcode"].ToString());

                kuminConnection.update(new List<object> { "NO", "NO" }, ATTENDANCE_SHEET, ATTENDANCE_SHEET_TEMP_RECORD
                    + "!N" + tempRowNum.ToString() + ":O" + tempRowNum.ToString());
            }

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
                rowNum = kuminConnection.getRowNum(ATTENDANCE_SHEET, tempColumn1, checkValues[lastNameIndex]
                    , tempColumn2, checkValues[firstNameIndex], tempColumn3, checkValues[barcodeIndex].ToUpper());
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

            string deleteStudentRange = ATTENDANCE_SHEET_TEMP_RECORD + "!A" + rowNum.ToString() + ":O" + rowNum.ToString();
            kuminConnection.update(deleteRow, ATTENDANCE_SHEET, deleteStudentRange);

            // Transfer row to permanent record
            string addStudentRange = ATTENDANCE_SHEET_PERM_RECORD + "!A1:A";

            kuminConnection.append(pasteRange, ATTENDANCE_SHEET, addStudentRange);

            txtUpdate.Focus();
            myTimer_Tick(this, this);
            MessageBox.Show(signOutStudent[0][1] + " " + signOutStudent[0][0] 
                + " is now signed out. Goodbye!");
        }





        
        private void signOut(object sender, RoutedEventArgs e)
        {
            try
            {
                //int rowIndex = dgdListing.ItemContainerGenerator
                //    .IndexFromContainer((DataGridRow)((FrameworkElement)sender).DataContext);
                int rowIndex = dgdListing.SelectedIndex;
                string barcode = dummyTable.Rows[rowIndex]["Barcode"].ToString();
                CheckBox outBox = dgdListing.Columns[8].GetCellContent(dgdListing.Items[rowIndex]) as CheckBox;

                if (outBox.IsChecked.Value)
                    sendOutEmail(dgdListing.SelectedItem as DataRowView);

                signOut(new string[] { barcode });

                txtUpdate.Focus();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Something wrong happened here. Maybe double clicked the button?\n\n"
                    + "Admin info: " + ex.Message);
            }
        }








        //************************************************
        // Definition of populateDataGrid()
        // Takes a string array of barcode as input. First name and last name are also 
        // in the array if manual input is enabled in btnUpdate_Click(). Initializes 
        // the ranges for the entire database and barcode, first name and last name
        // columns. Gets the row number in the database corresponding to the values
        // in the input array, and gets the row corresponding to the row num. Sets
        // last name, first name, barcode, and all information needed from the
        // database row, and adds these values to a new List<Object>



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
            string isPhone1 = "";
            string isPhone2 = "";
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

                if (row[DATABASE_PICKUP1_VERIF].ToString() == "YES")
                    isPhone1 = "YES";
                else
                    isPhone1 = "NO";

                if (row[DATABASE_PICKUP2_VERIF].ToString() == "YES")
                    isPhone2 = "YES";
                else
                    isPhone2 = "NO";


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
                studentRowList.Add(isPhone1);
                studentRowList.Add(isPhone2);
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
                assignmentRecordRowNum = kuminConnection.getRowNum(ASSIGNMENT_RECORD_SHEET, ASSIGNMENT_RECORD + "!D1:D", barcode);

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
                if (isPhone1 == "YES"
                            || isPhone2 == "YES")
                    dummyRow["OutEmail"] = true;
                else
                    dummyRow["OutEmail"] = false;
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






        private void sendOutEmail(DataRowView drv)
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
                string firstName = drv["FirstName"].ToString();
                string lastName = drv["LastName"].ToString();
                int rowNum = 1;
                string barcode = (drv["Barcode"]).ToString();

                rowNum = kuminConnection.getRowNum(ATTENDANCE_SHEET, ATTENDANCE_SHEET_TEMP_RECORD 
                    + "!B1:B", firstName, ATTENDANCE_SHEET_TEMP_RECORD + "!A1:A", lastName);
                IList<IList<Object>>values = kuminConnection.get(ATTENDANCE_SHEET
                    , ATTENDANCE_SHEET_TEMP_RECORD + "!A" + rowNum.ToString() +  ":AAA" + rowNum.ToString());

                var studentRow = values[0];

                
                phone1 = studentRow[TEMPSHEET_PHONE1].ToString();
                carrier1 = studentRow[TEMPSHEET_CARRIER1].ToString();
                phone2 = studentRow[TEMPSHEET_PHONE2].ToString();
                carrier2 = studentRow[TEMPSHEET_CARRIER2].ToString();

                carrierString1 = returnCarrierString(carrier1);
                carrierString2 = returnCarrierString(carrier2);


                pickUpVerif1 = studentRow[TEMPSHEET_OUT1_VERIF].ToString();
                pickUpVerif2 = studentRow[TEMPSHEET_OUT2_VERIF].ToString();

                if (phone1 != "" && carrierString1 != "" && pickUpVerif1 == "YES")
                {
                    mail1.From = new MailAddress("kumonsrn@gmail.com");
                    mail1.To.Add(phone1 + carrierString1); //phone
                    mail1.CC.Add("kumonsrn@gmail.com");
                    mail1.Subject = "Kumon Reminder";
                    mail1.Body = "Your child - " + firstName + " " + lastName + " - is ready to be picked up";
                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("kumonsrn@gmail.com"
                           , "letmeout");
                    SmtpServer.EnableSsl = true;
                    SmtpServer.Send(mail1);
                }

                if (phone2 != "" && carrierString2 != "" && pickUpVerif2 == "YES")
                {
                    mail2.From = new MailAddress("kumonsrn@gmail.com");
                    mail2.To.Add(phone2 + carrierString2); //phone
                    mail2.CC.Add("kumonsrn@gmail.com");
                    mail2.Subject = "Kumon Reminder";
                    mail2.Body = "Your child - " + firstName + " " + lastName + " - is ready to be picked up";
                    SmtpServer.Port = 587;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("kumonsrn@gmail.com"
                           , "letmeout");
                    SmtpServer.EnableSsl = true;
                    SmtpServer.Send(mail2);
                }               
                MessageBox.Show("Text(s) sent to the parents of " + firstName + " " + lastName + "." 
                    ,"Success");                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Texts not sent, perhaps parent is not signed up?\n\nAdmin details: " 
                    + ex.Message, "Error");
            }
        }              

        private void btnCMSManip_Click(object sender, RoutedEventArgs e)
        {
            CMSManip myCMSManip = new CMSManip();
        }




        private void btnAddNewStudent_Click(object sender, RoutedEventArgs e)
        {
            AssignWork myAssignWork = new AssignWork();
            myAssignWork.Show();
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

        private void btnHWEmail_Click(object sender, RoutedEventArgs e)
        {
            HWEmails myHWEmails = new HWEmails();
            myHWEmails.Show();
        }


        }
    }

