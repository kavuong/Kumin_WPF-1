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
        SpreadsheetConnection kuminConnection = new SpreadsheetConnection();
        public const string DATABASE_SHEET = "1Gav1wmBzJ9xwIwlLxRkQHy32d2pzOuJ6CCO_3jZMJDY";
        public const string DATABASE_SHEET_RECORD = "Database";

        public const int DATABASE_FIRST_NAME = 5;
        public const int DATABASE_LAST_NAME = 7;
        public Confirmation(string number)
        {
            InitializeComponent();
            Number = number;
            string[] nameArray = namesDisplayed();
            lblFirstNameText.Content = nameArray[0];
            lblLastNameText.Content = nameArray[1];
        }

        private void btnSignIn_Click(object sender, RoutedEventArgs e)
        {

            this.DialogResult = true;
            this.Close();
            
        }

        private void btnReconfirm_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            MessageBox.Show("Please re-input your barcode.");
            this.Close();
        }
        private string[] namesDisplayed()
        {
            int rowNum = kuminConnection.getRowNum(DATABASE_SHEET, DATABASE_SHEET_RECORD + "!D1:D", Barcode);
            string[] returnArray = new string[2];
            string studentRowRange = DATABASE_SHEET_RECORD + "!A" + rowNum.ToString() + ":AAA" + rowNum.ToString();
            var studentRow = kuminConnection.get(DATABASE_SHEET, studentRowRange);
            foreach (var row in studentRow)
            {
                returnArray[0] = row[DATABASE_FIRST_NAME].ToString();
                returnArray[1] = row[DATABASE_LAST_NAME].ToString();
            }
            
            return returnArray;

        }

        public string FirstName
        {
            get { return lblFirstNameText.Content.ToString(); }
        }

        public string LastName
        {
            get { return lblLastNameText.Content.ToString(); }
        }

        public string Number
        {
            get { return lblNumberText.Content.ToString(); }
            set { lblNumberText.Content = value.ToString(); }
        }

        public string Barcode
        {
            get
            {
                if (Number.Length == 1)
                    return "A000" + Number.ToString();
                else if (Number.Length == 2)
                    return "A00" + Number.ToString();
                else if (Number.Length == 3)
                    return "A0" + Number.ToString();
                else if (Number.Length == 4)
                    return "A" + Number.ToString();
                else
                    throw new NullReferenceException();
            }
        }


    }
}
