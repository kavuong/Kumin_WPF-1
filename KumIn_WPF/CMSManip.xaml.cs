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
using System.Data;

namespace KumIn_WPF
{
    /// <summary>
    /// Interaction logic for CMSManip.xaml
    /// </summary>
    public partial class CMSManip : Window
    {
        DataTable dtListing = new DataTable();
        SpreadsheetConnection CMSConnection = new SpreadsheetConnection();

        public const string ASSIGNMENT_SHEET = "1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s";
        public const string ASSIGNMENT_SHEET_RECORDS = "Test";


        public const int ASSIGNSHEET_NUMASSIGN = 9;
        public const int ASSIGNSHEET_CMSMANIP = 12;

        public CMSManip()
        {
            InitializeComponent();

            dtListing.Columns.Add("NumAssign");
            dtListing.Columns.Add("Assigned");
            dtListing.Columns.Add("Completed");
            dtListing.Columns.Add("Level");
            dtListing.Columns.Add("Sheet#");
        }









        public void populate(int rowNum)
        {
            string studentCells = ASSIGNMENT_SHEET_RECORDS + 
                "!A" + rowNum.ToString() + ":AAA" + rowNum.ToString();
            IList<IList<Object>> studentRecord = CMSConnection.get(ASSIGNMENT_SHEET, studentCells);

            foreach (var row in studentRecord)
            {
                lblName.Content = row[2].ToString() + row[1].ToString();
                lblSubject.Content = row[4];

                int cellIndex = ASSIGNSHEET_CMSMANIP + 1;
                for (int i = 0; i < int.Parse(row[ASSIGNSHEET_NUMASSIGN].ToString()); i++)
                {
                    DataRow dr = dtListing.NewRow();
                    string[] levelSheet = row[cellIndex + 1].ToString().Split(' ');


                    dr["NumAssign"] = i;
                    dr["Assigned"] = row[cellIndex];
                    dr["Level"] = levelSheet[0];
                    dr["Sheet#"] = levelSheet[1];

                    cellIndex += 2;
                }
                dgdListing.ItemsSource = dtListing.DefaultView;
                lblDateRange.Content = dtListing.Rows[0][1].ToString() + "-"
                    + dtListing.Rows[int.Parse(row[ASSIGNSHEET_NUMASSIGN].ToString()) - 1][1].ToString();
            }

        }

        private void btnNextRecord_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
