
/*
When this window is constructed, a datatable is created with columns.
*/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
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
using System.Printing;



namespace KumIn_WPF
{
    /// <summary>
    /// Interaction logic for AssignWork.xaml
    /// </summary>
    public partial class AssignWork : Window
    {
        DataTable dt = new DataTable();
        SpreadsheetConnection assignConnection = new SpreadsheetConnection();
        private string dayOff = "";
        private string pattern = "";
        private string subject = "";

        // Constants
        public const string ASSIGNMENT_SHEET = "1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s";
        public const string DATABASE_SHEET = "1Lxn9qUxUbNWt3cI70CuTEIxCfgpxjAlZPd6ARph4oCM";

        public const string ASSIGNMENT_SHEET_RECORD = "Test";
        public const string DATABASE_SHEET_RECORD = "DB-Master";

        public const int DATABASE_BARCODE = 3;

        public const int ASSIGNSHEET_FIRSTNAME = 2;
        public const int ASSIGNSHEET_LASTNAME = 1;
        public const int ASSIGNSHEET_SUBJECT = 4;
        public const int ASSIGNSHEET_NUMASSIGN = 9;
        public const int ASSIGNSHEET_LASTDAY = 7;
        public const int ASSIGNSHEET_DAYOFF = 11;
        public const int ASSIGNSHEET_PATTERN = 10;
        public const int ASSIGNSHEET_CMSMANIP = 12;

        public const int DATAGRID_ASSIGNDATE = 1;
        public const int DATAGRID_LEVEL = 3;
        public const int DATAGRID_PAGES = 4;


        public AssignWork()
        {
            InitializeComponent();


            dt.Columns.Add("#");
            dt.Columns.Add("Assigned");
            dt.Columns.Add("Completed");
            dt.Columns.Add("Level");
            dt.Columns.Add("Sheet#");

            this.dgdFormat.CellEditEnding += new EventHandler
                <DataGridCellEditEndingEventArgs>(dgdFormat_CellEditEnding);
        }







        //************************************************************************************
        // Definition of event handler dgdFormat_CellEditEnding()                            *
        // This handles when the level or pages are changed inside the view. If the level is *
        // changed, that level is continued on for the remainder of the rows. If the pages   *
        // are changed, then the same pattern is continued on with the pages picking up from *
        // what was left off. The underlying datatable itself is updated upon change. If     *
        // the student has corrections only (C), pages pick up at the next row               *
        //************************************************************************************
        private void dgdFormat_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int rowIndex = ((DataGrid)sender).ItemContainerGenerator.IndexFromContainer(e.Row);
            string text = ((TextBox)e.EditingElement).Text;

            if (e.Column.SortMemberPath.Equals("Level"))
            {
                // enter this level++ for subsequent rows
                KumonLevel nextLevel = new KumonLevel(Subject, text);

                for (int i = rowIndex; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["Level"] = nextLevel.Level;
                }
                dgdFormat.ItemsSource = dt.DefaultView;
            }
            else if (e.Column.SortMemberPath.Equals("Sheet#"))
            {

                string[] pages;
                int startPage;
                int endPage;
                int currentSheet;
                string newPattern;

                pages = text.Split('-');

                if (text.ToUpper() == "C")
                {
                    pages = dt.Rows[rowIndex]["Sheet#"].ToString().Split('-');
                    startPage = int.Parse(pages[0]);
                    endPage = int.Parse(pages[1]);

                    currentSheet = int.Parse(pages[0]);
                    dt.Rows[rowIndex]["Sheet#"] = "Corr. Only";
                    rowIndex++;
                }
                else
                {
                    startPage = int.Parse(pages[0]);
                    endPage = int.Parse(pages[1]);
                    currentSheet = startPage;
                }

                newPattern = getNewPattern(startPage, endPage);



                for (int i = rowIndex; i < dt.Rows.Count; i++)
                {
                    int nextSheet;

                    if (currentSheet > 200)
                        currentSheet -= 200;

                    nextSheet = calculateNextSheet(newPattern, currentSheet);

                    dt.Rows[i]["Sheet#"] = currentSheet.ToString() + "-"
                        + (nextSheet - 1).ToString();

                    currentSheet = nextSheet;

                }
                dgdFormat.ItemsSource = dt.DefaultView;
            }
        }







        //****************************************************************
        // Definition of event handler btnPrintRecord_Click()            *
        // Data is written to the spreadsheet and the grid is printed    *
        //****************************************************************
        private void btnPrintRecord_Click(object sender, RoutedEventArgs e)
        {
            writeData(getDateAssign());
            print();
        }
        






        //***********************************************************
        // Definition of event handler txtBarcode_KeyDown()         *
        // If enter is pressed, readAndPopulate is called.          *
        //***********************************************************
        private void txtBarcode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                readAndPopulate();
        }







        //*************************************************************************
        // Definition of event handler txtNumAssign_TextChanged()                 *
        // Data is updated in the view's grid.                                    *
        //*************************************************************************
        private void txtNumAssign_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateData();
        }







        //*************************************************************************
        // Definition of event handler txtStartDate_TextChanged()                 *
        // Data is updated in the view's grid.                                    *
        //*************************************************************************
        private void txtStartDate_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateData();
        }







        //***********************************************************************
        // Definition of readAndPopulate()                                      *
        // Searches through the Assignment Record spreadsheet to find the row   *
        // that matches the entered barcode and subject. Information from here  *
        // is populated into the fields and the dataGrid is updated to continue *
        // on from the previous record.                                         *
        //***********************************************************************
        private void readAndPopulate()
        {
            if (txtBarcode.Text != "" && Subject != "")
            {
                string barcodes = ASSIGNMENT_SHEET_RECORD + "!D1:D";
                string subjectColumn = ASSIGNMENT_SHEET_RECORD + "!E1:E";
                string sheetCells;
                IList<IList<Object>> sheet;

                bool found = assignConnection.isValuePresent(ASSIGNMENT_SHEET
                    , barcodes, Barcode);
                int rowNum = assignConnection.getRowNum(ASSIGNMENT_SHEET, barcodes
    , Barcode, subjectColumn, Subject);


                sheetCells = "Test!A" + rowNum.ToString() + ":" + "AAA" + rowNum.ToString();
                sheet = assignConnection.get(ASSIGNMENT_SHEET, sheetCells);

                try
                {
                    foreach (var row in sheet)
                    {
                        int lastDateIndex = row.Count - 2;
                        DateTime today = Convert.ToDateTime(row[ASSIGNSHEET_LASTDAY])
                            + new TimeSpan(1, 0, 0, 0);
                        string[] subStringLevel; 
                            
                        string[] subStringPage;

                        lblName.Content = row[ASSIGNSHEET_FIRSTNAME].ToString() + " "
                            + row[ASSIGNSHEET_LASTNAME].ToString();
                        if (!found)
                            throw new EntryPointNotFoundException();


                        //lblSubject.Content = row[ASSIGNSHEET_SUBJECT].ToString();
                        txtNumAssign.Text = row[ASSIGNSHEET_NUMASSIGN].ToString();
                        txtStartDate.Text = (today).ToString("MM/dd");

                        subStringLevel = row[ASSIGNSHEET_CMSMANIP + 2 
                            * int.Parse(txtNumAssign.Text)].ToString().Split(' ');
                        txtLevel.Text = subStringLevel[0];

                        subStringPage = subStringLevel[1].Split('-');
                        txtStartPage.Text = (int.Parse(subStringPage[1]) + 1).ToString();
                        cbxPattern.Text = row[ASSIGNSHEET_PATTERN].ToString();
                        cbxDayOff.Text = row[ASSIGNSHEET_DAYOFF].ToString();
                    }

                }
                catch (EntryPointNotFoundException eEx)
                {
                    MessageBox.Show("Student does not seem to do this subject.");
                }
                catch (ArgumentOutOfRangeException aEx)
                {
                    MessageBox.Show("This student does not exist in records.");
                }
                catch (NullReferenceException nEx)
                {
                    MessageBox.Show("Please make sure you have inputted the right barcode and subject " 
                        + "for the student you wish to assign work to.");
                }

                updateData();
                txtBarcode.Clear();

            }
        }







        
        //************************************************************************
        // Definition of function writeData()                                    *
        // This function takes in information from the dataGrid and user fields  *
        // and updates the Assignment Record spreadsheet accordingly.            *
        //************************************************************************
        private void writeData(string[] dateAssign)
        {

            string[] name = lblName.Content.ToString().Split(' ');
            int studentRowNum = assignConnection.getRowNum(DATABASE_SHEET, DATABASE_SHEET_RECORD
    + "!F1:F", name[0], DATABASE_SHEET_RECORD + "!H1:H", name[1]);

            string barcode = assignConnection.get(DATABASE_SHEET
                , DATABASE_SHEET_RECORD + "!D" + studentRowNum.ToString())[0][0].ToString();
            string barcodes = ASSIGNMENT_SHEET_RECORD + "!D1:D";
            string subjectColumn = ASSIGNMENT_SHEET_RECORD + "!E1:E";
            int rowNum = assignConnection.getRowNum(ASSIGNMENT_SHEET
                , barcodes, barcode, subjectColumn, Subject);
            string studentInfoCells = ASSIGNMENT_SHEET_RECORD + "!B" + rowNum.ToString()
                + ":E" + rowNum.ToString();
            string assignmentInfoCells = ASSIGNMENT_SHEET_RECORD + "!I" + rowNum.ToString()
                + ":AAA" + rowNum.ToString();
            var studentInfo = new List<object>();
            var assignmentInfo = new List<object>();

            studentInfo.Add(name[1]);
            studentInfo.Add(name[0]);
            studentInfo.Add(barcode);
            studentInfo.Add(Subject);

            assignConnection.update(studentInfo, ASSIGNMENT_SHEET, studentInfoCells);

            assignmentInfo.Add(DateTime.Now.ToString("MM/dd"));
            assignmentInfo.Add(txtNumAssign.Text);
            assignmentInfo.Add(Pattern);
            assignmentInfo.Add(DayOff);
            assignmentInfo.Add(dateAssign[0]);

            for(int i = 1; i <= 2 * int.Parse(txtNumAssign.Text); i+=2)
            {
                assignmentInfo.Add(dateAssign[i]);
                assignmentInfo.Add(dateAssign[i + 1]);
            }

            assignConnection.update(assignmentInfo, ASSIGNMENT_SHEET, assignmentInfoCells);
        }







        //**********************************************************************************
        // Definition of function updateDate()                                             *
        // If all user fields are filled, then data is extracted and populated accordingly *
        // inside the dataGrid.                                                            *
        //**********************************************************************************
        private void updateData()
        {
            try
            {
                dt.Clear();

                if (txtNumAssign.Text != "" && txtStartDate.Text.Length == 5 && Subject != ""
                    && txtStartPage.Text != "" && txtNumAssign.Text != "" && Pattern != "")
                {
                    KumonLevel level = new KumonLevel(cbxSubject.Text, txtLevel.Text);
                    DateTime assignDate = DateTime.Now;
                    int sheet = int.Parse(txtStartPage.Text);
                    int nextSheet = 0;
                    bool flag = false;


                    for (int i = 0; i < int.Parse(txtNumAssign.Text); i++)
                    {
                        DataRow myRow = dt.NewRow();
                        TimeSpan increment = new TimeSpan(i, 0, 0, 0);
                        assignDate = (DateTime.Parse(txtStartDate.Text + "/" + DateTime.Now.Year.ToString()) 
                            + increment);

                        if (flag || assignDate.DayOfWeek.ToString() == DayOff)
                        {
                            assignDate += new TimeSpan(1, 0, 0, 0);
                            flag = true;
                        }

                        myRow["#"] = i + 1;
                        myRow["Assigned"] = assignDate.ToString("MM/dd");

                        if (sheet > 200)
                        {
                            sheet = 1;
                            level++;
                        }

                        myRow["Level"] = level.Level;

                        nextSheet = calculateNextSheet(Pattern, sheet);

                        myRow["Sheet#"] = sheet.ToString() + "-" + (nextSheet - 1).ToString();

                        sheet = nextSheet;
                        dgdFormat.ItemsSource = dt.DefaultView;

                        dt.Rows.Add(myRow);


                    }

                    lblDateRange.Content = txtStartDate.Text + "-" + assignDate.ToString("MM/dd");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }







        //*****************************************************************************
        // Definition of function getDateAssign()                                     *
        // This function extracts information about assignments assigned for each     *
        // day and formats it in a string[] that is returned. This array is then used *
        // to append to the spreadsheet.                                              *
        //*****************************************************************************
        private string[] getDateAssign()
        {
            string[] dateAssign = new string[76];
            DataTable table = new DataTable();
            table.Columns.Add("Assigned");
            table.Columns.Add("Sheet#");
            int j = 0;

            for (int i = 0; i < dgdFormat.Items.Count; i++)
            {
                DataRow dr = table.NewRow();

                dr["Assigned"] = (dgdFormat.Items[i] as DataRowView)[DATAGRID_ASSIGNDATE];
                dr["Sheet#"] = (dgdFormat.Items[i] as DataRowView)[DATAGRID_LEVEL].ToString()
                    + " " + (dgdFormat.Items[i] as DataRowView)[DATAGRID_PAGES].ToString();

                table.Rows.Add(dr);
            }

            if (chxCMSManip.IsChecked.Value)
                dateAssign[j] = "CMS";
            else
                dateAssign[j] = "";

            j = 1;
            foreach (DataRowView row in table.DefaultView)
            {
                dateAssign[j] = row["Assigned"].ToString();
                dateAssign[j + 1] = row["Sheet#"].ToString();

                j += 2;
            }

            

            return dateAssign;
        }







        //*******************************************************************************
        // Definition of function print()                                               *
        // Page size is set to a quarter page (408x528) and a ticket containing a doc   *
        // with values extracted from the DataGrid is created and sent to the deafult   *
        // printer to be printed.                                                       *
        //*******************************************************************************
        private void print()
        {
            
            PrintDialog Printdlg = new PrintDialog();
            Printdlg.PrintTicket.PageMediaSize = new PageMediaSize(408, 528);

            double width = 408;
            double height = 528;

            // Create a FlowDocument dynamically.
            FlowDocument doc = CreateFlowDocument(width, height);
            doc.Name = "FlowDoc";

            // Create IDocumentPaginatorSource from FlowDocument
            IDocumentPaginatorSource idpSource = doc;

            // Call PrintDocument method to send document to printer
            Printdlg.PrintDocument(idpSource.DocumentPaginator, "Hello WPF Printing.");



        }






        //********************************************************************************
        // Definition of function CreateFlowDocument()                                   *
        // A FlowDocument is created to the specified width and heigh of the page. It    *
        // contains all the information from the datagrid, reformatted for printing.     *
        //********************************************************************************
        private FlowDocument CreateFlowDocument(double pageWidth, double pageHeight)
        {
            // Create a FlowDocument
            FlowDocument doc = new FlowDocument();
            doc.PageWidth = pageWidth;
            doc.PageHeight = pageHeight;
            doc.ColumnWidth = pageWidth;

            // Create a Section
            Section sec = new Section();

            // Create first Paragraph
            Paragraph p1 = new Paragraph();
            Paragraph p2 = new Paragraph();
            // Create and add a new Bold, Italic and Underline
            Bold bld = new Bold();
            Bold bld2 = new Bold();
            bld.Inlines.Add(new Run(lblName.Content.ToString()));
            bld2.Inlines.Add(new Run(lblSubjectBig.Content.ToString() + "\t" + lblDateRange.Content.ToString()));


            // Add Bold, Italic, Underline to Paragraph
            p1.FontFamily = new FontFamily("Lucida Sans");
            p1.FontSize = 25.0;
            p1.TextAlignment = TextAlignment.Center;
            p1.Inlines.Add(bld);


            p2.FontFamily = new FontFamily("Lucida Sans");
            p2.FontSize = 15.0;
            p2.TextAlignment = TextAlignment.Center;
            p2.Inlines.Add(bld2);

            // Add Paragraph to Section
            sec.Blocks.Add(p1);
            sec.Blocks.Add(p2);

            // Add Section to FlowDocument
            doc.Blocks.Add(sec);

            var table = new Table();
            var rowGroup = new TableRowGroup();
            table.RowGroups.Add(rowGroup);
            var header = new TableRow();
            rowGroup.Rows.Add(header);

            int i = 0;
            foreach (DataColumn column in dt.Columns)
            {
                
                string[] headers = new string[] { "#", "Assigned", "Completed", "Level", "Pages" };
                int[] columnWidths = new int[] { 30, 75, 100, 90, 100 };
                var tableColumn = new TableColumn();
                //configure width and such
                tableColumn.Width = new GridLength(columnWidths[i]);
                table.Columns.Add(tableColumn);
                var cell = new TableCell(new Paragraph(new Run(headers[i])));
                cell.FontFamily = new FontFamily("Lucida Sans");
                cell.FontSize = 12.0;
                cell.FontWeight = FontWeights.DemiBold;
                header.Cells.Add(cell);
                i++;
            }

            foreach (DataRow row in dt.Rows)
            {
                var tableRow = new TableRow();
                rowGroup.Rows.Add(tableRow);

                foreach (DataColumn column in dt.Columns)
                {
                    var value = row[column].ToString();//mayby some formatting is in order
                    var cell = new TableCell(new Paragraph(new Run(value)));
                    cell.LineHeight = pageHeight / 25;
                    cell.FontFamily = new FontFamily("Lucida Sans");
                    cell.FontSize = 12.0;
                    tableRow.Cells.Add(cell);
                }
            }

            doc.Blocks.Add(table);

            return doc;
        }







        //*************************************************************************
        // Definition of event handler txtLevel_TextChanged()                     *
        // Data is updated in the view's grid.                                    *
        //*************************************************************************
        private void txtLevel_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateData();
        }







        //*************************************************************************
        // Definition of event handler txtStartPage_TextChanged()                 *
        // Data is updated in the view's grid.                                    *
        //*************************************************************************
        private void txtStartPage_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateData();
        }







        //**************************************************************************
        // Definition of function calculateNextSheet()                             *
        // A currentSheet is taken in as arguments, and the subsequent             *
        // starting page is returned based on the specified pattern                *
        //**************************************************************************
        private int calculateNextSheet(string pattern, int currentSheet)
        {
            int nextSheet = currentSheet;

            switch (pattern)
            {
                case "5-5":
                    nextSheet = currentSheet + 5;
                    break;
                case "4-3-3":
                    switch (currentSheet % 10)
                    {
                        case 1:
                            nextSheet = currentSheet + 4;
                            break;
                        default:
                            nextSheet = currentSheet + 3;
                            break;
                    }
                    break;
                case "3-2":
                    switch (currentSheet % 5)
                    {
                        case 1:
                            nextSheet = currentSheet + 3;
                            break;
                        default:
                            nextSheet = currentSheet + 2;
                            break;
                    }
                    break;
                case "10-10":
                    nextSheet = currentSheet + 10;
                    break;
                case "20-20":
                    nextSheet = currentSheet + 20;
                    break;
                case "2-2":
                    nextSheet = currentSheet + 2;
                    break;
                default:
                    break;
            }

            return nextSheet;
        }







        //***************************************************************************
        // Definition of function getNewPattern()                                   *
        // A starting page and ending page are accepted as arguments. The function  *
        // parses those values and infers which pattern they beong to. This         *
        // calculated pattern is returned.                                          *
        //***************************************************************************
        private string getNewPattern(int startPage, int endPage)
        {

            switch(endPage - startPage + 1)
            {
                case 10:
                    return "10-10";
                case 20:
                    return "20-20";
                case 5:
                    return "5-5";
                case 4:
                    return "4-3-3";
                case 3:
                    if (startPage % 10 == 1 || startPage % 10 == 6)
                        return "3-2";
                    else
                        return "4-3-3";
                case 2:
                    if (startPage % 10 == 4 || startPage % 10 == 9)
                        return "3-2";
                    else
                        return "2-2";
                default:
                    break;
            }

            return Pattern;

        }








        public string DayOff
        {
            get { return dayOff; }
            set
            {
                if (value == null)
                    dayOff = "";
                else if (value.Substring(38) != "None")
                    dayOff = value.Substring(38);
                else
                    dayOff = "";

                updateData();
            }
        }

        public string Pattern
        {
            get { return pattern; }
            set
            {
                if (value == null)
                    dayOff = "";
                else if (value.Substring(38) != "None")
                    pattern = value.Substring(38);
                else
                    pattern = "";

                updateData();
            }
        }

        public string Subject
        {
            get { return subject; }
            set
            {
                if (value.Substring(38) != "None")
                    subject = value.Substring(38);
                else
                    subject = "";
                readAndPopulate();

                lblSubjectBig.Content = Subject;
            }
        }
        
        public string Barcode
        {
            get {
                if (txtBarcode.Text.Length == 5)
                    return txtBarcode.Text;
                else if (txtBarcode.Text.Length == 1)
                    return "A000" + txtBarcode.Text.ToString();
                else if (txtBarcode.Text.Length == 2)
                    return "A00" + txtBarcode.Text.ToString();
                else if (txtBarcode.Text.Length == 3)
                    return "A0" + txtBarcode.Text.ToString();
                else if (txtBarcode.Text.Length == 4)
                    return "A" + txtBarcode.Text.ToString();
                else
                    return null;
                    
            }
            set { txtBarcode.Text = value; }
        }
    }
}
