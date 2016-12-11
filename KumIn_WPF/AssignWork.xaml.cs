// FIX RANGES FOR ASSIGNMENT RECORD -  ADDED NEW COLUMN
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
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
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API KumIn";
        DataTable dt = new DataTable();
        private string dayOff = "";
        private string pattern = "";
        private string subject = "";


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

        private void dgdFormat_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int rowIndex = ((DataGrid)sender).ItemContainerGenerator.IndexFromContainer(e.Row);

            if (e.Column.SortMemberPath.Equals("Level"))
            {
                // enter this level++ for subsequent rows
                KumonLevel nextLevel = new KumonLevel(Subject, ((TextBox)e.EditingElement).Text);

                for (int i = rowIndex; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["Level"] = nextLevel.Level;
                }
            }
            else if (e.Column.SortMemberPath.Equals("Sheet#"))
            {
                int currentSheet = int.Parse(txtStartPage.Text);

                for (int i = 0; i < rowIndex; i++)
                {
                    currentSheet = calculateNextSheet(Pattern, currentSheet);
                }

                // if C, enter corrections and start next row
                if (((TextBox)e.EditingElement).Text == "C")
                {
                    dt.Rows[rowIndex]["Sheet#"] = "Corrections Only";
                    rowIndex++;
                }

                // if end page - start page != expected pattern ==> change patern
                string[] pages = ((TextBox)e.EditingElement).Text.Split('-');
                int startPage = int.Parse(pages[0]);
                int endPage = int.Parse(pages[1]);

                string newPattern = getNewPattern(startPage, endPage);

                for (int i = rowIndex; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["Sheet#"] = currentSheet.ToString() + "-" 
                        + calculateNextSheet(newPattern, currentSheet);

                }


                // start next sheet same pattern next rows
            }
        }

        private void btnPrintRecord_Click(object sender, RoutedEventArgs e)
        {
            writeData(getDateAssign());
            print();
        }

        private void txtBarcode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                readAndPopulate();
        }

        private void txtNumAssign_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateData();
        }

        private void txtStartDate_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateData();
        }

        private void readAndPopulate()
        {
            if (txtBarcode.Text != "" && Subject != "")
            {
                try
                {
                    
                    IList<IList<Object>> values = getSpreadsheetInfo("1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s", "Test!C1:D");

                    bool found = true;
                    string range = "";
                    if (values != null && values.Count > 0)
                    {
                        int rowNum = 1;

                        foreach (var row in values)
                        {
                            
                            if (row[0].ToString() == txtBarcode.Text && row[1].ToString() == Subject)
                            {
                                range = "Test!A" + rowNum.ToString() + ":" + "AAA" + rowNum.ToString();
                                found = true;
                                break;
                            }
                            else if (row[0].ToString() == txtBarcode.Text)
                            {
                                range = "Test!A" + rowNum.ToString() + ":" + "AAA" + rowNum.ToString();
                                found = false;
                            }
                            else
                                rowNum++;
                        }
                    }

                    values = getSpreadsheetInfo("1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s", range);


                    foreach (var row in values)
                    {
                        lblName.Content = row[1].ToString() + " " + row[0].ToString();
                        if (!found)
                            throw new EntryPointNotFoundException();
                        cbxSubject.Text = row[3].ToString();
                        txtNumAssign.Text = row[7].ToString();

                        int lastDateIndex = row.Count - 2;
                        DateTime lastDay = new DateTime(DateTime.Now.Year, int.Parse(string.Concat(row[lastDateIndex].ToString()[0]
                                    , row[lastDateIndex].ToString()[1])), int.Parse(string.Concat(row[lastDateIndex].ToString()[3]
                                    , row[lastDateIndex].ToString()[4]))) + new TimeSpan(1, 0, 0, 0);
                        txtStartDate.Text = (lastDay).ToString("MM/dd");
                        string[] subStringLevel = row[9 + 2 * int.Parse(txtNumAssign.Text)].ToString().Split(' ');
                        txtLevel.Text = subStringLevel[0];
                        string[] subStringPage = subStringLevel[1].Split('-');
                        txtStartPage.Text = (int.Parse(subStringPage[1]) + 1).ToString();
                        cbxPattern.Text = row[8].ToString();
                        cbxDayOff.Text = row[9].ToString();
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

                updateData();

                txtBarcode.Clear();
                //cbxSubject.Text = "None";

            }
        }

        private void clearData()
        {
            dt.Clear();

            lblName.Content = "";
            lblSubjectBig.Content = "";
            lblDateRange.Content = "";
            txtBarcode.Clear();
            cbxSubject.Text = "None";
            txtStartDate.Clear();
            txtNumAssign.Clear();
            txtLevel.Clear();
            txtStartPage.Clear();
            cbxPattern.Text = "None";
            cbxDayOff.Text = "None";

            txtBarcode.Focus();
        }

        private IList<IList<Object>> getSpreadsheetInfo(string spreadsheetId, string range)
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


        private void writeData(string[] dateAssign)
        {
            int rowNum = 1;
            string[] name = lblName.Content.ToString().Split(' ');
            string barcode = name[0].ToUpper() + "-" + string.Concat(name[1][0], name[1][1]).ToUpper();

            IList<IList<Object>> values = getSpreadsheetInfo("1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s", "Test!C1:D");
            string range = "";
            string spreadsheetId = "";
            if (values != null && values.Count > 0)
            {                

                foreach (var row in values)
                {

                    if (row[0].ToString() == barcode && row[1].ToString() == Subject)
                    {
                        range = "Test!A" + rowNum.ToString() + ":" + "AAA" + rowNum.ToString();
                        break;
                    }
                    else
                        rowNum++;
                }
            }

            spreadsheetId = "1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s";
            range = "Test!A" + rowNum.ToString() + ":D" + rowNum.ToString();
            ValueRange valueRange = new ValueRange();

            var oblist = new List<object>();
            var oblist2 = new List<object>();
            oblist.Add(name[1]);
            oblist.Add(name[0]);
            oblist.Add(barcode);
            oblist.Add(Subject);

            updateSpreadsheetInfo(oblist, spreadsheetId, range);

            range = "Test!G" + rowNum.ToString() + ":AAA" + rowNum.ToString();
            oblist2.Add(DateTime.Now.ToString("MM/dd"));
            oblist2.Add(txtNumAssign.Text);
            oblist2.Add(Pattern);
            oblist2.Add(DayOff);

            for(int i = 0; i < 2 * int.Parse(txtNumAssign.Text); i+=2)
            {
                oblist2.Add(dateAssign[i]);
                oblist2.Add(dateAssign[i + 1]);
            }

            updateSpreadsheetInfo(oblist2, spreadsheetId, range);
        }


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
                        assignDate = (new DateTime(DateTime.Now.Year, int.Parse(string.Concat(txtStartDate.Text[0]
                            , txtStartDate.Text[1])), int.Parse(string.Concat(txtStartDate.Text[3]
                            , txtStartDate.Text[4]))) + increment);

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

        private string[] getDateAssign()
        {
            string[] dateAssign = new string[76];
            DataTable table = new DataTable();
            table.Columns.Add("Assigned");
            table.Columns.Add("Sheet#");

            for (int i = 0; i < dgdFormat.Items.Count; i++)
            {
                DataRow dr = table.NewRow();

                dr["Assigned"] = (dgdFormat.Items[i] as DataRowView)[1];
                dr["Sheet#"] = (dgdFormat.Items[i] as DataRowView)[3].ToString()
                    + " " + (dgdFormat.Items[i] as DataRowView)[4].ToString();

                table.Rows.Add(dr);
            }

            int j = 0;
            foreach (DataRowView row in table.DefaultView)
            {
                dateAssign[j] = row["Assigned"].ToString();
                dateAssign[j + 1] = row["Sheet#"].ToString();

                j += 2;
            }

            return dateAssign;
        }

        private void print()
        {
            



            System.Windows.Controls.PrintDialog Printdlg = new System.Windows.Controls.PrintDialog();
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

        private void txtLevel_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateData();
        }

        private void txtStartPage_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateData();
        }

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
    }
}
