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

            /*
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

            String spreadsheetId = "1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s";
            String range = "Test";
            ValueRange valueRange = new ValueRange();

            var oblist = new List<object>() { "Srinath", "Is", "A", "Legend", "Yay!" };
            valueRange.Values = new List<IList<object>> { oblist };

            SpreadsheetsResource.ValuesResource.AppendRequest append = service.Spreadsheets.Values.Append(
                valueRange, spreadsheetId, range);
            append.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;

            AppendValuesResponse result = append.Execute();
            */
            dt.Columns.Add("#");
            dt.Columns.Add("Assigned");
            dt.Columns.Add("Completed");
            dt.Columns.Add("Level");
            dt.Columns.Add("Sheet#");
        }

        private void btnPrintRecord_Click(object sender, RoutedEventArgs e)
        {
            writeData(updateData());
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

                    string spreadsheetId = "1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s";
                    string range = "Test!C1:D";
                    SpreadsheetsResource.ValuesResource.GetRequest request =
                            service.Spreadsheets.Values.Get(spreadsheetId, range);

                    ValueRange response = request.Execute();
                    IList<IList<Object>> values = response.Values;

                    bool found = true;
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

                    request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    response = request.Execute();
                    values = response.Values;



                    foreach (var row in values)
                    {
                        lblName.Content = row[1].ToString() + " " + row[0].ToString();
                        if (!found)
                            throw new EntryPointNotFoundException();
                        cbxSubject.Text = row[3].ToString();
                        txtNumAssign.Text = row[5].ToString();
                        txtStartDate.Text = (new DateTime(DateTime.Now.Year, int.Parse(string.Concat(row[4].ToString()[0]
                                    , row[4].ToString()[1])), int.Parse(string.Concat(row[4].ToString()[3]
                                    , row[4].ToString()[4]))) + new TimeSpan(int.Parse(row[5].ToString()) + 1, 0, 0
                                    , 0)).ToString("MM/dd");
                        string[] subStringLevel = row[7 + 2 * int.Parse(txtNumAssign.Text)].ToString().Split(' ');
                        txtLevel.Text = subStringLevel[0];
                        string[] subStringPage = subStringLevel[1].Split('-');
                        txtStartPage.Text = (int.Parse(subStringPage[1]) + 1).ToString();
                        cbxPattern.Text = row[6].ToString();
                        cbxDayOff.Text = row[7].ToString();
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


        private void writeData(string[] dateAssign)
        {
            UserCredential credential;
            int rowNum = 1;
            string[] name = lblName.Content.ToString().Split(' ');
            string barcode = name[0].ToUpper() + "-" + string.Concat(name[1][0], name[1][1]).ToUpper();

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

            string spreadsheetId = "1rQvp2rNVHpCyVaOCgnDJQo_5Hzvq6217DfTEs1czm9s";
            string range = "Test!C1:D";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;

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
            range = "Test!A" + rowNum.ToString() + ":AAA" + rowNum.ToString();
            ValueRange valueRange = new ValueRange();

            var oblist = new List<object>();

            oblist.Add(name[1]);
            oblist.Add(name[0]);
            oblist.Add(barcode);
            oblist.Add(Subject);
            oblist.Add(DateTime.Now.ToString("MM/dd"));
            oblist.Add(txtNumAssign.Text);
            oblist.Add(Pattern);
            oblist.Add(DayOff);

            for(int i = 0; i < 2 * int.Parse(txtNumAssign.Text); i+=2)
            {
                oblist.Add(dateAssign[i]);
                oblist.Add(dateAssign[i + 1]);

            }

            valueRange.Values = new List<IList<object>> { oblist };

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(
                valueRange, spreadsheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

            UpdateValuesResponse result = update.Execute();



        }


        private string[] updateData()
        {
            string[] dateAssign = new string[50];
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
                    int j = 0;
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

                        switch (Pattern)
                        {
                            case "5-5":
                                nextSheet = sheet + 5;
                                break;
                            case "4-3-3":
                                switch(sheet % 10)
                                {
                                    case 1:
                                        nextSheet = sheet + 4;
                                        break;
                                    default:
                                        nextSheet = sheet + 3;
                                        break;
                                }
                                break;
                            case "3-2":
                                switch(sheet % 5)
                                {
                                    case 1:
                                        nextSheet = sheet + 3;
                                        break;
                                    default:
                                        nextSheet = sheet + 2;
                                        break;
                                }
                                break;
                            case "10-10":
                                nextSheet = sheet + 10;
                                break;
                            case "20-20":
                                nextSheet = sheet + 20;
                                break;
                            case "2-2":
                                nextSheet = sheet + 2;
                                break;
                            default:
                                break;

                        }
                        


                        myRow["Sheet#"] = sheet.ToString() + "-" + (nextSheet - 1).ToString();

                        dateAssign[j] = assignDate.ToString("MM/dd");
                        dateAssign[j + 1] = level.Level + " " + sheet.ToString() + "-" + (nextSheet - 1).ToString();

                        sheet = nextSheet;
                        dgdFormat.ItemsSource = dt.DefaultView;

                        dt.Rows.Add(myRow);

                        j += 2;
                    }
                    lblDateRange.Content = txtStartDate.Text + "-" + assignDate.ToString("MM/dd");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
