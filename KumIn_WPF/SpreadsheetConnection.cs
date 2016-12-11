/*
Establishes a model for the program to interact with google spreadsheets.
The object is constructed by initially reading "client_secret.json" in
order to extract login credentials into the google system. Upon success,
a spreadsheet service is created through which values may be obtained,
updated, or appended.
*/

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;

namespace KumIn_WPF
{
    class SpreadsheetConnection
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API KumIn";
        SheetsService service;


        
        public SpreadsheetConnection()
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
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }







        //***********************************************************************
        // Definition of get()                                                  *
        // Gets values in spreadsheet specified by spreadsheetID and range. The *
        // API getRequest is instantiated with the ID of the desired sheet and  *
        // range of cells. The values in a 2D List are returned.                *
        //***********************************************************************
        public IList<IList<Object>> get(string spreadsheetID, string range)
        {

            SpreadsheetsResource.ValuesResource.GetRequest request =
                        service.Spreadsheets.Values.Get(spreadsheetID, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            return values;
        }






        //************************************************************************
        // Definition of update()                                                *
        // Takes in a list of objects and forms a 2D list of objects which forms *
        // the ValueRange object. The google API UpdateRequest object is then    *
        // instantiated with the ValueRange, and desired end spreadsheet ID      *
        // number and range of cells. The request is executed.                   *
        //************************************************************************
        public void update(IList<Object> oblist, string spreadsheetID, string range)
        {
            List<IList<Object>> values = new List<IList<object>> { oblist };

            ValueRange valueRange = new ValueRange();
            valueRange.Values = values;

            SpreadsheetsResource.ValuesResource.UpdateRequest request =
                        service.Spreadsheets.Values.Update(valueRange, spreadsheetID, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            request.Execute();
        }








        //***********************************************************************
        // Definition of append()                                               *
        // Takes a list of objects as input and forms a 2D list of objects which*
        // forms a ValueRange object. The google API AppendRequest object is    * 
        // instantiated with the ValueRange, a Google spreadsheet ID number and *
        // a defined range of cells. The append request is executed.            *
        //***********************************************************************
        public void append(IList<Object> oblist, string spreadsheetID, string range)
        {
            List<IList<Object>> values = new List<IList<object>> { oblist };

            ValueRange valueRange = new ValueRange();
            valueRange.Values = values;

            SpreadsheetsResource.ValuesResource.AppendRequest request =
                        service.Spreadsheets.Values.Append(valueRange, spreadsheetID, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            request.Execute();
        }







        //*************************************************************************
        // Definition of isValuePresent()                                         *
        // Accepts a range and corresponding value to check, and returns a bool   *
        // on whether or not that value was found in that set of cells in the     *
        // spreadsheet. Function returns true if value is found in some cell in   *
        // the specified range. Else, function returns false.                     *
        //*************************************************************************
        public bool isValuePresent(string spreadsheetID, string columnRange, string value)
        {
            IList<IList<Object>> cellColumn = get(spreadsheetID, columnRange);

            foreach (var cell in cellColumn)
            {
                if (cell[0].ToString() == value)
                {
                    return true;
                }
            }

            return false;

        }





        //************************************************************************
        // Definition of getrowNum()                                             *
        // Obtains a column of cells specified by spreadsheetID and columnRange. *
        // Then, value is searched inside the retrieved list of cells and the row*
        // number of the match is returned. If not found, -1 is returned.        *
        //************************************************************************
        public int getRowNum(string spreadsheetID, string columnRange, string value)
        {
            IList<IList<Object>> column = get(spreadsheetID, columnRange);
            int rowNum = 1;
            foreach (var cell in column)
            {
                if (cell[0].ToString() == value)
                    return rowNum;
                else
                    rowNum++;
            }

            return -1;
        }







        //*************************************************************************
        // Definition of overload of getRowNum()                                  *
        // Same logic as getRowNum, but now crosschecks two different values found*
        // in 2 different colunns of the same row.                                *
        //*************************************************************************
        public int getRowNum(string spreadsheetID, string columnRange1, string value1
            , string columnRange2, string value2)
        {
            IList<IList<Object>> column1 = get(spreadsheetID, columnRange1);
            IList<IList<Object>> column2 = get(spreadsheetID, columnRange2);

            int rowNum = 1;

            for (int i = 0; i < column1.Count; i++)
            {
                if (column1[i].ToString() == value1 && column2[i].ToString() == value2)
                    return rowNum;
                else
                    rowNum++;
            }

            return -1;
        }


    }
}
