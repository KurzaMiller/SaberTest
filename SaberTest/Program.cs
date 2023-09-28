using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Threading;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Sheets.v4.Data;
using System.Data;

namespace SaberTest
{
    internal class Program
    {
        //Объявление всех необходимых для работы с Google Sheets переменных
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        static readonly string ApplicationName = "SaberTest";
        static readonly string SpreadsheetID = "1oCZkZKQmEHosTUmHhJGozapzm6s2CIjNVXxK1AhWAQA";

        static readonly string sourceSheet = "Source";
        static readonly string reportSheet = "ProgRepo";

        static readonly string posSourceLeftUp = "A2";
        static readonly string posSourceRightDown = "L242";

        enum column
        { 
            Issuekey = 0,
            IssueID = 1,
            IssueType = 2,
            Status = 3,
            Priority = 4,
            Resolution = 5,
            Assignee = 6,
            Reporter = 7,
            Created = 8,
            Resolved = 9,
            OriginalEstimate = 10,
            TimeSpent = 11
        }

        //Объявление переменной сервиса Google Sheets
        static SheetsService sheetService;

        //Метод подключения к Google Cloud через json файл с секретным ключом
        public static void ConnectToGoogle()
        {
            GoogleCredential userCredential;

            using (var fs = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                userCredential = GoogleCredential.FromStream(fs).CreateScoped(Scopes);
            }

            sheetService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = userCredential,
                ApplicationName = ApplicationName
            });
        }

        //Метод для считывания и обработки получаемых с указанной таблицы данных из указанного диапазона
        public static List<IList<object>> ReadEntries(string sheetName, string rangeFrom, string rangeTo)
        {
            var range = $"{sheetName}!{rangeFrom}:{rangeTo}";
            var request = sheetService.Spreadsheets.Values.Get(SpreadsheetID, range);

            var responce = request.Execute();
            var values = (List<IList<object>>)responce.Values;

            if (values != null && values.Count > 0)
                Console.WriteLine("Values are ok");
            else
                throw new Exception("!!! MISSING VALUES !!!");

            return values;
        }

        //Метод для создания списка записей в указанной таблице с ближайшего свободного места в указаном диапазоне
        public static void CreateEntry(List<List<object>> values, string sheetName, string rangeFrom, string rangeTo)
        {
            var range = $"{sheetName}!{rangeFrom}:{rangeTo}";
            var valueRange = new ValueRange();

            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    var row_obj = new List<object>();
                    row_obj = row;
                    valueRange.Values = new List<IList<object>> { row_obj };
                }
            }
                
            else
                throw new Exception("!!! MISSING VALUES !!!");

            var appendRequest = sheetService.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            var appendResponse = appendRequest.Execute();
        }

        public void AvgRatePriority(List<IList<object>> values)
        {
            var listOfRows = new List<List<object>>()
            { 
                new List<object>(),
                new List<object>(),
                new List<object>(),
                new List<object>(),
                new List<object>()
            };
            int blocker_pr = 0;
            int critical_pr = 0;
            int major_pr = 0;
            int minor_pr = 0;
            int trivial_pr = 0;
            
            foreach(var row in values)
            {
                switch (row[(int)column.Priority])
                {
                    case "Blocker":
                        blocker_pr += (int)row[(int)column.OriginalEstimate];
                        break;
                    case "Critical":
                        critical_pr += (int)row[(int)column.OriginalEstimate];
                        break;
                    case "Major":
                        major_pr += (int)row[(int)column.OriginalEstimate];
                        break;
                    case "Minor":
                        minor_pr += (int)row[(int)column.OriginalEstimate];
                        break;
                    case "Trivial":
                        trivial_pr += (int)row[(int)column.OriginalEstimate];
                        break;
                }
            }

            List<IList<object>> values_list = new List<IList<object>>();

            foreach
        }

        private static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            ConnectToGoogle();

            var values = new List<IList<object>>();
            values = ReadEntries(sourceSheet, posSourceLeftUp, posSourceRightDown);

            //Console.WriteLine(values[0].GetType());

            //foreach (var row in values)
            //{
            //    Console.WriteLine("{0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} | {8} | {9} | {10} | {11} \n",
            //                    row[(int)column.Issuekey], row[(int)column.IssueID], row[(int)column.IssueType],
            //                    row[(int)column.Status], row[(int)column.Priority], row[(int)column.Resolution],
            //                    row[(int)column.Assignee], row[(int)column.Reporter], row[(int)column.Created],
            //                    row[(int)column.Resolved], row[(int)column.OriginalEstimate], row[(int)column.TimeSpent]);
            //}



            Console.ReadKey();
        }
    }
}