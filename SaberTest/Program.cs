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
        static readonly string reportSheet = "Report";

        static readonly string posSourceLeftUp = "A2";
        static readonly string posSourceRightDown = "L242";

        enum column
        { 
            Issuekey,
            IssueID,
            IssueType,
            Status,
            Priority,
            Resolution,
            Assignee,
            Reporter,
            Created,
            Resolved,
            OriginalEstimate,
            TimeSpent
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
        public static List<object> ReadEntries(string sheetName, string rangeFrom, string rangeTo)
        {
            var range = $"{sheetName}!{rangeFrom}:{rangeTo}";
            var request = sheetService.Spreadsheets.Values.Get(SpreadsheetID, range);

            var responce = request.Execute();
            var values = (List<object>)responce.Values;

            if (values != null && values.Count > 0)
                Console.WriteLine("Values are ok");
            else
                throw new Exception("!!! MISSING VALUES !!!");

            return values;
        }

        //Метод для создания списка записей в указанной таблице с ближайшего свободного места в указаном диапазоне
        public static void CreateEntry(List<object> values, string sheetName, string rangeFrom, string rangeTo)
        {
            var range = $"{sheetName}!{rangeFrom}:{rangeTo}";
            var valueRange = new ValueRange();

            if (values != null && values.Count > 0)
                valueRange.Values = new List<IList<object>> { values };
            else
                throw new Exception("!!! MISSING VALUES !!!");

            var appendRequest = sheetService.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            var appendResponse = appendRequest.Execute();
        }

        private static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            ConnectToGoogle();

            var values = new List<object>();
            values = ReadEntries(sourceSheet, posSourceLeftUp, posSourceRightDown);

            Console.ReadKey();
        }
    }
}