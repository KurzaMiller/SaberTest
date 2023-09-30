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
using System.Security.Cryptography.X509Certificates;

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

        static readonly int secInDay = 86400;

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

        public class AnswerStrucure
        { 
            public string field { get; set; }
            public float answer { get; set; }
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
        public static void UpdateEntry(List<List<object>> values, string sheetName, string rangeFrom, string rangeTo)
        {
            //var rowCount = values.ToList<object>().Count;

            var range = $"{sheetName}!{rangeFrom}:{rangeTo}";
            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>>();

            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    var row_obj = new List<object>();
                    row_obj = row;
                    //Console.WriteLine(row[0]);
                    valueRange.Values.Add(row_obj);                    
                    //Console.WriteLine(valueRange.Values.Count + "\n");
                }
            }
                
            else
                throw new Exception("!!! MISSING VALUES !!!");

            var updateRequest = sheetService.Spreadsheets.Values.Update(valueRange, SpreadsheetID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            var updateResponse = updateRequest.Execute();
        }

        public static void AvgRatePriority(List<IList<object>> values)
        {
            var listOfRows = new List<AnswerStrucure>()
            {
                new AnswerStrucure(){ answer = 0, field = "Blocker" },
                new AnswerStrucure(){ answer = 0, field = "Critical" },
                new AnswerStrucure(){ answer = 0, field = "Major" },
                new AnswerStrucure(){ answer = 0, field = "Minor" },
                new AnswerStrucure(){ answer = 0, field = "Trivial" }
            };

            int bl_count = 0;
            int cr_count = 0;
            int maj_count = 0;
            int min_count = 0;
            int tr_count = 0;

            foreach (var row in values)
            {
                switch (row[(int)column.Priority])
                {
                    case "Blocker":
                        listOfRows[0].answer += Convert.ToInt32(row[(int)column.OriginalEstimate]);
                        bl_count++;
                        //Console.WriteLine(bl_count);
                        break;
                    case "Critical":
                        listOfRows[1].answer += Convert.ToInt32(row[(int)column.OriginalEstimate]);
                        cr_count++;
                        break;
                    case "Major":
                        listOfRows[2].answer += Convert.ToInt32(row[(int)column.OriginalEstimate]);
                        maj_count++;
                        break;
                    case "Minor":
                        listOfRows[3].answer += Convert.ToInt32(row[(int)column.OriginalEstimate]);
                        min_count++;
                        break;
                    case "Trivial":
                        listOfRows[4].answer += Convert.ToInt32(row[(int)column.OriginalEstimate]);
                        tr_count++;
                        break;
                }
            }

            List<List<object>> listToTable = new List<List<object>>();
            foreach(var row in listOfRows)
            {
                switch (row.field)
                {
                    case "Blocker":
                        //Console.WriteLine(row.answer);
                        //Console.WriteLine((row.answer/bl_count)/secInDay);
                        row.answer = (row.answer / bl_count) / secInDay;
                        break;
                    case "Critical":
                        row.answer = (row.answer /cr_count) / secInDay;
                        break;
                    case "Major":
                        row.answer = (row.answer / maj_count) / secInDay;
                        break;
                    case "Minor":
                        row.answer = (row.answer / min_count) / secInDay;
                        break;
                    case "Trivial":
                        row.answer = (row.answer / tr_count) / secInDay;
                        break;
                }
                //Console.WriteLine(row.answer);
                listToTable.Add(new List<object>() { row.field, row.answer });
            }

            UpdateEntry(listToTable, reportSheet, "A2", "B7");
        }

        public static void StatsOnClosedTasks(List<IList<object>> values)
        {
            var listOfRows = new List<AnswerStrucure>()
            {
                new AnswerStrucure(){ answer = 0, field = "Blocker" },
                new AnswerStrucure(){ answer = 0, field = "Critical" },
                new AnswerStrucure(){ answer = 0, field = "Major" }
            };

            foreach (var row in values) 
            {
                if (row[(int)column.Status].ToString() == "Closed")
                {
                    switch (row[(int)column.Priority])
                    {
                        case "Blocker":
                            listOfRows[0].answer++;
                            break;
                        case "Critical":
                            listOfRows[1].answer++;
                            break;
                        case "Major":
                            listOfRows[2].answer++;
                            break;
                        default: break;
                    }
                }                    
            }

            List<List<object>> listToTable = new List<List<object>>();
            foreach (var row in listOfRows)
            {
                //Console.WriteLine(row.answer);
                listToTable.Add(new List<object>() { row.field, row.answer });
            }

            UpdateEntry(listToTable, reportSheet, "A9", "B11");
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

            AvgRatePriority(values);
            StatsOnClosedTasks(values);

            Console.ReadKey();
        }
    }
}