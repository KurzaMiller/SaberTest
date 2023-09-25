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

        //Метод для считывания и обработки получаемых с указанной таблицы данных
        public static void ReadEntries()
        {
            var range = $"{sourceSheet}!A1:A242";
            var request = sheetService.Spreadsheets.Values.Get(SpreadsheetID, range);

            var responce = request.Execute();
            var values = responce.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    Console.WriteLine(row[0]);
                }
            }
            else Console.WriteLine("тут пусто бля");
        }

        private static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            ConnectToGoogle();
            ReadEntries();

            Console.ReadKey();
        }
    }
}