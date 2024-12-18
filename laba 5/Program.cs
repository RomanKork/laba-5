using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using OfficeOpenXml;

namespace laba_5
{

    class Program
    {
        static void Main()
        {
            Console.WriteLine("Добро пожаловать в программу управления инвестиционными счетами!");

            Console.WriteLine("Введите имя для файла протокола (по умолчанию: file):");
            string logFileName = "C:\\Users\\User\\Desktop\\" + Console.ReadLine() + ".txt";
            if (string.IsNullOrWhiteSpace(logFileName)) logFileName = "C:\\Users\\User\\Desktop\\file.txt";

            Console.WriteLine("Вы хотите создать новый файл протокола или дописать в существующий? (новый/добавить):");
            bool appendLog = Console.ReadLine()?.Trim().ToLower() == "добавить";

            string excelFilePath = "C:\\Users\\User\\Desktop\\laba 5.xlsx";

            var appHelper = new ApplicationHelper(logFileName, excelFilePath);
            appHelper.Start();
            appHelper.SaveExcelData(excelFilePath);
        }
    }
}

