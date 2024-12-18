using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace laba_5
{
    
        class ApplicationHelper
        {
            private readonly string logFileName;
            private readonly string excelFilePath;
            private List<Account> accounts;
            private List<Currency> currencies;
            private List<CurrencyRate> exchangeRates;
            private List<Transaction> transactions;

            public ApplicationHelper(string logFileName, string excelFilePath)
            {
                this.logFileName = logFileName;
                this.excelFilePath = excelFilePath;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                accounts = new List<Account>();
                currencies = new List<Currency>();
                exchangeRates = new List<CurrencyRate>();
                transactions = new List<Transaction>();
            }

            public void Start()
            {
                using (StreamWriter logWriter = new StreamWriter(logFileName, true))
                {
                    try
                    {
                        logWriter.WriteLine($"Сессия началась: {DateTime.Now}");

                        ReadExcelData();

                        while (true)
                        {
                            ShowMenu();
                            if (!int.TryParse(Console.ReadLine(), out int choice))
                            {
                                Console.WriteLine("Некорректный ввод. Повторите попытку.");
                            }

                            switch (choice)
                            {
                                case 1:
                                    ViewData(logWriter);
                                    break;
                                case 2:
                                    DeleteElement(logWriter);
                                    break;
                                case 3:
                                    UpdateElement(logWriter);
                                    break;
                                case 4:
                                    AddElement(logWriter);
                                    break;
                                case 5:
                                    ExecuteQueries(logWriter);
                                    break;
                                case 0:
                                    logWriter.WriteLine($"Сессия завершилась: {DateTime.Now}");
                                    Console.WriteLine("Выход из программы. До свидания!");
                                    return;
                                default:
                                    Console.WriteLine("Некорректный выбор. Повторите попытку.");
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logWriter.WriteLine($"Ошибка: {ex.Message}");
                        Console.WriteLine($"Произошла ошибка: {ex.Message}");
                    }
                }
            }

            private void ReadExcelData()
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    // Чтение таблицы "Счета"
                    var accountsWorksheet = package.Workbook.Worksheets["Счета"];
                    if (accountsWorksheet == null) throw new Exception("Не удалось найти лист 'Счета' в файле Excel.");

                    for (int row = 2; row <= accountsWorksheet.Dimension.End.Row; row++)
                    {
                        accounts.Add(new Account
                        {
                            ID = int.Parse(accountsWorksheet.Cells[row, 1].Text),
                            Name = accountsWorksheet.Cells[row, 2].Text,
                            OpenDate = DateTime.Parse(accountsWorksheet.Cells[row, 3].Text)
                        });
                    }

                    // Чтение таблицы "Валюты"
                    var currenciesWorksheet = package.Workbook.Worksheets["Валюты"];
                    if (currenciesWorksheet == null) throw new Exception("Не удалось найти лист 'Валюты' в файле Excel.");

                    for (int row = 2; row <= currenciesWorksheet.Dimension.End.Row; row++)
                    {
                        currencies.Add(new Currency
                        {
                            ID = int.Parse(currenciesWorksheet.Cells[row, 1].Text),
                            LetterCode = currenciesWorksheet.Cells[row, 2].Text,
                            Name = currenciesWorksheet.Cells[row, 3].Text
                        });
                    }

                    // Чтение таблицы "Курсы валют"
                    var exchangeRatesWorksheet = package.Workbook.Worksheets["Курсы валют"];
                    if (exchangeRatesWorksheet == null) throw new Exception("Не удалось найти лист 'Курс валют' в файле Excel.");

                    for (int row = 2; row <= exchangeRatesWorksheet.Dimension.End.Row; row++)
                    {
                        exchangeRates.Add(new CurrencyRate
                        {
                            ID = int.Parse(exchangeRatesWorksheet.Cells[row, 1].Text),
                            CurrencyID = int.Parse(exchangeRatesWorksheet.Cells[row, 2].Text),
                            Date = DateTime.Parse(exchangeRatesWorksheet.Cells[row, 3].Text),
                            Rate = decimal.Parse(exchangeRatesWorksheet.Cells[row, 4].Text)
                        });
                    }

                    // Чтение таблицы "Поступления"
                    var transactionsWorksheet = package.Workbook.Worksheets["Поступления"];
                    if (transactionsWorksheet == null) throw new Exception("Не удалось найти лист 'Начисления' в файле Excel.");

                    for (int row = 2; row <= transactionsWorksheet.Dimension.End.Row; row++)
                    {
                        transactions.Add(new Transaction
                        {
                            ID = int.Parse(transactionsWorksheet.Cells[row, 1].Text),
                            AccountID = int.Parse(transactionsWorksheet.Cells[row, 2].Text),
                            CurrencyID = int.Parse(transactionsWorksheet.Cells[row, 3].Text),
                            Date = DateTime.Parse(transactionsWorksheet.Cells[row, 4].Text),
                            Amount = decimal.Parse(transactionsWorksheet.Cells[row, 5].Text)
                        });
                    }
                }
            }

            private void ViewData(StreamWriter logWriter)
            {
                Table();
                if (!int.TryParse(Console.ReadLine(), out int choice))
                {
                    Console.WriteLine("Некорректный ввод. Повторите попытку.");
                }

                switch (choice)
                {
                    case 1:
                        Console.WriteLine("Счета:");
                        accounts.ForEach(Console.WriteLine);
                        Console.WriteLine("Выведена таблица 'Счета'.");
                        break;
                    case 2:
                        Console.WriteLine("Валюты:");
                        currencies.ForEach(Console.WriteLine);
                        Console.WriteLine("Выведена таблица 'Валюты'.");
                        break;
                    case 3:
                        Console.WriteLine("Курсы валют:");
                        exchangeRates.ForEach(Console.WriteLine);
                        Console.WriteLine("Выведена таблица 'Курсы валют'.");
                        break;
                    case 4:
                        Console.WriteLine("Поступления:");
                        transactions.ForEach(Console.WriteLine);
                        Console.WriteLine("Выведена таблица 'Поступления'.");
                        break;
                    default:
                        Console.WriteLine("Некорректный выбор. Повторите попытку.");
                        break;

                }
            }

            private void DeleteElement(StreamWriter logWriter)
            {
                Table();
                if (!int.TryParse(Console.ReadLine(), out int choice))
                {
                    Console.WriteLine("Некорректный ввод. Повторите попытку.");
                }
                else
                {
                    Console.WriteLine("Введите Валюта ID для удаления:");
                    if (int.TryParse(Console.ReadLine(), out int id))
                    {
                        switch (choice)
                        {
                            case 1:
                                var account = accounts.FirstOrDefault(a => a.ID == id);
                                if (account != null)
                                {
                                    accounts.Remove(account);
                                    logWriter.WriteLine($"Удален элемент: {account}");
                                    Console.WriteLine("Элемент успешно удален.");
                                }
                                else
                                {
                                    Console.WriteLine("Элемент с таким ID не найден.");
                                }
                                break;
                            case 2:
                                var currency = currencies.FirstOrDefault(a => a.ID == id);
                                if (currency != null)
                                {
                                    currencies.Remove(currency);
                                    logWriter.WriteLine($"Удален элемент: {currency}");
                                    Console.WriteLine("Элемент успешно удален.");
                                }
                                else
                                {
                                    Console.WriteLine("Элемент с таким ID не найден.");
                                }
                                break;
                            case 3:
                                Console.WriteLine("Введите id валюты для удаления:");
                                if (int.TryParse(Console.ReadLine(), out int currencyid))
                                {

                                    var removedCurrencyRate = exchangeRates.FirstOrDefault(a => (a.CurrencyID == currencyid && a.ID == id));

                                    if (removedCurrencyRate != null)
                                    {
                                        exchangeRates.Remove(removedCurrencyRate);
                                        Console.WriteLine("Элемент успешно удален.");
                                        logWriter.WriteLine($"Удален элемент: {removedCurrencyRate}");
                                    }
                                    else
                                    {
                                        Console.WriteLine("Элемент с таким ID или ID валюты не найден.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Некорректный ввод id валюты. ");
                                }
                                break;
                            case 4:
                                var transaction = transactions.FirstOrDefault(a => a.ID == id);
                                if (transaction != null)
                                {
                                    transactions.Remove(transaction);
                                    logWriter.WriteLine($"Удален элемент: {transaction}");
                                    Console.WriteLine("Элемент успешно удален.");
                                }
                                else
                                {
                                    Console.WriteLine("Элемент с таким ID не найден.");
                                }
                                break;
                            default:
                                Console.WriteLine("Некорректный выбор. Повторите попытку.");
                                break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Некорректный ввод ID.");
                    }
                }
            }

            private void UpdateElement(StreamWriter logWriter)
            {
                Table();
                if (!int.TryParse(Console.ReadLine(), out int choice))
                {
                    Console.WriteLine("Некорректный ввод. Повторите попытку.");
                }
                else
                {
                    Console.WriteLine("Введите ID: ");
                    if (int.TryParse(Console.ReadLine(), out int id))
                    {
                        switch (choice)
                        {
                            case 1:
                                var account = accounts.FirstOrDefault(a => a.ID == id);
                                if (account != null)
                                {
                                    Console.WriteLine("Введите новое имя:");
                                    string name = Console.ReadLine();
                                    if (!string.IsNullOrWhiteSpace(name))
                                    {
                                        account.Name = name;
                                        Console.WriteLine("Введите новую дату открытия:");
                                        if (DateTime.TryParse(Console.ReadLine(), out DateTime newDate))
                                        {
                                            account.OpenDate = newDate;
                                            logWriter.WriteLine($"Обновлен элемент: {account}");
                                            Console.WriteLine("Элемент успешно обновлен.");
                                        }
                                        else
                                        {
                                            Console.WriteLine("Некорректный ввод даты.");
                                        }
                                    }
                                    {
                                        Console.WriteLine("Вы ввели пустую сторку или строку из пробелов.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Элемент с таким ID не найден.");
                                }
                                break;
                            case 2:
                                var currency = currencies.FirstOrDefault(a => a.ID == id);
                                if (currency != null)
                                {
                                    Console.WriteLine("Введите новое имя:");
                                    string name = Console.ReadLine();
                                    if (!string.IsNullOrWhiteSpace(name))
                                    {
                                        currency.Name = name;
                                        Console.WriteLine("Введите новую буквенный код:");
                                        string currencyLetterCode = Console.ReadLine();
                                        if (currencyLetterCode.All(char.IsLetter) == true && !string.IsNullOrWhiteSpace(currencyLetterCode))
                                        {
                                            currency.LetterCode = currencyLetterCode;
                                            logWriter.WriteLine($"Обновлен элемент: {currency}");
                                            Console.WriteLine("Элемент успешно обновлен.");
                                        }
                                        else
                                        {
                                            Console.WriteLine("Некорректный ввод буквенного кода.");
                                        }
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Элемент с таким ID не найден.");
                                }
                                break;
                            case 3:
                                Console.WriteLine("Введите id валюты для изменения данных:");
                                if (int.TryParse(Console.ReadLine(), out int currencyid))
                                {

                                    var CurrencyRate = exchangeRates.FirstOrDefault(a => (a.CurrencyID == currencyid && a.ID == id));
                                    if (CurrencyRate != null)
                                    {
                                        Console.WriteLine("Введите курс:");
                                        if (decimal.TryParse(Console.ReadLine(), out decimal newRate) && newRate >= 0)
                                        {
                                            CurrencyRate.Rate = newRate;
                                            Console.WriteLine("Введите новую дату(в диапазоне с 24 по 30 декабря 2021 года):");
                                            DateTime startDate = new DateTime(2021, 12, 24);
                                            DateTime endDate = new DateTime(2021, 12, 30);

                                            if (DateTime.TryParse(Console.ReadLine(), out DateTime newDate) && newDate <= endDate && newDate >= startDate)
                                            {
                                                CurrencyRate.Date = newDate;
                                                logWriter.WriteLine($"Обновлен элемент: {CurrencyRate}");
                                                Console.WriteLine("Элемент успешно обновлен.");
                                            }
                                            else
                                            {
                                                Console.WriteLine("Некорректный ввод даты.");
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("Некорректный ввод курса. ");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Элемент с таким ID не найден.");
                                    }
                                }
                                break;
                            case 4:
                                var transaction = transactions.FirstOrDefault(a => a.ID == id);
                                if (transaction != null)
                                {
                                    Console.WriteLine("Введите поступления:");
                                    if (decimal.TryParse(Console.ReadLine(), out decimal newAmount))
                                    {
                                        transaction.Amount = newAmount;
                                        Console.WriteLine("Введите новую дату(в диапазоне с 24 по 30 декабря включительно 2021 года):");
                                        DateTime startDate = new DateTime(2021, 12, 24);
                                        DateTime endDate = new DateTime(2021, 12, 30);

                                        if (DateTime.TryParse(Console.ReadLine(), out DateTime newDate) && newDate <= endDate && newDate >= startDate)
                                        {
                                            transaction.Date = newDate;
                                            logWriter.WriteLine($"Обновлен элемент: {transaction}");
                                            Console.WriteLine("Элемент успешно обновлен.");
                                        }
                                        else
                                        {
                                            Console.WriteLine("Некорректный ввод даты.");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Некорректный ввод поступления. ");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Элемент с таким ID не найден.");
                                }
                                break;
                            default:
                                Console.WriteLine("Некорректный выбор. Повторите попытку.");
                                break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Некорректный ввод ID.");
                    }
                }
            }

            private void AddElement(StreamWriter logWriter)
            {
                Table();
                if (!int.TryParse(Console.ReadLine(), out int choice))
                {
                    Console.WriteLine("Некорректный ввод. Повторите попытку.");
                }
                else
                {
                    switch (choice)
                    {
                        case 1:
                            Console.WriteLine("Введите уникальный ID нового элемента:");
                            if (int.TryParse(Console.ReadLine(), out int id) && !accounts.Any(a => a.ID == id))
                            {
                                Console.WriteLine("Введите имя нового элемента:");
                                string name = Console.ReadLine();
                                if (!string.IsNullOrWhiteSpace(name))
                                {
                                    Console.WriteLine("Введите дату открытия нового элемента(в диапазоне 2021 года):");
                                    DateTime startDate = new DateTime(2021, 1, 1);
                                    DateTime endDate = new DateTime(2021, 12, 31);

                                    if (DateTime.TryParse(Console.ReadLine(), out DateTime openDate) && openDate <= endDate && openDate >= startDate)
                                    {
                                        var newAccount = new Account { ID = id, Name = name, OpenDate = openDate };
                                        accounts.Add(newAccount);
                                        logWriter.WriteLine($"Добавлен элемент: {newAccount}");
                                        Console.WriteLine("Элемент успешно добавлен.");
                                    }
                                    else
                                    {
                                        Console.WriteLine("Некорректный ввод даты.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Имя не должно быть пустым или состоять из пробелов. ");
                                }
                            }
                            else
                            {
                                Console.WriteLine("Некорректный ввод ID.");
                            }
                            break;
                        case 2:
                            Console.WriteLine("Введите уникальный ID нового элемента:");
                            if (int.TryParse(Console.ReadLine(), out int currid) && !currencies.Any(a => a.ID == currid))
                            {
                                Console.WriteLine("Введите имя нового элемента:");
                                string currname = Console.ReadLine();
                                if (!string.IsNullOrWhiteSpace(currname))
                                {
                                    Console.WriteLine("Введите буквенный код нового элемента:");
                                    string currlettercode = Console.ReadLine();
                                    if (currlettercode.All(char.IsLetter) == true && !string.IsNullOrWhiteSpace(currlettercode))
                                    {
                                        var newCurrency = new Currency { ID = currid, Name = currname, LetterCode = currlettercode };
                                        currencies.Add(newCurrency);
                                        logWriter.WriteLine($"Добавлен элемент: {newCurrency}");
                                        Console.WriteLine("Элемент успешно добавлен.");
                                    }
                                    else
                                    {
                                        Console.WriteLine("Некорректный ввод даты.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Буквенный код должна состоять только из букв. ");
                                }
                            }
                            else
                            {
                                Console.WriteLine("Некорректный ввод ID.");
                            }
                            break;
                        case 3:
                            Console.WriteLine("Введите уникальный ID нового элемента:");
                            if (int.TryParse(Console.ReadLine(), out int rateid) && !exchangeRates.Any(a => a.ID == rateid))
                            {
                                Console.WriteLine("Введите ID валюты(может быть неуникальным, но такой ID должен существовать в таблице 'Валюты') нового элемента:");
                                if (int.TryParse(Console.ReadLine(), out int ratecurrid) && currencies.Any(a => a.ID == ratecurrid))
                                {
                                    Console.WriteLine("Введите дату нового элемента(в диапазоне от 24 по 30 декабря включительно 2021 года):");
                                    DateTime startDate = new DateTime(2021, 12, 24);
                                    DateTime endDate = new DateTime(2021, 12, 30);

                                    if (DateTime.TryParse(Console.ReadLine(), out DateTime rateDate) && rateDate <= endDate && rateDate >= startDate)
                                    {
                                        Console.WriteLine("Введите курс валюты(не менее нуля): ");
                                        if (decimal.TryParse(Console.ReadLine(), out decimal rate) && rate >= 0)
                                        {
                                            var newRate = new CurrencyRate { ID = rateid, CurrencyID = ratecurrid, Date = rateDate, Rate = rate };
                                            exchangeRates.Add(newRate);
                                            logWriter.WriteLine($"Добавлен элемент: {newRate}");
                                            Console.WriteLine("Элемент успешно добавлен.");
                                        }
                                        else
                                        {
                                            Console.WriteLine("Некорректный ввод курса валюты. ");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Некорректный ввод даты.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Имя не должно быть пустым или состоять из пробелов. ");
                                }
                            }
                            else
                            {
                                Console.WriteLine("Некорректный ввод ID.");
                            }
                            break;
                        case 4:
                            Console.WriteLine("Введите уникальный ID нового элемента:");
                            if (int.TryParse(Console.ReadLine(), out int transactionid) && !transactions.Any(a => a.ID == transactionid))
                            {
                                Console.WriteLine("Введите ID валюты(может быть неуникальным, но такой ID должен существовать в таблице 'Валюты') нового элемента:");
                                if (int.TryParse(Console.ReadLine(), out int transactioncurrid) && currencies.Any(a => a.ID == transactioncurrid))
                                {
                                    Console.WriteLine("Введите ID счёта(может быть неуникальным, но такой ID должен существовать в таблице 'Счета') нового элемента:");
                                    if (int.TryParse(Console.ReadLine(), out int transactionaccountid) && accounts.Any(a => a.ID == transactionaccountid))
                                    {
                                        Console.WriteLine("Введите дату нового элемента(в диапазоне от 24 по 30 декабря включительно 2021 года):");
                                        DateTime startDate = new DateTime(2021, 12, 24);
                                        DateTime endDate = new DateTime(2021, 12, 30);

                                        if (DateTime.TryParse(Console.ReadLine(), out DateTime transactionDate) && transactionDate <= endDate && transactionDate >= startDate)
                                        {
                                            Console.WriteLine("Введите поступления: ");
                                            if (decimal.TryParse(Console.ReadLine(), out decimal amount))
                                            {
                                                var newTransaction = new Transaction { ID = transactionid, CurrencyID = transactioncurrid, Date = transactionDate, Amount = amount, AccountID = transactionaccountid };
                                                transactions.Add(newTransaction);
                                                logWriter.WriteLine($"Добавлен элемент: {newTransaction}");
                                                Console.WriteLine("Элемент успешно добавлен.");
                                            }
                                            else
                                            {
                                                Console.WriteLine("Некорректный ввод поступления. ");
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("Некорректный ввод даты.");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Некорректный ввод ID счёта. ");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Некорректный ввод ID валюты. ");
                                }
                            }
                            else
                            {
                                Console.WriteLine("Некорректный ввод ID.");
                            }
                            break;
                        default:
                            Console.WriteLine("Некорректный выбор. Повторите попытку.");
                            break;
                    }

                }
            }

            private void ExecuteQueries(StreamWriter logWriter)
            {
                Console.WriteLine("Запрос 1: Все счета, открытые в феврале 2021 года. ");
                var result1 = accounts.Where(a => (a.OpenDate >= new DateTime(2021, 2, 1) && a.OpenDate <= new DateTime(2021, 2, 28))).ToList();
                result1.ForEach(Console.WriteLine);


                Console.Write("Запрос 2: Название всех валют, чей курс когда-либо был выше 80. ");
                var result2 = exchangeRates
            .Where(rate => rate.Rate > 80)
            .Join(currencies, rate => rate.CurrencyID, currency => currency.ID, (rate, currency) => currency.Name)
            .Distinct().ToList();
                Console.WriteLine("Названия всех валют, чей курс был выше 80:");
                result2.ForEach(Console.Write);

                Console.WriteLine("Запрос 3: найти сумму пополнений среди счетов, открытых в январе 2021 года, и в валюте 'Евро' .");
                var result = transactions
            .Join(
                accounts,
                transaction => transaction.AccountID,
                account => account.ID,
                (transaction, account) => new { Transaction = transaction, Account = account }
            )
            .Where(x => x.Account.OpenDate >= new DateTime(2021, 1, 1)
                         && x.Account.OpenDate <= new DateTime(2021, 1, 31))
            .Join(
                currencies,
                x => x.Transaction.CurrencyID,
                currency => currency.ID,
                (x, currency) => new { x.Transaction, x.Account, Currency = currency }
            )
            .Where(x => x.Currency.Name.Trim() == "Евро")
            .ToList();
                decimal totalSum = result.Sum(x => x.Transaction.Amount);
                Console.WriteLine($"Сумма пополнений для счетов, открытых в январе 2021 года и в валюте 'Евро': {totalSum}");

                Console.WriteLine("Запрос 4: Максимальная сумма операции в валюте 'Доллар США' для счетов, открытых первого числа любого месяца. ");

                var maxTransactionInDollar = transactions
                    .Join(
                        accounts,
                        transaction => transaction.AccountID,
                        account => account.ID,
                        (transaction, account) => new { Transaction = transaction, Account = account }
                    )
                    .Where(x => x.Account.OpenDate.Day == 1)
                    .Join(
                        currencies,
                        x => x.Transaction.CurrencyID,
                        currency => currency.ID,
                        (x, currency) => new { x.Transaction, x.Account, Currency = currency }
                    )
                    .Where(x => x.Currency.Name.Trim() == "Доллар США")
                    .Max(x => x.Transaction.Amount);

                Console.WriteLine($"Максимальная сумма операции: {maxTransactionInDollar}");
                logWriter.WriteLine("Выполнение всех запросов завершено успешно. ");
            }
            public void SaveExcelData(string excelFilePath)
            {
                // Убедимся, что EPPlus использует лицензирование (обязательно для работы с библиотекой)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Создаем новый Excel файл или перезаписываем существующий
                using (var package = new ExcelPackage())
                {
                    // ====== Сохранение таблицы "Счета" ======
                    var accountsWorksheet = package.Workbook.Worksheets.Add("Счета");

                    accountsWorksheet.Cells[1, 1].Value = "ID";
                    accountsWorksheet.Cells[1, 2].Value = "ФИО";
                    accountsWorksheet.Cells[1, 3].Value = "Дата открытия";

                    for (int i = 0; i < accounts.Count; i++)
                    {
                        accountsWorksheet.Cells[i + 2, 1].Value = accounts[i].ID;
                        accountsWorksheet.Cells[i + 2, 2].Value = accounts[i].Name;
                        accountsWorksheet.Cells[i + 2, 3].Value = accounts[i].OpenDate.ToString("dd.MM.yyyy");
                    }

                    // ====== Сохранение таблицы "Валюты" ======
                    var currenciesWorksheet = package.Workbook.Worksheets.Add("Валюты");

                    currenciesWorksheet.Cells[1, 1].Value = "ID";
                    currenciesWorksheet.Cells[1, 2].Value = "Буквенный код";
                    currenciesWorksheet.Cells[1, 3].Value = "Наименование валюты";

                    for (int i = 0; i < currencies.Count; i++)
                    {
                        currenciesWorksheet.Cells[i + 2, 1].Value = currencies[i].ID;
                        currenciesWorksheet.Cells[i + 2, 2].Value = currencies[i].LetterCode;
                        currenciesWorksheet.Cells[i + 2, 3].Value = currencies[i].Name;
                    }

                    // ====== Сохранение таблицы "Курсы валют" ======
                    var exchangeRatesWorksheet = package.Workbook.Worksheets.Add("Курсы валют");

                    // Заголовки для "Курсы валют"
                    exchangeRatesWorksheet.Cells[1, 1].Value = "ID";
                    exchangeRatesWorksheet.Cells[1, 2].Value = "ID валюты";
                    exchangeRatesWorksheet.Cells[1, 3].Value = "Дата";
                    exchangeRatesWorksheet.Cells[1, 4].Value = "Курс";

                    for (int i = 0; i < exchangeRates.Count; i++)
                    {
                        exchangeRatesWorksheet.Cells[i + 2, 1].Value = exchangeRates[i].ID;
                        exchangeRatesWorksheet.Cells[i + 2, 2].Value = exchangeRates[i].CurrencyID;
                        exchangeRatesWorksheet.Cells[i + 2, 3].Value = exchangeRates[i].Date.ToString("dd.MM.yyyy");
                        exchangeRatesWorksheet.Cells[i + 2, 4].Value = exchangeRates[i].Rate;
                    }

                    // ====== Сохранение таблицы "Поступления" ======
                    var transactionsWorksheet = package.Workbook.Worksheets.Add("Поступления");

                    transactionsWorksheet.Cells[1, 1].Value = "ID";
                    transactionsWorksheet.Cells[1, 2].Value = "ID счёта";
                    transactionsWorksheet.Cells[1, 3].Value = "ID валюты";
                    transactionsWorksheet.Cells[1, 4].Value = "Дата";
                    transactionsWorksheet.Cells[1, 5].Value = "Поступление";

                    for (int i = 0; i < transactions.Count; i++)
                    {
                        transactionsWorksheet.Cells[i + 2, 1].Value = transactions[i].ID;
                        transactionsWorksheet.Cells[i + 2, 2].Value = transactions[i].AccountID;
                        transactionsWorksheet.Cells[i + 2, 3].Value = transactions[i].CurrencyID;
                        transactionsWorksheet.Cells[i + 2, 4].Value = transactions[i].Date.ToString("dd.MM.yyyy");
                        transactionsWorksheet.Cells[i + 2, 5].Value = transactions[i].Amount;
                    }

                    // ====== Сохранение файла ======
                    var fileInfo = new FileInfo(excelFilePath);
                    package.SaveAs(fileInfo);

                    Console.WriteLine($"Данные успешно сохранены в файл: {excelFilePath}");
                }
            }

            private void ShowMenu()
            {
                Console.WriteLine("\nВыберите действие:");
                Console.WriteLine("1. Просмотр базы данных");
                Console.WriteLine("2. Удаление элемента");
                Console.WriteLine("3. Корректировка элемента");
                Console.WriteLine("4. Добавление элемента");
                Console.WriteLine("5. Выполнение запросов");
                Console.WriteLine("0. Выход");
            }
            private void Table()
            {
                Console.WriteLine("\nВыберите с какой таблицей работаем:");
                Console.WriteLine("1. Счета");
                Console.WriteLine("2. Валюты");
                Console.WriteLine("3. Курсы Валют");
                Console.WriteLine("4. Поступления");
            }
        }
}
