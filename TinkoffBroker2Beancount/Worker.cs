using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace TinkoffBroker2Beancount
{
    public class Worker
    {
        public static void ReadTransactions(string filePath)
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            while (true)
            {
                if (!reader.Read())
                    return;

                // Find transactions section.
                var firstCell = reader.GetString(0);
                if (firstCell?.StartsWith("1.1 ") == true)
                {
                    // Move to before the first transaciton row.
                    reader.Read();
                    break;
                }
            }

            // Read columns
            var columns = new Dictionary<string, int>();
            FillDataColumns(reader, columns);

            while (true)
            {
                if (!reader.Read())
                    return;

                var firstCell = reader.GetString(0);

                if (MoveUtilFind(reader, "1.2 "))
                    // End of executed transactions section.
                    break;

                if (!long.TryParse(firstCell, out var _))
                {
                    // It's ok, some lines are info about pages and whatnot.
                    continue;
                }

                ParseAndPrintTransactions(reader, columns);
            }

            var currencyList = new List<string>();
            while (true)
            {
                if (!reader.Read())
                    return;

                // Find money section.
                if (MoveUtilFind(reader, "2. "))
                {
                    // Move to before the first transaciton row.
                    reader.Read();
                    break;
                }
            }

            while (true)
            {
                if (!reader.Read())
                    return;

                var firstCell = reader.GetString(0);

                if (string.IsNullOrWhiteSpace(firstCell))
                    continue;

                if (currencyList.Contains(firstCell))
                {
                    // Move to first currency
                    reader.Read();
                    break;
                }

                currencyList.Add(firstCell);

                // We should never arrive there
                if (MoveUtilFind(reader, "3.1"))
                    // End of executed transactions section.
                    break;
            }
            var currentCurrency = currencyList.FirstOrDefault();

            if (currentCurrency != default)
            {
                while (true)
                {
                    var columnsMoney = new Dictionary<string, int>();
                    FillMoneySectionColumns(reader, columnsMoney);

                    while (true)
                    {
                        if (!reader.Read())
                            return;

                        var firstCell = reader.GetString(0);
                        if (MoveUtilFind(reader, "3.1"))
                            return;

                        if (currencyList.Contains(firstCell))
                        {
                            // Move to next currency
                            currentCurrency = firstCell;
                            reader.Read();
                            break;
                        }

                        var tryDate = reader.GetString(columnsMoney[nameof(MonetaryTransaction.Date)]);
                        if (string.IsNullOrWhiteSpace(tryDate))
                            continue;
                        ParseAndPrintMoneyTransactions(reader, currentCurrency, columnsMoney, tryDate);
                    }
                }
            }

        }

        private static void ParseAndPrintMoneyTransactions(IExcelDataReader reader, string currentCurrency, Dictionary<string, int> columnsMoney, string tryDate)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("ru-RU");
            var date = DateTime.Parse(tryDate);
            var operation = reader.GetString(columnsMoney[nameof(MonetaryTransaction.Operation)]);
            var amountp = decimal.Parse(reader.GetString(columnsMoney[nameof(Transaction.Amount) + "+"]));
            var amountm = decimal.Parse(reader.GetString(columnsMoney[nameof(Transaction.Amount) + "-"]));

            PrintTransaction(new MonetaryTransaction
            {
                Amount = amountp - amountm,
                Date = date,
                Operation = operation,
                Currency = currentCurrency
            });
        }

        private static void FillMoneySectionColumns(IExcelDataReader reader, Dictionary<string, int> columnsMoney)
        {
            for (var i = 0; i < reader.FieldCount; i++)
            {
                var colNoSpace = reader.GetString(i)?.Replace(" ", "")?.Replace("\n", "")?.ToLowerInvariant();
                var key = colNoSpace switch
                {
                    "датаисполнения" => nameof(MonetaryTransaction.Date),
                    "операция" => nameof(MonetaryTransaction.Operation),
                    "суммазачисления" => nameof(MonetaryTransaction.Amount) + "+",
                    "суммасписания" => nameof(MonetaryTransaction.Amount) + "-",
                    _ => null
                };
                if (key != null)
                    columnsMoney.Add(key, i);
            }
        }

        private static void ParseAndPrintTransactions(IExcelDataReader reader, Dictionary<string, int> columns)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("ru-RU");
            var date = DateTime.Parse(reader.GetString(columns[nameof(Transaction.Date)]));

            var typeRaw = reader.GetString(columns[nameof(Transaction.Type)]);
            TransactionType type = TransactionType.Unknown;
            if (typeRaw.Contains("Покупка"))
                type = TransactionType.Buy;
            else if (typeRaw.Contains("Продажа"))
                type = TransactionType.Sell;

            var name = reader.GetString(columns[nameof(Transaction.Name)]);
            var ticker = reader.GetString(columns[nameof(Transaction.Ticker)]);
            var price = decimal.Parse(reader.GetString(columns[nameof(Transaction.Price)]));
            var priceCur = reader.GetString(columns[nameof(Transaction.Currency)]);
            var amount = int.Parse(reader.GetString(columns[nameof(Transaction.Amount)]));
            var accumulatedCoupon = decimal.Parse(reader.GetString(columns[nameof(Transaction.AccumulatedCoupon)]));
            var comission = decimal.Parse(reader.GetString(columns[nameof(Transaction.Comission)]));
            var mode = reader.GetString(columns[nameof(Transaction.Mode)]);

            PrintTransaction(new Transaction
            {
                AccumulatedCoupon = accumulatedCoupon,
                Amount = amount,
                Comission = comission,
                Date = date,
                Price = price,
                Ticker = ticker,
                Type = type,
                Name = name.Replace("\"", ""),
                Currency = priceCur,
                Mode = mode
            });
        }

        private static void FillDataColumns(IExcelDataReader reader, Dictionary<string, int> columns)
        {
            for (var i = 0; i < reader.FieldCount; i++)
            {
                var colNoSpace = reader.GetString(i)?.Replace(" ", "")?.Replace("\n", "")?.ToLowerInvariant();
                var key = colNoSpace switch
                {
                    "датазаключения" => nameof(Transaction.Date),
                    "видсделки" => nameof(Transaction.Type),
                    "сокращенноенаименованиеактива" => nameof(Transaction.Name),
                    "кодактива" => nameof(Transaction.Ticker),
                    "ценазаединицу" => nameof(Transaction.Price),
                    "валютацены" => nameof(Transaction.Currency),
                    "количество" => nameof(Transaction.Amount),
                    "нкд" => nameof(Transaction.AccumulatedCoupon),
                    "комиссияброкера" => nameof(Transaction.Comission),
                    "режимторгов" => nameof(Transaction.Mode),
                    _ => null
                };
                if (key != null)
                    columns.Add(key, i);
            }
        }

        private static bool MoveUtilFind(IExcelDataReader reader, string searchString)
        {
            bool foundBreak = false;
            for (var i = 0; i < reader.FieldCount; i++)
            {
                try
                {
                    var cell = reader.GetString(i);

                    if (cell?.StartsWith(searchString) == true)
                    {
                        foundBreak = true;
                        break;
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }

            return foundBreak;
        }


        public static void PrintTransaction(MonetaryTransaction transaction)
        {
            var oper = transaction.Operation.ToLowerInvariant();
            if (oper.StartsWith("налог"))
            {
                PrintTransaction(transaction, Config.TaxAccount);
                return;
            }
            if (oper.Equals("пополнение счета") || oper.Equals("вывод средств"))
            {
                PrintTransaction(transaction, Config.BankAccount);
                return;
            }
            if (oper.Equals("комиссия по тарифу"))
            {
                PrintTransaction(transaction, Config.ComissionAccount);
                return;
            }
            if (oper.Equals("выплата дивидендов"))
            {
                PrintTransaction(transaction, Config.DividendsAccount);
                return;
            }
            if (oper.Equals("выплата купонов"))
            {
                PrintTransaction(transaction, Config.CouponsAccount);
                return;
            }
        }

        private static void PrintTransaction(MonetaryTransaction transaction, string otherAccount)
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            Console.WriteLine($"{transaction.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)} * \"Брокерский счёт: {transaction.Operation}\"");
            Console.WriteLine($"\t{Config.BrokerAccount}\t{transaction.Amount} {transaction.Currency}");
            Console.WriteLine($"\t{otherAccount}\t{0 - transaction.Amount} {transaction.Currency}");
            Console.WriteLine();
        }

        public static void PrintTransaction(Transaction transaction)
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            if (string.Equals(transaction.Mode, "CNGD", StringComparison.OrdinalIgnoreCase))
            {
                var commodity = string.Concat(transaction.Ticker.Take(3));
                Console.WriteLine($"{transaction.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)} * \"{transaction.Type switch { TransactionType.Buy => "Покупка", TransactionType.Sell => "Продажа", _ => "???" }} {transaction.Amount} {commodity}\"");
                Console.WriteLine($"\t{Config.BrokerAccount}\t{(transaction.Type == TransactionType.Buy ? "+" : "-")}{transaction.Amount} {commodity} @ {transaction.Price} {transaction.Currency}");
                Console.WriteLine($"\t{Config.BrokerAccount}\t{(transaction.Type == TransactionType.Buy ? "-" : "+")}{transaction.Amount * transaction.Price} {transaction.Currency}");
                if (transaction.Comission != 0)
                    Console.WriteLine($"\t{Config.BrokerAccount}\t-{transaction.Comission} {transaction.Currency}");
            }
            else
            {
                Console.WriteLine($"{transaction.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)} * \"{transaction.Type switch { TransactionType.Buy => "Покупка", TransactionType.Sell => "Продажа", _ => "???" }} {transaction.Amount} {transaction.Name}\"");
                switch (transaction.Type)
                {
                    case TransactionType.Buy:
                        Console.WriteLine($"\t{Config.BrokerAccount}\t+{transaction.Amount} {transaction.Ticker} {{{transaction.Price} {transaction.Currency}}}");
                        Console.WriteLine($"\t{Config.BrokerAccount}");
                        break;
                    case TransactionType.Sell:
                        Console.WriteLine($"\t{Config.BrokerAccount}\t-{transaction.Amount} {transaction.Ticker} {{}} @ {transaction.Price} {transaction.Currency}");
                        Console.WriteLine($"\t{Config.BrokerAccount} +{(transaction.Amount * transaction.Price + transaction.AccumulatedCoupon)} {transaction.Currency}");
                        Console.WriteLine($"\t{Config.CapitalGainsAccount}");
                        if (transaction.Comission != 0)
                            Console.WriteLine($"\t{Config.BrokerAccount} -{transaction.Comission} {transaction.Currency}");
                        break;
                    default:
                        Console.WriteLine($"Unable to pretty-print because type is {transaction.Type}. Amount = {transaction.Amount}, Price = {transaction.Price} {transaction.Currency}");
                        break;
                }
                if (transaction.AccumulatedCoupon != 0)
                    Console.WriteLine($"\t{Config.CouponsAccount}\t{(transaction.Type == TransactionType.Sell ? "-" : "+")}{transaction.AccumulatedCoupon} {transaction.Currency}");
            }
            if (transaction.Comission != 0)
                Console.WriteLine($"\t{Config.ComissionAccount}\t{transaction.Comission} {transaction.Currency}");

            Console.WriteLine();
        }
    }
}
