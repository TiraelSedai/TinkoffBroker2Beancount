using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace TinkoffBroker2Beancount
{
    class ExcelReader
    {
        public static ICollection<Transaction> ReadTransactions(string filePath, out int linesRead)
        {
            linesRead = 0;
            var result = new List<Transaction>();
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            while (true)
            {
                if (!reader.Read())
                    return result;
                linesRead++;

                // Find transactions section.
                var firstCell = reader.GetString(0);
                if (firstCell?.StartsWith("1.1") == true)
                {
                    // Move to before the first transaciton row.
                    reader.Read();
                    break;
                }
            }

            // Read columns
            var columns = new Dictionary<string, int>();
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
                    "валютацены" => nameof(Transaction.PriceCurrency),
                    "количество" => nameof(Transaction.Amount),
                    "нкд" => nameof(Transaction.AccumulatedCoupon),
                    "комиссияброкера" => nameof(Transaction.Comission),
                    "валютакомиссии" => nameof(Transaction.ComissionCurrency),
                    _ => null
                };
                if (key != null)
                    columns.Add(key, i);
            }

            while (true)
            {
                if (!reader.Read())
                    return result;
                linesRead++;

                var firstCell = reader.GetString(0);
                if (firstCell?.StartsWith("1.2") == true)
                    // End of executed transactions section.
                    return result;

                if (!long.TryParse(firstCell, out var _))
                {
                    // It's ok, some lines are info about pages and whatnot.
                    reader.Read();
                    continue;
                }
                var ru = CultureInfo.GetCultureInfo("RU-ru");

                var date = DateTime.Parse(reader.GetString(columns[nameof(Transaction.Date)]));
                var type = reader.GetString(columns[nameof(Transaction.Type)]) switch { "Покупка" => TransactionType.Buy, "Продажа" => TransactionType.Sell, _ => TransactionType.Unknown };
                var name = reader.GetString(columns[nameof(Transaction.Name)]);
                var ticker = reader.GetString(columns[nameof(Transaction.Ticker)]);
                var price = decimal.Parse(reader.GetString(columns[nameof(Transaction.Price)]), CultureInfo.GetCultureInfo("RU-ru"));
                var priceCur = reader.GetString(columns[nameof(Transaction.PriceCurrency)]);
                var amount = int.Parse(reader.GetString(columns[nameof(Transaction.Amount)]));
                var accumulatedCoupon = decimal.Parse(reader.GetString(columns[nameof(Transaction.AccumulatedCoupon)]), ru);
                var comission = decimal.Parse(reader.GetString(columns[nameof(Transaction.Comission)]), ru);
                var comissionCur = reader.GetString(columns[nameof(Transaction.ComissionCurrency)]);
                
                result.Add(new Transaction
                {
                    AccumulatedCoupon = accumulatedCoupon,
                    Amount = amount,
                    Comission = comission,
                    Date = date,
                    Price = price,
                    Ticker = ticker,
                    Type = type,
                    Name = name,
                    PriceCurrency = priceCur,
                    ComissionCurrency = comissionCur
                });
            }
        }
    }
}
