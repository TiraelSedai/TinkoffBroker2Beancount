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
        public ICollection<Transaction> Read(string filePath)
        {
            var result = new List<Transaction>();
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            do
            {
                if (!reader.Read())
                    return result;

                // Find transactions section.
                var firstCell = reader.GetString(0);
                if (firstCell?.StartsWith("1.1") == true)
                {
                    // Move to first transaciton row.
                    reader.Read();
                    reader.Read();
                    break;
                }
            }
            while (true);

            do
            {
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

                var date = DateTime.Parse(reader.GetString(4));
                var type = reader.GetString(19) switch { "Покупка" => TransactionType.Buy, "Продажа" => TransactionType.Sell, _ => TransactionType.Unknown };
                var name = reader.GetString(21);
                var ticker = reader.GetString(27);
                var price = decimal.Parse(reader.GetString(32), CultureInfo.GetCultureInfo("RU-ru"));
                var amount = int.Parse(reader.GetString(37));
                var accumulatedCoupon = decimal.Parse(reader.GetString(47), ru);
                var comission = decimal.Parse(reader.GetString(54), ru);
                var comissionCur = reader.GetString(57);
                var priceCur = reader.GetString(35);
                
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


                if (!reader.Read())
                    return result;
            }
            while (true);

        }
    }
}
