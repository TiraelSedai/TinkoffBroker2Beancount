using System;
using System.Globalization;

namespace TinkoffBroker2Beancount
{
    public class BeancountFormatter
    {
        public static void PrintTransaction(Transaction transaction)
        {
            Console.WriteLine($"{transaction.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)} * \"{transaction.Type switch { TransactionType.Buy => "Покупка", TransactionType.Sell => "Продажа", _ => "???" }} {transaction.Amount} {transaction.Name}\"");
            
            switch (transaction.Type)
            {
                case TransactionType.Buy:
                    Console.WriteLine($"\t{Config.BrokerAccount}\t+{transaction.Amount} {transaction.Ticker} {{{transaction.Price:0.00} {transaction.PriceCurrency}}}");
                    Console.WriteLine($"\t{Config.BrokerAccount}");
                    break;
                case TransactionType.Sell:
                    Console.WriteLine($"\t{Config.BrokerAccount}\t-{transaction.Amount} {transaction.Ticker} {{}}");
                    Console.WriteLine($"\t{Config.BrokerAccount} +{(transaction.Amount * transaction.Price):0.00} {transaction.PriceCurrency}");
                    Console.WriteLine($"\t{Config.InvestmentsIncomeAccount}");
                    Console.WriteLine($"\t{Config.BrokerAccount} -{transaction.Comission:0.00} {transaction.ComissionCurrency}");
                    break;
                default:
                    Console.WriteLine($"Unable to pretty-print because type is {transaction.Type}. Amount = {transaction.Amount}, Price = {transaction.Price:0.00} {transaction.PriceCurrency}");
                    break;
            }
            Console.WriteLine($"\t{Config.ComissionAccount}\t{transaction.Comission:0.00} {transaction.ComissionCurrency}");
            if (transaction.AccumulatedCoupon != 0)
                Console.WriteLine($"\t{Config.CouponsAccount}\t{(transaction.Type == TransactionType.Sell ? "-":"+")}{transaction.AccumulatedCoupon} {transaction.PriceCurrency}");
            Console.WriteLine();
        }
    }
}
