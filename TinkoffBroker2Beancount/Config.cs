using Newtonsoft.Json;
using System.IO;

namespace TinkoffBroker2Beancount
{
    public static class Config
    {
        static Config()
        {
            var configTxt = File.ReadAllText("settings.json");
            dynamic? config = JsonConvert.DeserializeObject(configTxt);
            BrokerAccount = config?["BrokerAccount"] ?? "Assets:Broker:Tinkoff";
            CouponsAccount = config?["CouponsAccount"] ?? "Income:Coupons";
            DividendsAccount = config?["DividendsAccount"] ?? "Income:Dividends";
            ComissionAccount = config?["ComissionAccount"] ?? "Expenses:TaxesFees";
            TaxAccount = config?["TaxAccount"] ?? "Expenses:TaxesFees";
            CapitalGainsAccount = config?["CapitalGainsAccount"] ?? "Income:Investments";
            BankAccount = config?["BankAccount"] ?? "Assets: Tinkoff:Bank";
        }
        public static string BrokerAccount { get; }
        public static string CouponsAccount { get; }
        public static string DividendsAccount { get; }
        public static string TaxAccount { get; }
        public static string ComissionAccount { get; }
        public static string CapitalGainsAccount { get; }
        public static string BankAccount { get; }
    }
}
