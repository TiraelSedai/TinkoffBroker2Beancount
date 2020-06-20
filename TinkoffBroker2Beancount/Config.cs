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
            ComissionAccount = config?["ComissionAccount"] ?? "Expenses:TaxesFees";
            InvestmentsIncomeAccount = config?["InvestmentsIncomeAccount"] ?? "Income:Investments";
        }
        public static string BrokerAccount { get; private set; }
        public static string CouponsAccount { get; private set; }
        public static string ComissionAccount { get; private set; }
        public static string InvestmentsIncomeAccount { get; private set; }
    }
}
