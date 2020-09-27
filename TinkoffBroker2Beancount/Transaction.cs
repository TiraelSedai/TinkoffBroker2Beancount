using System;

namespace TinkoffBroker2Beancount
{
    public struct Transaction
    {
        public DateTime Date { get; set; }
        public TransactionType Type { get; set; }
        public string Ticker { get; set; }
        public decimal Price { get; set; }
        public int Amount { get; set; }
        public decimal AccumulatedCoupon { get; set; }
        public decimal Comission { get; set; }
        public string Currency { get; set; }
        public string Name { get; set; }
        public string Mode { get; set; }
    }

    public struct MonetaryTransaction
    {
        public DateTime Date { get; set; }
        public string Operation { get; set; }
        public decimal Amount { get; set; }
        public string Currency { get; set; }
    }

    public enum TransactionType
    {
        Unknown,
        Buy,
        Sell
    }
}
