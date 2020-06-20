using System;
using System.Collections.Generic;
using System.Text;

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
        public string ComissionCurrency { get; set; }
        public string PriceCurrency { get; set; }
        public string Name { get; set; }
    }

    public enum TransactionType
    {
        Unknown,
        Buy,
        Sell
    }
}
