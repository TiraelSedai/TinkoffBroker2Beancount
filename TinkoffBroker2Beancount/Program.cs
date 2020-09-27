using System;
using System.IO;
using System.Linq;

namespace TinkoffBroker2Beancount
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!args.Any())
                throw new ArgumentException("Please provide Excel report as first argument");

            var filename = args.First();
            if (!File.Exists(filename))
                throw new ArgumentException($"Unable to find specified report {filename}");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            Worker.ReadTransactions(filename);
        }
    }
}
