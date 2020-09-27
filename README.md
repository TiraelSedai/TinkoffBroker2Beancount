# TinkoffBroker2Beancount
Converts Excel reports from Tinkoff Broker to Beancount format

## Usage
- Tweak account names in settings.json if necessary
- Run with .xlsx file as first argument in command line

### Example usage
- TinkoffBroker2Beancount.exe "D:\Finance\Reports\broker_rep-2020-may.xlsx"
- dotnet TinkoffBroker2Beancount.dll C:\Downloads\broker_rep-5.xlsx

### Features

- [x] Buy/Sell, including comission
- [x] Equity
- [x] Bonds
- [x] Tariff comission
- [x] Dividents
- [x] Coupons