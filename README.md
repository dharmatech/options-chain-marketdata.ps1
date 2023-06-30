
# chart-open-interest-marketdata.ps1

![image](https://github.com/dharmatech/options-chain-marketdata.ps1/assets/20816/45031217-5591-4b8f-b1c6-4d4b17276adc)

## Example invocations

SPY with all expirations:

    $chains_spy  = .\chart-open-interest-marketdata.ps1 SPY
    
SPY with specific expirations:

    $chains_spy  = .\chart-open-interest-marketdata.ps1 SPY  '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'

SPY with expirations up to 90 days:

    $chains_spy = .\chart-open-interest-marketdata.ps1 SPY -dte 90

## API key

Update the `$api_key` variable accordingly.

## Data source

MarketData.app is being used for the data source:

https://www.marketdata.app/
