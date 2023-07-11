
Param([string]$symbol, [string[]]$expirations, $dte, $as_of)

$api_key = Get-Content C:\Users\dharm\Dropbox\api-keys\marketdata-app

class Option
{
    [string]$optionSymbol
    [string]$underlying      
    [string]$expiration      
    [string]$side            
    [decimal]$strike          
    [string]$firstTraded     
    [int]$dte             
    [string]$updated         
    [decimal]$bid             
    [int]$bidSize         
    [decimal]$mid             
    [decimal]$ask             
    [int]$askSize         
    [decimal]$last            
    [int]$openInterest    
    [int]$volume          
    [string]$inTheMoney      
    [decimal]$intrinsicValue  
    [decimal]$extrinsicValue  
    [decimal]$underlyingPrice 
    [decimal]$iv              
    [decimal]$delta           
    [decimal]$gamma           
    [decimal]$theta           
    [decimal]$vega            
    [decimal]$rho                 
}
# ----------------------------------------------------------------------
class Chain
{
    [Option[]]$calls
    [Option[]]$puts
}
# ----------------------------------------------------------------------
# function get-expirations ($symbol)
# {
#     Write-Host "Downloading expirations for $symbol" -ForegroundColor Yellow

#     $result = Invoke-RestMethod ('https://api.marketdata.app/v1/options/expirations/{0}?token={1}' -f $symbol, $api_key)

#     $result.expirations
# }

function get-expirations ($symbol, $as_of)
{
    Write-Host "Downloading expirations for $symbol" -ForegroundColor Yellow

    if ($as_of -eq $null)
    {
        $result = Invoke-RestMethod ('https://api.marketdata.app/v1/options/expirations/{0}?token={1}' -f $symbol, $api_key)
    }
    else
    {
        $result = Invoke-RestMethod ('https://api.marketdata.app/v1/options/expirations/{0}?date={1}&token={2}' -f $symbol, $as_of, $api_key)
    }

    $result.expirations
}

function get-expirations-dte ($symbol, $dte)
{
    $expirations = get-expirations $symbol

    foreach ($expiration in $expirations)
    {
        if (([math]::Ceiling(((Get-Date $expiration) - (Get-Date)).TotalDays) -le 0) -and (((Get-Date $expiration) - (Get-Date)).TotalHours -lt -16))
        {
            continue
        }

        if ($dte -ge [math]::Ceiling(((Get-Date $expiration) - (Get-Date)).TotalDays))
        {
            $expiration
        }
    }
}

# ----------------------------------------------------------------------
# function get-options-chain ($symbol, $expiration)
# {
#     Write-Host "Downloading options chain for $symbol $expiration" -ForegroundColor Yellow


#     $result = Invoke-RestMethod ('https://api.marketdata.app/v1/options/chain/{0}/?expiration={1}&dateformat=timestamp&token={2}' -f $symbol, $expiration, $api_key)
    
#     $options = foreach ($i in 0..($result.optionSymbol.Count - 1))
#     {
#         $obj = [Option]::new()
    
#         $obj.optionSymbol = $result.optionSymbol[$i]
#         $obj.underlying      = $result.underlying[$i]
#         $obj.expiration      = $result.expiration[$i]
#         $obj.side            = $result.side[$i]
#         $obj.strike          = $result.strike[$i]
#         $obj.firstTraded     = $result.firstTraded[$i]
#         $obj.dte             = $result.dte[$i]
#         $obj.updated         = $result.updated[$i]
#         $obj.bid             = $result.bid[$i]
#         $obj.bidSize         = $result.bidSize[$i]
#         $obj.mid             = $result.mid[$i]
#         $obj.ask             = $result.ask[$i]
#         $obj.askSize         = $result.askSize[$i]
#         $obj.last            = $result.last[$i]
#         $obj.openInterest    = $result.openInterest[$i]
#         $obj.volume          = $result.volume[$i]
#         $obj.inTheMoney      = $result.inTheMoney[$i]
#         $obj.intrinsicValue  = $result.intrinsicValue[$i]
#         $obj.extrinsicValue  = $result.extrinsicValue[$i]
#         $obj.underlyingPrice = $result.underlyingPrice[$i]
#         $obj.iv              = $result.iv[$i]
#         $obj.delta           = $result.delta[$i]
#         $obj.gamma           = $result.gamma[$i]
#         $obj.theta           = $result.theta[$i]
#         $obj.vega            = $result.vega[$i]
#         $obj.rho             = $result.rho[$i]
    
#         $obj
#     }

#     $chain = [Chain]::new()

#     $chain.calls = $options | ? side -EQ call
#     $chain.puts  = $options | ? side -EQ put
    
#     $chain
# }

function get-options-chain ($symbol, $expiration, $as_of)
{
    Write-Host "Downloading options chain for $symbol $expiration" -ForegroundColor Yellow

    if ($as_of -eq $null)
    {
        $result = Invoke-RestMethod ('https://api.marketdata.app/v1/options/chain/{0}/?expiration={1}&dateformat=timestamp&token={2}' -f $symbol, $expiration, $api_key)
    }
    else
    {
        $result = Invoke-RestMethod ('https://api.marketdata.app/v1/options/chain/{0}/?expiration={1}&dateformat=timestamp&date={2}&token={3}' -f $symbol, $expiration, $as_of, $api_key)
    }
    
    $options = foreach ($i in 0..($result.optionSymbol.Count - 1))
    {
        $obj = [Option]::new()
    
        $obj.optionSymbol = $result.optionSymbol[$i]
        $obj.underlying      = $result.underlying[$i]
        $obj.expiration      = $result.expiration[$i]
        $obj.side            = $result.side[$i]
        $obj.strike          = $result.strike[$i]
        $obj.firstTraded     = $result.firstTraded[$i]
        $obj.dte             = $result.dte[$i]
        $obj.updated         = $result.updated[$i]
        $obj.bid             = $result.bid[$i]
        $obj.bidSize         = $result.bidSize[$i]
        $obj.mid             = $result.mid[$i]
        $obj.ask             = $result.ask[$i]
        $obj.askSize         = $result.askSize[$i]
        $obj.last            = $result.last[$i]
        $obj.openInterest    = $result.openInterest[$i]
        $obj.volume          = $result.volume[$i]
        $obj.inTheMoney      = $result.inTheMoney[$i]
        $obj.intrinsicValue  = $result.intrinsicValue[$i]
        $obj.extrinsicValue  = $result.extrinsicValue[$i]
        $obj.underlyingPrice = $result.underlyingPrice[$i]
        $obj.iv              = $result.iv[$i]
        $obj.delta           = $result.delta[$i]
        $obj.gamma           = $result.gamma[$i]
        $obj.theta           = $result.theta[$i]
        $obj.vega            = $result.vega[$i]
        $obj.rho             = $result.rho[$i]
    
        $obj
    }

    $chain = [Chain]::new()

    $chain.calls = $options | ? side -EQ call
    $chain.puts  = $options | ? side -EQ put
    
    $chain
}
# ----------------------------------------------------------------------           


function round-to-strike($strikes, $val)
{
    $strikes | Select-Object @{ Label = 'strike'; Expression = { $_ } }, @{ Label = 'dist'; Expression = { [math]::Abs($_ - $val) } } | Sort-Object dist | Select-Object -First 1 | % strike
}

# round-to-strike $strikes_union 123

# $symbol = 'COIN'
# $dte = $null
# $expirations = $null

function chart-open-interest ($symbol, $expirations, $dte, $as_of)
{
    if ($dte -ne $null)
    {
        Write-Host "calling get-expirations-dte $symbol $dte"

        $expirations = get-expirations-dte $symbol $dte
    }

    if ($expirations -eq $null)
    {
        $expirations = get-expirations $symbol $as_of
    }



    # Write-Host "expirations: $expirations"

    Write-Host 'Retrieving options chain data' -ForegroundColor Yellow
    
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date $as_of
    }

    $chains


    $strikes_union = $chains | % { $chain = $_; $chain.calls | % strike } | Sort-Object -Unique

    $colors = @(
        "#4E79A7"
        "#F28E2B"
        "#E15759"
        "#76B7B2"
        "#59A14F"
        "#EDC948"
        "#B07AA1"
        "#FF9DA7"
        "#9C755F"
        "#BAB0AC"
    )

    Write-Host 'Building calls datasets' -ForegroundColor Yellow
            
    $i = 0    

    $datasets_calls = foreach ($chain in $chains)
    {
        $data = $strikes_union | % {
            $strike = $_

            $call = $chain.calls | ? strike -EQ $strike

            if ($call -eq $null) { 0 } else { $call.openInterest }
        }

        $date = $chain.calls[0].expiration.Substring(0, 10)

        $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)

        @{ label = "C $date ${dte}d"; data = $data; backgroundColor = $colors[$i++ % $colors.Count] }
    }

    Write-Host 'Building puts datasets' -ForegroundColor Yellow

    $i = 0

    $datasets_puts = foreach ($chain in $chains)
    {
        $data = $strikes_union | % {
            $strike = $_

            # $option = $chain.puts | ? strike -EQ $strike

            $option = $chain.puts | ? strike -EQ $strike | Select-Object -First 1

            if ($option -eq $null) { 0 } else { -$option.openInterest }
        }

        $date = $chain.puts[0].expiration.Substring(0, 10)

        $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)

        @{ label = "P $date ${dte}d"; data = $data; backgroundColor = $colors[$i++ % $colors.Count] }
    }
    # ----------------------------------------------------------------------
    $total_calls_oi = $chains | % { $chain = $_; $chain.calls | Measure-Object openInterest -Sum | % Sum } | Measure-Object -Sum | % Sum
    $total_puts_oi  = $chains | % { $chain = $_; $chain.puts  | Measure-Object openInterest -Sum | % Sum } | Measure-Object -Sum | % Sum    
    # ----------------------------------------------------------------------
    Write-Host 'Creating chart' -ForegroundColor Yellow
    # ----------------------------------------------------------------------
    $json = @{
        chart = @{
            type = 'bar'
            data = @{
                                
                labels = $strikes_union
                                
                datasets = $datasets_calls + $datasets_puts
            }
            options = @{
                title = @{ 
                    display = $true
                    text = @(
                        ('{0} Open Interest as of {1}' -f $symbol, $(if ($as_of) { $as_of } else { (Get-Date -Format 'yyyy-MM-dd') })),
                        ("total calls OI: {0}" -f $total_calls_oi.ToString('N0')),
                        ("total puts OI: {0}" -f $total_puts_oi.ToString('N0')),
                        'data source: MarketData.app'
                    ) 
                }

                scales = @{ 
                    xAxes = @( @{ stacked = $true } ) 
                    yAxes = @( @{ stacked = $true } ) 
                }
                plugins = @{ datalabels = @{ display = $false } }
                legend = @{ position = 'left' }


                annotation = @{

                    annotations = @(
    
                        @{
                            # type = 'line'; mode = 'vertical'; value = $chains[0].calls[0].underlyingPrice; scaleID = 'X1'; borderColor = 'red'; borderWidth = 1

                            type = 'line'; mode = 'vertical'; value = (round-to-strike $strikes_union $chains[0].calls[0].underlyingPrice); scaleID = 'X1'; borderColor = 'red'; borderWidth = 1

                            # type = 'line'; mode = 'vertical'; value = [math]::Round($chains[0].calls[0].underlyingPrice) ; scaleID = 'X1'; borderColor = 'red'; borderWidth = 1

                            # type = 'line'; mode = 'vertical'; value = 250 ; scaleID = 'X1'; borderColor = 'red'; borderWidth = 1
                            
                            # type = 'line'; mode = 'vertical'; value = 250; scaleID = 'X1'; borderColor = 'red'; borderWidth = 1

                            # type = 'line'; mode = 'vertical'; value = 256.89; scaleID = 'X1'; borderColor = 'red'; borderWidth = 2

                            label = @{
                                # enabled = $true
                                # content = 'Fed Funds Lower'
                                # position = 'end'
                            }
                        }
                    )
                }                
            }
        }
    } | ConvertTo-Json -Depth 100
    
    $result_quickchart = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    $id = ([System.Uri] $result_quickchart.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)    
    # ----------------------------------------------------------------------           
}

chart-open-interest $symbol $expirations $dte $as_of

# ----------------------------------------------------------------------           
exit
# ----------------------------------------------------------------------           

# $chains_spy  = .\chart-open-interest-marketdata.ps1 SPY  '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'
$chains_spy = .\chart-open-interest-marketdata.ps1 SPY
# $chains_spy = .\chart-open-interest-marketdata.ps1 SPY -dte 90

$chains_nvda = .\chart-open-interest-marketdata.ps1 NVDA '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'

$chains_qqq  = .\chart-open-interest-marketdata.ps1 QQQ  '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'
$chains_qqq  = .\chart-open-interest-marketdata.ps1 QQQ -dte 90

$chains_tsla = .\chart-open-interest-marketdata.ps1 TSLA '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'
$chains_tlt  = .\chart-open-interest-marketdata.ps1 TLT  '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'
$chains_gld  = .\chart-open-interest-marketdata.ps1 GLD  '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'
$chains_uso  = .\chart-open-interest-marketdata.ps1 USO  '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'
# $chains_ung  = .\chart-open-interest-marketdata.ps1 UNG  '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18'
$chains_ung = .\chart-open-interest-marketdata.ps1 UNG
# $chains_kre  = .\chart-open-interest-marketdata.ps1 KRE  '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'
$chains_kre  = .\chart-open-interest-marketdata.ps1 KRE
$chains_coin = .\chart-open-interest-marketdata.ps1 COIN 


$chains_nke = .\chart-open-interest-marketdata.ps1 NKE
$chains_tgt = .\chart-open-interest-marketdata.ps1 TGT

$chains_envx = .\chart-open-interest-marketdata.ps1 ENVX
$chains_trup = .\chart-open-interest-marketdata.ps1 TRUP

$chains_trup = .\chart-open-interest-marketdata.ps1 CVNA

# ----------------------------------------------------------------------           
# as-of
# ----------------------------------------------------------------------           



$result_expirations = get-expirations SPY '2022-08-12'


$symbol = 'SPY'
# $as_of = '2022-08-16'
# $as_of = '2022-08-12'
$as_of = '2022-03-29'
$api_key

$result = Invoke-RestMethod ('https://api.marketdata.app/v1/options/expirations/{0}?date={1}&token={2}' -f $symbol, $as_of, $api_key)

# SPY high 2022-03-08
# SPY high 2022-03-29
# SPY low  2022-06-17
# SPY high 2022-08-16

get-expirations SPY '2022-01-04'
get-expirations SPY '2022-03-08'
get-expirations SPY '2022-03-29'
get-expirations SPY '2022-10-13'
get-expirations SPY '2022-12-01'
get-expirations SPY '2022-12-28'
get-expirations SPY '2023-02-02'
get-expirations SPY '2023-03-13'












$chains_spy = .\chart-open-interest-marketdata.ps1 -symbol SPY -as_of '2022-01-04' # high
$chains_spy = chart-open-interest -symbol SPY -as_of '2022-03-08' # low
$chains_spy = chart-open-interest -symbol SPY -as_of '2022-03-29' # high
$chains_spy = chart-open-interest -symbol SPY -as_of '2022-06-17' # low
$chains_spy = chart-open-interest -symbol SPY -as_of '2022-08-12' # high
$chains_spy = chart-open-interest -symbol SPY -as_of '2022-10-13' # low
$chains_spy = chart-open-interest -symbol SPY -as_of '2022-12-01' # high
$chains_spy = chart-open-interest -symbol SPY -as_of '2022-12-28' # low
$chains_spy = chart-open-interest -symbol SPY -as_of '2023-02-02' # high
$chains_spy = chart-open-interest -symbol SPY -as_of '2023-03-13' # low
$chains_spy = .\chart-open-interest-marketdata.ps1 -symbol SPY


$chains_spy = chart-open-interest -symbol SPY -expirations $result_expirations -as_of '2022-08-12' # high


$chain = $chains_spy[0]

foreach ($chain in $chains_spy)
{
    
} 


$chains_spy | % { $chain = $_; $chain.calls | Measure-Object openInterest -Sum | % Sum } | Measure-Object -Sum | % Sum



