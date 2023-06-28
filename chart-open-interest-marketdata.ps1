
Param([string]$symbol, [string[]]$expirations)

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
function get-options-chain ($symbol, $expiration)
{
    $result = Invoke-RestMethod ('https://api.marketdata.app/v1/options/chain/{0}/?expiration={1}&dateformat=timestamp&token={2}' -f $symbol, $expiration, $api_key)
    
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
function chart-open-interest ($symbol, $expirations)
{
    
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
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

    $i = 0

    $datasets_puts = foreach ($chain in $chains)
    {
        $data = $strikes_union | % {
            $strike = $_

            $option = $chain.puts | ? strike -EQ $strike

            if ($option -eq $null) { 0 } else { -$option.openInterest }
        }

        $date = $chain.puts[0].expiration.Substring(0, 10)

        $dte = [math]::Ceiling(((Get-Date $date) - (Get-Date)).TotalDays)

        @{ label = "P $date ${dte}d"; data = $data; backgroundColor = $colors[$i++ % $colors.Count] }
    }

    # ----------------------------------------------------------------------
    $json = @{
        chart = @{
            type = 'bar'
            data = @{
                                
                labels = $strikes_union
                                
                datasets = $datasets_calls + $datasets_puts
            }
            options = @{
                title = @{ display = $true; text = ('{0} Open Interest' -f $symbol) }
                scales = @{ 
                    xAxes = @( @{ stacked = $true } ) 
                    yAxes = @( @{ stacked = $true } ) 
                }
                plugins = @{ datalabels = @{ display = $false } }
                legend = @{ position = 'left' }
            }
        }
    } | ConvertTo-Json -Depth 100
    
    $result_quickchart = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    $id = ([System.Uri] $result_quickchart.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)    
    # ----------------------------------------------------------------------           
}

chart-open-interest $symbol $expirations
# ----------------------------------------------------------------------           
exit
# ----------------------------------------------------------------------           
$symbol = 'SPY'
$expirations = '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'

$chains_spy = .\chart-open-interest-marketdata.ps1 SPY '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'

$chains = $chains_spy