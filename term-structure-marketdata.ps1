
$api_key = Get-Content C:\Users\dharm\Dropbox\api-keys\marketdata-app


$expiration = '2023-07-21'

$result = Invoke-RestMethod ('https://api.marketdata.app/v1/options/chain/{0}/?expiration={1}&dateformat=timestamp&token={2}' -f $symbol, $expiration, $api_key)

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

# $result.iv[1].GetType()

class Chain
{
    [Option[]]$calls
    [Option[]]$puts
}


$result_spx = get-options-chain 'SPX' '2023-06-30'

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

$symbol = 'SPY'
$expirations = '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18'


# function nearest-delta ($options, $delta)
# {

# }


function chart-term-structure ($symbol, $expirations)
{
    $chains = foreach ($date in $expirations)
    {
        get-options-chain $symbol $date
    }

    $chains

    # $chain = $chains[0]

    $atm_calls = foreach ($chain in $chains)
    {
        $atm_strike = $chain.calls | Select-Object strike, @{ Label = 'dist'; Expression = { [math]::Abs($_.strike - $_.underlyingPrice) } } | Sort-Object dist | Select-Object -First 1 | % strike

        $chain.calls | ? strike -EQ $atm_strike | Select-Object -First 1
    }

    $atm_puts = foreach ($chain in $chains)
    {
        $atm_strike = $chain.puts | Select-Object strike, @{ Label = 'dist'; Expression = { [math]::Abs($_.strike - $_.underlyingPrice) } } | Sort-Object dist | Select-Object -First 1 | % strike

        $chain.puts | ? strike -EQ $atm_strike | Select-Object -First 1
    }    
    
    $delta_10_calls = foreach ($chain in $chains)
    {
        $delta_10_strike = $chain.calls | Select-Object strike, delta, @{ Label = 'dist'; Expression = { [math]::Abs($_.delta - 0.10) } } | Sort-Object dist | Select-Object -First 1 | % strike

        $chain.calls | ? strike -EQ $delta_10_strike | Select-Object -First 1
    }

    $delta_10_puts = foreach ($chain in $chains)
    {
        $delta_10_strike = $chain.puts | Select-Object strike, delta, @{ Label = 'dist'; Expression = { [math]::Abs($_.delta - -0.10) } } | Sort-Object dist | Select-Object -First 1 | % strike

        $chain.puts | ? strike -EQ $delta_10_strike | Select-Object -First 1
    }    



    $atm_net_dex = foreach ($i in 0..($atm_calls.Count-1))
    {
        # $atm_calls[$i].delta - $atm_puts[$i].delta

        # $atm_puts[$i].delta - $atm_calls[$i].delta

        # -$atm_puts[$i].delta - $atm_calls[$i].delta

        $atm_calls[$i].delta - -$atm_puts[$i].delta
    }


    # ----------------------------------------------------------------------
    $json = @{
        chart = @{
            type = 'line'
            data = @{
                                
                labels = $chains | % { 
                    $chain = $_; 
                    
                    $chain.calls[0].expiration.Substring(0,10)
                }

                datasets = @(
                    @{ label = 'ATM Call';                           data = $atm_calls | % { $_.iv.ToString('N4') }; fill = $false }
                    @{ label = 'ATM Put' ;                           data = $atm_puts  | % { $_.iv.ToString('N4') }; fill = $false }

                    @{ label = ('OTM Call Delta 10'); data = $delta_10_calls | % { $_.iv.ToString('N4') }; fill = $false }
                    @{ label = ('OTM Put Delta 10');  data = $delta_10_puts  | % { $_.iv.ToString('N4') }; fill = $false }     
                    
                    @{ label = 'NET DEX'; data = $atm_net_dex | % { $_.ToString('N4') }; fill = $false }
                    
                    # @{ label = ('OTM Call {0}' -f $otm_call_strike); data = $otm_calls | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    # @{ label = ('OTM Put {0}'  -f $otm_put_strike ); data = $otm_puts  | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    # @{ label = ('OTM Call {0}' -f $partial_otm_call_strike); data = $partial_otm_calls | % { $_.impliedVolatility.ToString('N4') }; fill = $false }
                    # @{ label = ('OTM Put {0}'  -f $partial_otm_put_strike) ; data = $partial_otm_puts  | % { $_.impliedVolatility.ToString('N4') }; fill = $false }                    
                )
            }
            options = @{
                title = @{ display = $true; text = ('{0} IV term structure' -f $symbol) }
                scales = @{ }   
                plugins = @{ datalabels = @{ display = $true } }                
            }
        }
    } | ConvertTo-Json -Depth 100
    
    $result_quickchart = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    $id = ([System.Uri] $result_quickchart.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)    
    # ----------------------------------------------------------------------    

}

$result_spy = chart-term-structure SPY '2023-06-30', '2023-07-07', '2023-07-14', '2023-07-21', '2023-07-28', '2023-08-04', '2023-08-18', '2023-09-15'

$chains = $result_spy



$expiration = '2023-08-18'

function chart-by-strikes ($symbol, $expiration)
{
    
    $chain = get-options-chain $symbol $expiration
    
    # ----------------------------------------------------------------------
    $json = @{
        chart = @{
            type = 'bar'
            data = @{
                                
                labels = $chain.calls | % strike
                                
                datasets = @(

                    # @{ label = 'call delta'; data = $chain.calls | % delta }
                    # @{ label = 'put delta';  data = $chain.puts  | % delta }

                    # @{ label = 'call delta'; data = $chain.calls | % { $_.delta * $_.strike } }
                    # @{ label = 'put delta';  data = $chain.puts  | % { $_.delta * $_.strike } }                    

                    @{ label = 'call OI'; data = $chain.calls | % {  $_.openInterest } }
                    @{ label = 'put OI';  data = $chain.puts  | % { -$_.openInterest } }
                )
            }
            options = @{
                title = @{ display = $true; text = ('{0}' -f $symbol) }
                scales = @{ xAxes = @( @{ stacked = $true } ) }   
                plugins = @{ datalabels = @{ display = $false } }                
            }
        }
    } | ConvertTo-Json -Depth 100
    
    $result_quickchart = Invoke-RestMethod -Method Post -Uri 'https://quickchart.io/chart/create' -Body $json -ContentType 'application/json'
    
    $id = ([System.Uri] $result_quickchart.url).Segments[-1]
    
    Start-Process ('https://quickchart.io/chart-maker/view/{0}' -f $id)    
    # ----------------------------------------------------------------------           
}

chart-by-strikes SPY '2023-08-18'


# ----------------------------------------------------------------------


