Param(
    $Log
)
if (!(Get-PSSession | where ConfigurationName -eq "Microsoft.Exchange")) {
    try{
        Import-PSSession (New-PSSession -ConnectionUri http://fogex01/Powershell -ConfigurationName Microsoft.Exchange -Authentication kerberos) -ErrorAction Stop
    }catch{
        Write-Host $_.Exception.Message
    }
}

$exchangeServers = Get-ExchangeServer | where ServerRole -NotMatch "edge"

switch ($Log) {
    "Tracking" {
        $trackingLogs = @()
        $exchangeServers | %{
            try{
                $trackingLogs += Get-MessageTrackingLog -Server $_.Name -EventId Receive -Start (Get-Date).AddDays(-5) -End (Get-Date) -ResultSize Unlimited -ErrorAction Stop | ?{$_.Sender -NotMatch "HealthMailbox" -And $_.Sender -NotMatch "InboundProxy" -And $_.SourceContext -NotMatch "System Probe"}
            }catch{
                Write-Host $_.Exception.Message -ForegroundColor Red
            }
        }
        
        $trackingLogs | Export-Csv .\trackingLogs_$(Get-Date -Format 'dd-MM-yyyy_hh-mm-ss').csv -NoTypeInformation
    }
    "Mailbox" {
        try{
            $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | Select Name, Alias, PrimarySmtpAddress
            $mailboxes | Export-Csv .\mailboxes_$(Get-Date -Format 'dd-MM-yyyy_hh-mm-ss').csv -NoTypeInformation
        }catch{
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }
    "POP" {
        $popServersSettings = $exchangeServer | %{Get-PopSettings -Server $_.Name}
        foreach($popServerSettings in $popServersSettings) {
            $popLogNetworkPath = "\\$($popServerSettings.Server)\" + $popServerSettings.LogFileLocation.Replace(":","$")
            try{
                Set-PopSettings -Server $popServerSettings.Server -ProtocolLogEnabled $true -ErrorAction Stop
                $sourceFiles = Get-ChildItem -Path $popLogNetworkPath -ErrorAction Stop
                New-Item -ItemType Directory -Name "POP" -Path .\ -ErrorAction Stop
                New-Item -ItemType Directory -Name "$($popServerSettings.Server)" -Path .\POP -ErrorAction Stop
                Copy-Item $sourceFiles -Destination .\POP\$($popServerSettings.Server)
            }catch{
                Write-Host $_.Exception.GetType() -ForegroundColor Yellow
            }
            
        }
    }
    Default {
        Write-Host "Use parameter 'Log' to choose which log to collect: Tracking, " -ForegroundColor Yellow
    }
}
