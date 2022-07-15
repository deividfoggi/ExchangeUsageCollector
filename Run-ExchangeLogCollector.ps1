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
        try{
            $popServersSettings = $exchangeServers | %{Get-PopSettings -Server $_.Name}
            foreach($popServerSettings in $popServersSettings) {
                $popLogNetworkPath = "\\$($popServerSettings.Server)\" + $popServerSettings.LogFileLocation.Replace(":","$")
                try{
                    if((Get-PopSettings -Server $popServerSettings.Server -ErrorAction Stop).ProtocolLogEnabled -eq $false){
                        Set-PopSettings -Server $popServerSettings.Server -ProtocolLogEnabled $true -ErrorAction Stop
                        Invoke-Command -ComputerName $popServerSettings.Server -ScriptBlock {Restart-Service MSExchangePop3;Restart-Service MSexchangePop3BE} -ErrorAction Stop
                    }
                    $popSourceFiles = Get-ChildItem -Path $popLogNetworkPath -ErrorAction Stop | ?{$_.LastWriteTime -gt (Get-Date).AddDays(-7)}
                    if(!(Test-Path -Path .\POP)){
                        New-Item -ItemType Directory -Name "POP" -Path .\ -ErrorAction Stop
                    }
                    if(!(Test-Path -Path .\POP\$($popServerSettings.Server))){
                        New-Item -ItemType Directory -Name "$($popServerSettings.Server)" -Path .\POP -ErrorAction Stop
                    }
                    foreach($file in $popSourceFiles){
                        Copy-Item $file.FullName -Destination .\POP\$($popServerSettings.Server) -ErrorAction Stop
                    }                
                }catch{
                    Write-Host $_.Exception.GetType() -ForegroundColor Yellow
                }
            }
        }catch{
            Write-Host $_.Exception.GetType() -ForegroundColor Yellow
        }
        
    }
    "Http" {
        try{
            foreach($server in $exchangeServers) {
                $iisLogPath = Invoke-Command -ComputerName $server.Name -ScriptBlock{
                    Import-Module WebAdministration
                    (Get-Item 'IIS:\Sites\Default Web Site').logFile.Directory
                } -ErrorAction Stop
                if($iisLogPath.Split("\")[0] -match ":") {
                    $iisNetworkLogPath = "\\$($server.Name)\" + $iisLogPath.Replace(":","$") + "\W3SVC1"
                } elseif($iisLogPath.Split("\")[0] -match "%SystemDrive%") {
                    $iisNetworkLogPath = "\\$($server.Name)\" + $iisLogPath.Replace("%SystemDrive%","C$") + "\W3SVC1"
                }
                $httpSourceFiles = Get-ChildItem -Path $iisNetworkLogPath -ErrorAction Stop | ?{$_.LastWriteTime -gt (Get-Date).AddDays(-7)}
                if(!(Test-Path -Path .\HTTP)){
                    New-Item -ItemType Directory -Name "HTTP" -Path .\ -ErrorAction Stop
                }
                if(!(Test-Path -Path .\HTTP\$($server.Name))){
                    New-Item -ItemType Directory -Name "$($server.Name)" -Path .\HTTP -ErrorAction Stop
                }
                foreach($file in $httpSourceFiles){
                   Copy-Item $file.FullName -Destination .\HTTP\$($server.Name) -ErrorAction Stop
                } 
            }
        }catch{
            Write-Host $_.Exception.GetType() -ForegroundColor Yellow
        }
        
    }
    "Smtp" {
        foreach($server in $exchangeServers) {
            $receiveProtocolLogPath = Get-TransportService $server.Name
            $receiveProtocolLogNetworkPath = "\\$($server.Name)\" + $receiveProtocolLogPath.ReceiveProtocolLogPath.Replace(":","$")
            $smtpSourceFiles = Get-ChildItem -Path $receiveProtocolLogNetworkPath -ErrorAction Stop | where LastWriteTime -gt (Get-Date).AddDays(-7)
            if(!(Test-Path -Path .\Smtp)){
                New-Item -ItemType Directory -Name "SMTP" -Path .\ -ErrorAction Stop
            }
            if(!(Test-Path -Path .\SMTP\$($server.Name))){
                New-Item -ItemType Directory -Name "$($server.Name)" -Path .\SMTP -ErrorAction Stop
            }
            foreach($file in $smtpSourceFiles){
                Copy-Item $file.FullName -Destination .\SMTP\$($server.Name) -ErrorAction Stop
            } 
        }
    }
    Default {
        Write-Host "Use parameter 'Log' to choose which log to collect: Tracking, " -ForegroundColor Yellow
    }
}
