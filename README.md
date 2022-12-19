# ExchangeUsageCollector
Log collector for various Exchange on premises workloads to define the usage of them.

## Log Collector

The script file .\Collector\Run-ExchangeLogCollector.ps1 can be used to collect Exchange protocol log files from all servers in the organization. This script will copy all files from remote servers to the localhost. For each protocol log, the script will proceed as follows:

### POP

The script will create a folder "POP" in the running directory. Inside "POP" folder, it creates a folder with the server name for each Exchange server in the organization and copies all POP logs from the servers into the folder with the respective server name. The script will look in the (Get-PopSettings SERVER).LogFileLocation to determine where in each Exchange Server the POP logs are.

### Http

Same as POP, but now it will create a folder "HTTP" and proceed as it does with POP. To determine where HTTP logs are, it used WebAdministration PowerShell module to remotely invoke the command (Get-Item 'IIS:\Sites\Default Web Site').logFile.Directory.

### Smtp

Same as POP and HTTP, but using command (Get-TransportService SERVER).ReceiveProtocolLogPath to determine where the Receive Connector's verbose protocol logging are.

### Smtp_FE

Sam as POP, HTTP and SMTP, but using command Get-FrontendTransportService to determine where the Frontend Receive Connector's verbose protocol logging are.

### Tracking

It uses Get-MessageTrackingLog to consolidate all tracking logs into one .csv file named as .\trackingLogs_$(Get-Date -Format 'dd-MM-yyyy_hh-mm-ss').csv"

### Mailboxes

It runs Get-Mailbox to list all mailboxes in the environment. It will be used to check if there is senders that are not a mailbox in the enviroment.

### How to run the collector?

To run the collector, just run the script providing which logs you'd like to collect using LogType parameter, as follows:

.\Run-ExchangeLogCollector.ps1 -LogType Pop,Http,Smtp,Tracking,Mailbox

## Log Analyzer

The log analyzer script .\Analyzer\Run-ExchangeLogAnalyzer.ps1 is intended to sanitize a few information from log files either not important or that would mess with the log processing in PowerBI later. You just need to copy/move each protocol log folder into .\Analyzer directory in order to start. For instance, if you collect all protocol logs: Pop, Http and Smtp, just copy/move all three folders into .\Analyzer and run the script as follows:

.\Run-ExchangeLogAnalyzer.ps1 -LogType Pop,Http,Smtp

It will start to clean all logs up. For each log folder, it creates a folder "sanitized", for example, for Pop it will be .\Analyzer\Pop\Sanitized\SERVERNAME. After all logs are cleaned, will can go ahead and point the Sanitized folder for each log as a "Folder" data source in PowerBI to start analyze it.