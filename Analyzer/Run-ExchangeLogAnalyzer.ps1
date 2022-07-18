Param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Tracking","Mailbox","Pop","Http","Smtp")]
    [string[]]$LogType
)

switch ($LogType) {
    "Pop" { 
        if(!(Test-Path .\POP)){
            Write-Host "No POP folder found. Please create a folder 'POP' and paste all collected data inside."
            Exit
        }
        $zipFiles = Get-ChildItem .\POP | where name -Match ".zip" 
        foreach($file in $zipFiles){
            Expand-Archive -Path $file.FullName -DestinationPath ".\POP\$($file.Name.Replace('.zip',''))" -Force
        }
        $logFiles = Get-ChildItem .\POP -Recurse | where name -NotMatch ".zip"
        foreach($item in $logFiles){
            if($item.GetType().Name -eq "FileInfo"){
                $content = Get-Content $item | ForEach-Object{
                    if($line -match '^#'){
                        $line -replace $line, ''
                    }
                }
                $content | Out-File $item.FullName.Replace(".LOG","_analyzed.LOG")
            }
        }
     }
    "Http" {  }
    "Smtp" {  }
    Default {}
}