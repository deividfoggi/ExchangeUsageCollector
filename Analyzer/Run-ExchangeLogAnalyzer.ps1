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
        if(!(Test-Path .\POP\Sanitized)){
            New-Item -ItemType Directory -Path .\POP\Sanitized
        }
        $zipFiles = Get-ChildItem .\POP | where name -Match ".zip" 
        foreach($file in $zipFiles){
            Expand-Archive -Path $file.FullName -DestinationPath ".\POP\Sanitized\$($file.Name.Replace('.zip',''))" -Force
        }
        $logFiles = Get-ChildItem .\POP -Recurse | where name -NotMatch ".zip"
        $i = 1
        foreach($item in $logFiles){
            Write-Progress -Activity "Sanitizing file $($item.Name)" -Status "File $i of $($logFiles.Length)" -Id 1 -PercentComplete (($i/$logFiles.Length)*100)
            if($item.GetType().Name -eq "FileInfo"){
                Get-Content $item | Select-String -Pattern '^#' -NotMatch | Out-File $item.FullName.Replace(".LOG","_sanitized.LOG")
                Remove-Item $item -Confirm:$false
            }
            $i++
        }
     }
    "Http" {  }
    "Smtp" {  }
    Default {}
}