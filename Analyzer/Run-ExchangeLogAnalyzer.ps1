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
        $zipFiles = Get-ChildItem .\POP
        foreach($file in $zipFiles){
            Expand-Archive -Path $file.FullName -DestinationPath ".\POP\$($file.Name.Replace('.zip',''))"
        }
        
     }
    "Http" {  }
    "Smtp" {  }
    Default {}
}