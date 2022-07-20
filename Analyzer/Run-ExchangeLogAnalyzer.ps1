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
        $folders = Get-ChildItem .\HTTP -ErrorAction Stop | where FullName -NotMatch "sanitized"
        $y = 1
        foreach($folder in $folders){
            Write-Progress -Activity "Processing folder $($folder.Name)" -Status "Folder $y of $($folders.Length)" -Id 1 -PercentComplete (($y/$folders.Length)*100)
            $logFiles = Get-ChildItem ".\POP\$($folder.Name)" -Recurse -ErrorAction Stop
            if($logFiles){
                New-Item -ItemType Directory -Path ".\POP\Sanitized\$($folder.Name)" -ErrorAction Stop
                $i = 1
                foreach($item in $logFiles){
                    Write-Progress -Activity "Sanitizing file $($item.Name)" -Status "File $i of $($logFiles.Length)" -Id 2 -ParentId 1 -PercentComplete (($i/$logFiles.Length)*100)
                    if($item.GetType().Name -eq "FileInfo"){
                        Get-Content $item -ErrorAction Stop | Select-String -Pattern '^#' -NotMatch -ErrorAction Stop | %{$_.Line} | Out-File $item.FullName.Replace("\POP\","\POP\Sanitized\").Replace(".LOG","_sanitized_$($folder.Name).LOG")
                    }
                    $i++
                }
            }
            $y++
        }
     }
    "Http" { 
        try{
            if(!(Test-Path .\HTTP)){
                Write-Host "No HTTP folder found. Please create a folder 'HTTP' and paste all collected data inside."
                Exit
            }
            if(!(Test-Path .\HTTP\Sanitized)){
                New-Item -ItemType Directory -Path .\HTTP\Sanitized -ErrorAction Stop
            }
            $folders = Get-ChildItem .\HTTP -ErrorAction Stop | where FullName -NotMatch "sanitized"
            $y = 1
            foreach($folder in $folders){
                Write-Progress -Activity "Processing folder $($folder.Name)" -Status "Folder $y of $($folders.Length)" -Id 1 -PercentComplete (($y/$folders.Length)*100)
                $logFiles = Get-ChildItem ".\HTTP\$($folder.Name)" -Recurse -ErrorAction Stop
                if($logFiles){
                    New-Item -ItemType Directory -Path ".\HTTP\Sanitized\$($folder.Name)" -ErrorAction Stop
                    $i = 1
                    foreach($item in $logFiles){
                        Write-Progress -Activity "Sanitizing file $($item.Name)" -Status "File $i of $($logFiles.Length)" -Id 2 -ParentId 1 -PercentComplete (($i/$logFiles.Length)*100)
                        if($item.GetType().Name -eq "FileInfo"){
                            (Get-Content $item -ErrorAction Stop).Replace("#Fields: ", "") | Select-String -Pattern '^#' -NotMatch | %{$_.Line} | Out-File $item.FullName.Replace("\HTTP\","\HTTP\Sanitized\").Replace(".log","_sanitized_$($folder.Name).log")
                        }
                        $i++
                    }
                }
                $y++
            }
        }
        catch{
            Write-Host $_.Exception.Message -ForegroundColor Yellow
        }
        
     }
    "Smtp" {  }
    Default {}
}