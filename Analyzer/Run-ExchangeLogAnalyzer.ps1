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

        $logFiles = Get-ChildItem .\POP -Recurse | where name -NotMatch ".zip"
        $i = 1
        foreach($item in $logFiles){
            Write-Progress -Activity "Sanitizing file $($item.Name)" -Status "File $i of $($logFiles.Length)" -Id 1 -PercentComplete (($i/$logFiles.Length)*100)
            if($item.GetType().Name -eq "FileInfo"){
                Get-Content $item | Select-String -Pattern '^#' -NotMatch | %{$_.Line} | Out-File $item.FullName.Replace(".LOG","_sanitized.LOG").Replace("/POP/","/POP/Sanitized/")
                Remove-Item $item -Confirm:$false
            }
            $i++
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
                New-Item -ItemType Directory -Path ".\HTTP\Sanitized\$($folder.Name)" -ErrorAction Stop
                $i = 1
                foreach($item in $logFiles){
                    Write-Progress -Activity "Sanitizing file $($item.Name)" -Status "File $i of $($logFiles.Length)" -Id 2 -ParentId 1 -PercentComplete (($i/$logFiles.Length)*100)
                    if($item.GetType().Name -eq "FileInfo"){
                        (Get-Content $item -ErrorAction Stop).Replace("#Fields: ", "") | Select-String -Pattern '^#' -NotMatch | %{$_.Line} | Out-File $item.FullName.Replace("\HTTP\","\HTTP\Sanitized\").Replace(".log","_sanitized_$($folder.Name).log")
                        #Get-Content $item.FullName.Replace("\HTTP\","\HTTP\Sanitized\").Replace(".log","_sanitizedTEMP.log") | Select-String -Pattern '^#' -NotMatch | Out-File $item.FullName.Replace("\HTTP\","\HTTP\Sanitized\").Replace("_sanitizedTEMP.log","_$($folder.Name)_sanitized.log") -ErrorAction Stop
                        #Remove-Item $item.FullName.Replace("\HTTP\","\HTTP\Sanitized\").Replace(".log","_sanitizedTEMP.log") -ErrorAction Stop
                    }
                    $i++
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