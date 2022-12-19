#Function to split an array of objects
Function Split-Array{
    Param(
    $Array,
    $ObjectLimit
)
    #Define the chunk size
    [int]$blockLimit = $ObjectLimit
    #Math the number of jobs/chunks
    $numberOfJobs = [math]::Floor($array.length / $blockLimit)
    #Get the rest/mod for the last one
    $lastJobCount = $array.length % $blockLimit
    #interaction control variable
    $i = 1
    #Define an array to store all the arrays
    $result = @()
    #Do the calculations until $i variable is greater than the number of chunks plus 1
    Do{
        #If the interaction variable is equal to the number of chunks plus 1 this is the last chunk
        if($i -eq ($numberOfJobs + 1)){
            #The variable of the first object in the current chunk becomes the rest/mod all chunks divided the chunk size
            $varFirst = $lastJobCount
            #If the array length is less than object limit, then the number of objects should be equal to array length
            if($array.length -lt $ObjectLimit){
                $numberOfObj = $array.length
            }
            #The variable of the number of objects to skip becomes the number of objects multipled by current interaction value minus 1
            $varSkip = $numberOfObj * ($i - 1)
        }else{
            #If the interaction variable is not equanto to the number of chunks plus 1 then this is not the last chunk
            #The number of objects becomes the chunk size
            $numberOfObj = $blockLimit
            #The first object is the number of objects multiplied by the current interaction
            $varFirstTmp = $numberOfObj * $i
            #The number of objects to skip is the number of the first object minus the number of objects
            $varSkip = $varFirstTmp - $numberOfObj
            #The very first object is the number of objects
            $varFirst = $numberOfObj
        }

        #Append to the main array the array of objects considering the first and last objects as the current chunk of objects (Starting in one specific and skipping some accordingly)
        $result+=,@($Array | Select-Object -First $varFirst -Skip $varSkip)

        $i++
    }
    Until($i -gt $numberOfJobs + 1)
    #Return an array of arrays as a result
    return,$result
}

Function New-PowerShellRunspace {
    Param(
        [array]$Files,
        [string]$Protocol,
        [string]$CurrentDir
    )

    begin {
        switch ($Protocol) {
            "Http" { $skipCount = 3 }
            "Smtp" { $skipCount = 4 }
            "Smtp_FE" { $skipCount = 4; write-Host "skip count 4" }
            default {} 
        }
        #total files count used for progress output
        $totalFiles = $files.Count
        #Max threads
        $maxThreads = 20
        #Creates a Runspace pool limited to 10 threads
        $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
        $RunspacePool.Open()
        #Define an array to store all runspaces
        $Jobs = @()
        #Split all files int the list in smaller chunks
        $fileChunks = Split-Array -Array $Files -ObjectLimit 10
        #Increment control variable to be used to name temp files
        $i = 1
    }
    process {

        #For each chunk of files in files chunks
        foreach($files in $fileChunks){
            $ParamList = @{
                Files = $files
                Protocol = $Protocol
                CurrentDir = $CurrentDir
                SkipCount = $skipCount
            }
            #Create the powershell runspace
            $PowerShell = [powershell]::Create()
            #Add runspace into runspace pool
            $PowerShell.RunspacePool = $RunspacePool
            #Define the script block of the current run space using the current files chunk
            $PowerShell.AddScript({
                param ($Files,$Protocol,$CurrentDir,$skipCount)
                Set-Location -Path $CurrentDir
                #For each user in array users
                foreach($file in $Files){
                    $file.FullName
                    if($file.GetType().Name -eq "FileInfo"){
                        if(!(Test-Path ".\$($Protocol)\Sanitized\$($file.Directory.Name)")) {
                            New-Item -ItemType Directory -Path ".\$($Protocol)\Sanitized\$($file.Directory.Name)"
                        }
                    }
                    (Get-Content $file.FullName -ErrorAction Stop | Select-Object -Skip $skipCount).Replace("#Fields: ", "") | Out-File $file.FullName.Replace("\$($Protocol)\","\$($Protocol)\Sanitized\").Replace(".log","_sanitized_$($file.Directory.Name).log")
                }
            })
            $PowerShell.AddParameters($ParamList)
            #Invoke the execution of current runspace and add it to an array of runspaces to allow tracking
            $jobs += $PowerShell.BeginInvoke()
            #Increment the control variable
            $i++
        }
        
        #While at least on job is not completed, wait for 2 seconds and check again. Script will not move further until all jobs are completed.
        While($Jobs.IsCompleted -contains $false){
            $sanitizedFiles = Get-ChildItem ./$($Protocol)/Sanitized -Recurse -File | Measure-Object
            Write-Host "Progress: $((($sanitizedFiles.count/$totalFiles)*100).ToString("##.##"))%. Already processed $($sanitizedFiles.Count) files. Pending to process $($totalFiles - $sanitizedFiles.count) files. Last update $(Get-Date)."
            Start-Sleep -Seconds 5
            Write-Host "done" -ForegroundColor Green
        }
    }
}

function Run-ExchangeLogCleaner {
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Tracking","Mailbox","Pop","Http","Smtp","Smtp_FE")]
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
            $folders = Get-ChildItem .\Pop -ErrorAction Stop | Where-Object FullName -NotMatch "sanitized"
            $y = 1
            foreach($folder in $folders){
                Write-Progress -Activity "Processing folder $($folder.Name)" -Status "Folder $y of $(($folders | Measure-Object).count)" -Id 1 -PercentComplete (($y/($folders | Measure-Object).count)*100)
                $logFiles = Get-ChildItem ".\POP\$($folder.Name)" -Recurse -ErrorAction Stop
                if($logFiles){
                    New-Item -ItemType Directory -Path ".\POP\Sanitized\$($folder.Name)" -ErrorAction Stop
                    $i = 1
                    foreach($item in $logFiles){
                        Write-Progress -Activity "Sanitizing file $($item.Name)" -Status "File $i of $(($logFiles | Measure-Object).count)" -Id 2 -ParentId 1 -PercentComplete (($i/($logFiles | Measure-Object).count)*100)
                        if($item.GetType().Name -eq "FileInfo"){
                            Get-Content $item -ErrorAction Stop | Select-String -Pattern '^#' -NotMatch -ErrorAction Stop | Where-Object{$_.Line} | Out-File $item.FullName.Replace("\POP\","\POP\Sanitized\").Replace(".LOG","_sanitized_$($folder.Name).LOG")
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
                $logFiles = Get-ChildItem .\HTTP -Recurse -File -ErrorAction Stop | Where-Object FullName -NotMatch "sanitized"
                $y = 1

                New-PowerShellRunspace -Files $logFiles -Protocol "HTTP" -CurrentDir (Get-Location).Path.ToString()
            }
            catch{
                Write-Host $_.Exception.Message -ForegroundColor Yellow
            }
            
        }
        "Smtp" {  
            try{
                if(!(Test-Path .\Smtp)){
                    Write-Host "No SMTP folder found. Please create a folder 'SMTP' and paste all collected data inside."
                    Exit
                }
                if(!(Test-Path .\SMTP\Sanitized)){
                    New-Item -ItemType Directory -Path .\SMTP\Sanitized -ErrorAction Stop
                }
                $logFiles = Get-ChildItem .\SMTP -Recurse -File -ErrorAction Stop | Where-Object FullName -NotMatch "sanitized"
                $y = 1

                New-PowerShellRunspace -Files $logFiles -Protocol "SMTP" -CurrentDir (Get-Location).Path.ToString()
            }
            catch{ Write-Host $_.Exception.Message -ForegroundColor Yellow }
        }
        "Smtp_FE" {  
            try{
                if(!(Test-Path .\SMTP_FE)){
                    Write-Host "No SMTP_FE folder found. Please create a folder 'SMTP_FE' and paste all collected data inside."
                    Exit
                }
                if(!(Test-Path .\SMTP_FE\Sanitized)){
                    New-Item -ItemType Directory -Path .\SMTP_FE\Sanitized -ErrorAction Stop
                }
                $logFiles = Get-ChildItem .\SMTP_FE -Recurse -File -ErrorAction Stop | Where-Object FullName -NotMatch "sanitized"
                $y = 1

                New-PowerShellRunspace -Files $logFiles -Protocol "SMTP_FE" -CurrentDir (Get-Location).Path.ToString()
            }
            catch{ Write-Host $_.Exception.Message -ForegroundColor Yellow }
        }
        Default {}
    }
}