# Written by Midas van Oene
# Contact: Midas.van.Oene@gmail.com 
# Repository: https://github.com/midasvo/CSV2SSIS

$package = "\path\to\Package.dtsx"

$SSISWorkingDirectory = "\path\to\SSISWorkingDirectory"
$MonitoringDirectory = '\path\to\MonitoringDirectory'

$filter = '*.*'  # Wildcard filter                        
$fsw = New-Object IO.FileSystemWatcher $MonitoringDirectory, $filter -Property @{IncludeSubdirectories = $false;NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'} 

$timer = New-Object timers.timer
$timer.Interval = 5000

$global:files=@()

Write-Host "Starting script" -fore green

function Run-SSIS-Package {
    dtexec /f $package # Execute SSIS Package
    Write-Host $LASTEXITCODE
    if($LASTEXITCODE -eq 0) {
        Write-Host 'Executed SSIS package succesfully. Time for cleanup...' -fore green

        Write-Host 'Removing files from SSIS Staging folder' -fore green
        Remove-Item $SSISWorkingDirectory\*
        
        Write-Host 'Clearing files array' -fore green
        $global:files = @()
    } else { # add error handling for the other error codes
        Write-Host "Encountered an error ('$LASTEXITCODE') while executing SSIS package.. " -fore red
        Write-Host 'Removing files from SSIS Staging folder'
        Remove-Item $SSISWorkingDirectory\*
        
        Write-Host 'Clearing files array'
        $global:files = @()
    }
}

function Get-Machine-Name($logfilename) {

    Write-Host $MonitoringDirectory\$logfilename -foreground Green
    $data = Import-CSV $MonitoringDirectory\$logfilename -Delimiter ';'
    $machinename = $data[0].Machine # Assume a file is for one machine, so only need first element

    return $machinename
}

Register-ObjectEvent -InputObject $timer -EventName Elapsed -SourceIdentifier Timer.Output -Action {
    $timer.Enabled = $False
    foreach ($name in $global:files) {
        $machinename = Get-Machine-Name($name)
        Write-Host 'Machinename: ' $machinename
        Write-Host 'Copying ' $name 'to SSIS location and renaming'
        Copy-Item $MonitoringDirectory\$name -Destination $SSISWorkingDirectory\$machinename'.csv'
    }
    Write-Host 'Start SSIS package now' -fore green
    Run-SSIS-Package
}

Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated -Action { 
    $timer.Enabled = $False

    $name = $Event.SourceEventArgs.Name 
    $timeStamp = $Event.TimeGenerated 

    Write-Host "Detected new file '$name' on $timeStamp" -fore green 

    Write-Host 'File: '$MonitoringDirectory\$name -foreground Green    
    
    $global:files+=$name
    $timer.Enabled = $True
} 

    