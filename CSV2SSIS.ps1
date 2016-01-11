# Written by Midas van Oene
# Contact: Midas.van.Oene@gmail.com 
# Repository: https://github.com/midasvo/CSV2SSIS
$package = "\path\to\Package.dtsx"

$SSISWorkingDirectory = "\path\to\SSISWorkingDirectory"
$MonitoringDirectory = "\path\to\MonitoringDirectory"

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
        Send-Mail($global:files)
        Write-Host 'Removing files from SSIS Staging folder' -fore green
        Remove-Item $SSISWorkingDirectory\*
        Write-Host 'Clearing files array' -fore green
        $global:files = @()
    } else { # add error handling for the other error codes
        Write-Host "Encountered an error ('$LASTEXITCODE') while executing SSIS package.. " -fore red
        Send-Mail($global:files)
        Write-Host 'Not removing files from SSIS Staging folder, please check it manually.' -fore red        
        Write-Host 'Clearing files array' -fore green
        $global:files = @()
    }
}

function Send-Mail($report) {
    $exitcodes = @{
        0 = "The package executed successfully"; 
        1 = "The package failed";     
        3 = "The package was canceled by the user";     
        4 = "The utility was unable to locate the requested package. The package could not be found";     
        5 = "The utility was unable to load the requested package. The package could not be loaded"; 
        6 = "The utility encountered an internal error of syntactic or semantic errors in the command line"
        }

    Write-Host -NoNewline "Sending mail... "
    $date = Get-Date
    $exitcode_descr = $exitcodes.Get_Item($LASTEXITCODE)

    $emailSmtpServer = "smtp.gmail.com"
    $emailSmtpServerPort = "587"
    $emailSmtpUser = "user"
    $emailSmtpPass = "password"
 
    $emailMessage = New-Object System.Net.Mail.MailMessage
    $emailMessage.From = "AUTOCSV2SSIS Service <autocsv2ssis@midasvo.nl>"
    $emailMessage.To.Add( "recipient@gmail.com" )
    $emailMessage.Subject = "[AUTOCSV2SSIS] Execution report from $date"
    $emailMessage.IsBodyHtml = $true
    if($LASTEXITCODE -eq 0) {
        $emailMessage.Body = @"
<p>Hello,</p>
<p>Your SSIS package returned an exitcode of $LASTEXITCODE ($exitcode_descr) on $date.</p>
<p>The following files were processed by SSIS: $report.</p>
<p>Package: <strong>$package</strong></p>
"@
     
    } else {
        $emailMessage.Body = @"
<p>Hello,</p>
<p>Your SSIS package returned an exitcode of $LASTEXITCODE ($exitcode_descr) on $date.</p>
<p>Your package failed. The files have not been deleted from the SSIS Working Directory. Please run SSIS manually to troubleshoot.</p>
<p>Package: <strong>$package</strong></p>
"@
    }

    $SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
 
    $SMTPClient.Send( $emailMessage )

    Write-Host "Sent on $date"
}

function Get-Machine-Name($logfilename) {

    Write-Host "Hello " $MonitoringDirectory\$logfilename -foreground Green
    $data = Import-CSV $MonitoringDirectory\$logfilename -Delimiter ';'
    $machinename = $data[1].Machine # Name of the row that will become the filename, SSIS points to this file.
    Write-Host "Filename: $machinename"
    return $machinename
}

Register-ObjectEvent -InputObject $timer -EventName Elapsed -SourceIdentifier Timer.Output -Action {
    $timer.Enabled = $False
    foreach ($name in $global:files) {
        $machinename = Get-Machine-Name($name)
        Write-Host $machinename #change to $extractedname for clarity
        Write-Host 'Copying ' $name 'to SSIS location and renaming'
        Copy-Item $MonitoringDirectory\$name -Destination $SSISWorkingDirectory\$machinename'.csv'
    }
    Write-Host 'Start SSIS package now' -fore green
    Run-SSIS-Package
}

Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated -Action { 
    $timer.Enabled = $False

    $filename = $Event.SourceEventArgs.Name # change to $filename for clarity
    $timeStamp = $Event.TimeGenerated 

    Write-Host "Detected new file '$filename' on $timeStamp" -fore green 

    Write-Host 'File: '$MonitoringDirectory\$filename -foreground Green    
    
    $global:files+=$filename
    $timer.Enabled = $True
} 

    