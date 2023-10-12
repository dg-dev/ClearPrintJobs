# ClearPrintJobs.ps1 v1.3 2023-07-05 by Dennis G.
 
# Fixes stuck printer queues. Yeets all print jobs when erroneous or old jobs are found.
# To be run by Task Scheduler on an interval until a better solution is found.

$jobsDir = "C:\Windows\System32\spool\PRINTERS"
$jobExts = ".spl", ".shd"
$maxJobAgeMins = 17
$maxLogSizeBytes = 1024 * 1024 * 10
$logMoreDetails = $True

$scriptName = (Get-Item $PSCommandPath).BaseName
$scriptDir = (Get-Item $PSCommandPath).Directory
$logDir = "{0}\Logs" -f $scriptDir
$logPath = "{0}\{1}.log" -f $logDir, $scriptName

function Write-Log {
    param(
        [Parameter()]
        $Value = ''
    )
    $finalOutput = ("{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Value)
    Write-Host $finalOutput
    Add-Content -Path $logPath -Value $finalOutput -Force
}

Write-Log "Starting"

if (-Not (Test-Path -Path $logDir -PathType Container)) {
    New-Item -ItemType Directory -Path $logDir
}

# Rotate log
if ((Test-Path -Path $logPath -PathType Leaf) -and ((Get-Item -Path $logPath).length -ge $maxLogSizeBytes)) {
    $archivePath = "{0}\{1}-{2}.zip" -f $logDir, $scriptName, (Get-Date -Format "yyyyMMddHHmmss")
    Write-Log "  Rotating log"
    Write-Log ("    Archiving to {0}" -f $archivePath)
    try {
        Compress-Archive -Path $logPath -CompressionLevel Optimal -DestinationPath $archivePath -ErrorAction Stop
        Remove-Item -Path $logPath -Force -ErrorAction Stop
        Write-Log "Continuing"
    } catch {
        Write-Log ("     ERROR: {0}" -f $Error[0])
    }
}

$SpoolerService = Get-Service -ServiceName Spooler
Write-Log ('  Spooler service status: {0}' -f $SpoolerService.Status)

Write-Log -Value ("  Checking printers")
$allPrinters = @(Get-Printer | Where-Object -Property Shared -EQ $True)
$abnormalPrinters = @($allPrinters | Where-Object -Property PrinterStatus -NE "Normal")
Write-Log -Value ("    Found {0} out of {1} printers with abnormal status" -f $abnormalPrinters.Length, $allPrinters.Length)
Write-Log -Value ("  Checking print job queue")
$allPrintJobs = @($allPrinters | Get-PrintJob)
$abnormalPrintJobs = @($allPrintJobs | Where-Object -Property JobStatus -NE "Normal")
Write-Log -Value ("    Found {0} out of {1} jobs with abnormal status" -f $abnormalPrintJobs.Length, $allPrintJobs.Length)
#$allPrintJobs | ForEach-Object -Process { }
Write-Log -Value ("  Checking {0}" -f $jobsDir)
$allJobFiles = Get-ChildItem -Path $jobsDir -File -Force |
    Where-Object -FilterScript { $PSItem.Extension -in $jobExts }
$oldJobFiles = $allJobFiles | 
    Where-Object -FilterScript { $PSItem.CreationTime -lt (Get-Date).AddMinutes(-1 * $maxJobAgeMins) }
Write-Log -Value ("    Found {1} out of {0} job files exceeding {2} minutes" -f $allJobFiles.Length, $oldJobFiles.Length, $maxJobAgeMins)
$allJobFiles | ForEach-Object -Process {
        $ageMins = (New-TimeSpan -Start $PSItem.CreationTime -End (Get-Date)).TotalMinutes
        #$ageString = ''
        #if ($ageMins -gt $maxJobAgeMins) { $ageString = ' [{0} mins]' -f [int]$ageMins }
        $ageNote = ''
        if ($ageMins -gt $maxJobAgeMins) { $ageNote = ' !!!' }
        $ageString = ' [{0} mins]' -f [int]$ageMins
        Write-Log -Value ("      {0}{1}{2}" -f $PSItem.Name, $ageString, $ageNote)
    }
if ($allJobFiles -and $logMoreDetails) {
    Write-Log -Value ("    Additional information")
    Write-Log -Value ("      Start")
    #$additionalInfo = Get-Printer | Where-Object -Property Shared -EQ True | Where-Object -Property JobCount -NE 0
    $additionalInfo = Get-Printer | Where-Object -Property Shared -EQ $True
    $additionalInfo | Format-List | Out-File -FilePath $logPath -Encoding utf8 -Append -NoClobber -Force
    $additionalInfo | Format-List
    $additionalInfo = $additionalInfo | Get-PrintJob
    $additionalInfo | Format-List | Out-File -FilePath $logPath -Encoding utf8 -Append -NoClobber -Force
    $additionalInfo | Format-List
    Write-Log -Value ("      End")
}
#if (!$oldJobFiles -and !$abnormalPrintJobs) { exit 0 }
if (!$oldJobFiles) { exit 0 }

Write-Log ('  Preparing to delete all printer jobs; {0} files' -f $allJobFiles.Length)
if (-Not ($SpoolerService.Status -EQ 'Stopped')) {
    Write-Log '    Stopping Spooler service'
    $SpoolerService | Stop-Service -Force
    $SpoolerService.WaitForStatus("Stopped", '00:00:15')
    $SpoolerService.Refresh()
    Write-Log ("    Spooler service status: {0}" -f $SpoolerService.Status)
}
if ($SpoolerService.Status -EQ "Stopped") {
    Write-Log ("    Deleting {0} job files" -f $allJobFiles.Length)
    $allJobFiles | Remove-Item -Force
    $allJobFiles | ForEach-Object -Process {
            $jobFile = $PSItem
            if (-Not (Test-Path -Path $jobFile.FullName -PathType Leaf)) {
                Write-Log ("       {0} [DELETED]" -f $jobFile.Name)
            } else {
                Write-Log ("       {0} [ERROR]" -f $jobFile.Name)
            }
        }
} else {
    Write-Log ("    Unable to stop Spooler service and nothing was deleted")
}
Write-Log '    Starting Spooler service'
$SpoolerService | Start-Service
$SpoolerService.WaitForStatus("Running", '00:00:15')
$SpoolerService.Refresh()
Write-Log ('    Spooler service status: {0}' -f $SpoolerService.Status)

exit 0
