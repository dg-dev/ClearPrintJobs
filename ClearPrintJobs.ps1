# ClearPrintJobs.ps1 v1.1 2023-06-12 by Dennis G.
 
# Yeets all print jobs when jobs over certain age are found to fix stuck printer queues.
# To be run by Task Scheduler every ~5 minutes until a better solution is found.

$jobsDir = "C:\Windows\System32\spool\PRINTERS"
$jobExts = ".spl", ".shd"
$maxJobAgeMins = 17

$scriptName = (Get-Item $PSCommandPath).BaseName
$scriptDir = (Get-Item $PSCommandPath).Directory
$logPath = "{0}\{1}.log" -f $scriptDir, $scriptName

function Write-Log {
    param(
        [Parameter()]
        $Value = ''
    )
    Add-Content -Path $logPath -Value ("{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Value) -Force
}

Write-Log "Starting"

$SpoolerService = Get-Service -ServiceName Spooler
Write-Log ('  Spooler service status: {0}' -f $SpoolerService.Status)

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
if (!$oldJobFiles) { exit 0 }

Write-Log ('  Preparing to delete all printer jobs; {0} files' -f $allJobFiles.Length)
if (-Not ($SpoolerService.Status -EQ 'Stopped')) {
    Write-Log '    Stopping Spooler service'
    $SpoolerService | Stop-Service -Force
    $SpoolerService.WaitForStatus("Stopped", '00:00:30')
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
$SpoolerService.WaitForStatus("Running", '00:00:30')
$SpoolerService.Refresh()
Write-Log ('    Spooler service status: {0}' -f $SpoolerService.Status)

exit 0