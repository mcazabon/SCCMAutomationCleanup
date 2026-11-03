#region Script Configuration
class ScriptConfig {
    [string]$ScriptRoot
    [string]$ScriptPath
    [string]$LogDir
    [string]$ReportDir
    [string]$LogFile
    [string]$XmlPath
    [string]$Environment
    [int]$Days
    [datetime]$Timestamp

    ScriptConfig([string]$scriptRoot, [string]$environment, [int]$days) {
        $this.ScriptRoot = $scriptRoot
        $this.ScriptPath = $scriptRoot
        $this.LogDir = Join-Path $scriptRoot "logs"
        $this.ReportDir = Join-Path $scriptRoot "reports"

        # Create daily log file with timestamp and environment
        $dateStamp = Get-Date -Format "yyyy-MM-dd"
        $envUpper = $environment.ToUpper()
        $this.LogFile = Join-Path $this.LogDir "SCCMCollectionScan_${envUpper}_$dateStamp.log"

        $this.XmlPath = Join-Path $scriptRoot "Library\Resources.xml"
        $this.Environment = $envUpper
        $this.Days = $days
        $this.Timestamp = Get-Date
    }
}
#endregion

#region Logging Framework
class Logger {
    [string]$LogFile
    [string]$LogDir
    [string]$BaseLogName
    [string]$Environment
    [System.Collections.ArrayList]$LogBuffer
    [int]$BufferSize = 100
    [datetime]$CurrentLogDate
    [int]$MaxLogAgeDays = 30  # Keep logs for 30 days
    [bool]$VerboseEnabled = $false

    Logger([string]$logFile, [string]$environment) {
        $this.LogFile = $logFile
        $this.LogDir = Split-Path $logFile -Parent
        $this.BaseLogName = "SCCMCollectionScan"
        $this.Environment = $environment.ToUpper()
        $this.LogBuffer = New-Object System.Collections.ArrayList
        $this.CurrentLogDate = Get-Date -Hour 0 -Minute 0 -Second 0

        # Ensure log directory exists
        if (-not (Test-Path $this.LogDir)) {
            New-Item -ItemType Directory -Path $this.LogDir -Force | Out-Null
        }

        # Clean up old log files
        $this.CleanupOldLogs()

        # Write session header
        $this.WriteSessionHeader()
    }

    [void] WriteSessionHeader() {
        $separator = "=" * 80

        # Get PowerShell version using Get-Variable cmdlet
        try {
            $psVersion = (Get-Variable -Name PSVersionTable -ValueOnly -ErrorAction Stop).PSVersion.ToString()
        }
        catch {
            $psVersion = "Unknown"
        }

        $sessionInfo = @(
            $separator
            "SCCM Collection Scan - Session Started"
            "Environment: $($this.Environment)"
            "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            "User: $env:USERNAME"
            "Computer: $env:COMPUTERNAME"
            "PowerShell Version: $psVersion"
            $separator
        )

        $sessionInfo -join "`r`n" | Out-File -FilePath $this.LogFile -Append -Encoding ASCII
    }

    [void] CheckLogRotation() {
        $currentDate = Get-Date -Hour 0 -Minute 0 -Second 0

        # If date has changed, rotate to new log file
        if ($currentDate -gt $this.CurrentLogDate) {
            # Flush current buffer
            $this.Flush()

            # Update to new log file with environment
            $dateStamp = Get-Date -Format "yyyy-MM-dd"
            $this.LogFile = Join-Path $this.LogDir "$($this.BaseLogName)_$($this.Environment)_$dateStamp.log"
            $this.CurrentLogDate = $currentDate

            # Write new session header
            $this.WriteSessionHeader()

            # Clean up old logs
            $this.CleanupOldLogs()
        }
    }

    [void] CleanupOldLogs() {
        try {
            $cutoffDate = (Get-Date).AddDays(-$this.MaxLogAgeDays)
            $logFiles = Get-ChildItem -Path $this.LogDir -Filter "$($this.BaseLogName)_$($this.Environment)_*.log" -ErrorAction SilentlyContinue

            foreach ($logFile in $logFiles) {
                if ($logFile.LastWriteTime -lt $cutoffDate) {
                    Remove-Item $logFile.FullName -Force -ErrorAction SilentlyContinue
                    Write-Verbose "Deleted old log file: $($logFile.Name)"
                }
            }
        }
        catch {
            # Silently fail if cleanup fails
            Write-Verbose "Log cleanup failed: $_"
        }
    }

    [void] WriteLog([string]$level, [string]$message, [string]$component) {
        # Check if we need to rotate to a new log file
        $this.CheckLogRotation()

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $logEntry = "$($timestamp)|$($level)|$($component)|$message"

        # Add to buffer
        [void]$this.LogBuffer.Add($logEntry)

        # Flush if buffer is full
        if ($this.LogBuffer.Count -ge $this.BufferSize) {
            $this.Flush()
        }

        # Also write to console for errors and warnings
        if ($level -in @('ERROR', 'FATAL', 'WARNING')) {
            $consoleMessage = "[$timestamp] $level - $component : $message"
            Write-Host $consoleMessage -ForegroundColor $(
                switch ($level) {
                    'ERROR' { 'Red' }
                    'FATAL' { 'DarkRed' }
                    'WARNING' { 'Yellow' }
                    default { 'White' }
                }
            )
        }
    }

    [void] Info([string]$message, [string]$component = "General") {
        $this.WriteLog("INFO", $message, $component)
    }

    [void] Warning([string]$message, [string]$component = "General") {
        $this.WriteLog("WARNING", $message, $component)
    }

    [void] Error([string]$message, [string]$component = "General") {
        $this.WriteLog("ERROR", $message, $component)
    }

    [void] Fatal([string]$message, [string]$component = "General") {
        $this.WriteLog("FATAL", $message, $component)
        $this.Flush()
    }

    [void] Debug([string]$message, [string]$component = "General") {
        # The module has its own scope, so the caller's -Verbose is passed in through VerboseEnabled
        $verbosePref = Get-Variable -Name VerbosePreference -ValueOnly -ErrorAction SilentlyContinue
        if ($this.VerboseEnabled -or $verbosePref -eq 'Continue') {
            $this.WriteLog("DEBUG", $message, $component)
        }
    }

    [void] Success([string]$message, [string]$component = "General") {
        $this.WriteLog("SUCCESS", $message, $component)
    }

    [void] Flush() {
        if ($this.LogBuffer.Count -gt 0) {
            try {
                $this.LogBuffer | Out-File -FilePath $this.LogFile -Append -Encoding ASCII
                $this.LogBuffer.Clear()
            }
            catch {
                Write-Error "Failed to flush log buffer: $_"
            }
        }
    }

    [string] GetCurrentLogFile() {
        return $this.LogFile
    }
}
#endregion

#region Progress Tracking Framework
class ProgressTracker {
    [string]$Activity
    [int]$TotalSteps
    [int]$CurrentStep
    [hashtable]$StepWeights
    [datetime]$StartTime
    [Logger]$Logger

    ProgressTracker([string]$activity, [Logger]$logger) {
        $this.Activity = $activity
        $this.TotalSteps = 0
        $this.CurrentStep = 0
        $this.StepWeights = @{}
        $this.StartTime = Get-Date
        $this.Logger = $logger
    }

    [void] DefineSteps([hashtable]$steps) {
        $this.StepWeights = $steps
        $this.TotalSteps = ($steps.Values | Measure-Object -Sum).Sum
    }

    [void] UpdateProgress([string]$stepName, [string]$status, [int]$percentComplete) {
        if ($this.StepWeights.ContainsKey($stepName)) {
            $stepWeight = $this.StepWeights[$stepName]
            $baseProgress = ($this.CurrentStep / $this.TotalSteps) * 100
            $stepProgress = ($percentComplete / 100) * ($stepWeight / $this.TotalSteps) * 100
            $totalProgress = [math]::Min(100, [math]::Round($baseProgress + $stepProgress, 0))

            $elapsed = (Get-Date) - $this.StartTime
            $estimatedTotal = if ($totalProgress -gt 0) {
                [TimeSpan]::FromSeconds($elapsed.TotalSeconds / ($totalProgress / 100))
            } else {
                [TimeSpan]::Zero
            }
            $remaining = $estimatedTotal - $elapsed

            Write-Progress -Activity $this.Activity `
                -Status "$status ($totalProgress%)" `
                -PercentComplete $totalProgress `
                -CurrentOperation $stepName `
                -SecondsRemaining $remaining.TotalSeconds

            $this.Logger.Debug("Progress: $totalProgress% - $stepName - $status", "ProgressTracker")
        }
    }

    [void] CompleteStep([string]$stepName) {
        if ($this.StepWeights.ContainsKey($stepName)) {
            $this.CurrentStep += $this.StepWeights[$stepName]
            $totalProgress = [math]::Min(100, [math]::Round(($this.CurrentStep / $this.TotalSteps) * 100, 0))

            Write-Progress -Activity $this.Activity `
                -Status "Completed: $stepName ($totalProgress%)" `
                -PercentComplete $totalProgress `
                -CurrentOperation $stepName

            $this.Logger.Info("Completed: $stepName (Progress: $totalProgress%)", "ProgressTracker")
        }
    }

    [void] Complete() {
        Write-Progress -Activity $this.Activity -Completed
        $elapsed = (Get-Date) - $this.StartTime
        $this.Logger.Info("$($this.Activity) completed in $($elapsed.ToString('mm\:ss'))", "ProgressTracker")
    }

    [void] StartStep([string]$stepName, [string]$status) {
        $this.UpdateProgress($stepName, $status, 0)
    }
}
#endregion

#region Error Handling Framework
class ErrorTracker {
    [int]$FatalErrors = 0
    [int]$RecoverableErrors = 0
    [int]$Warnings = 0
    [System.Collections.ArrayList]$ErrorDetails

    ErrorTracker() {
        $this.ErrorDetails = New-Object System.Collections.ArrayList
    }

    [void] RecordError([string]$component, [string]$message, [bool]$isFatal, [System.Exception]$exception) {
        $errorRecord = @{
            Timestamp = Get-Date
            Component = $component
            Message = $message
            IsFatal = $isFatal
            Exception = $exception.Message
            StackTrace = $exception.StackTrace
        }

        [void]$this.ErrorDetails.Add($errorRecord)

        if ($isFatal) {
            $this.FatalErrors++
        } else {
            $this.RecoverableErrors++
        }
    }

    [void] RecordWarning([string]$component, [string]$message) {
        $this.Warnings++
        $warningRecord = @{
            Timestamp = Get-Date
            Component = $component
            Message = $message
            Type = "Warning"
        }
        [void]$this.ErrorDetails.Add($warningRecord)
    }

    [bool] HasFatalErrors() {
        return $this.FatalErrors -gt 0
    }

    [bool] HasErrors() {
        return ($this.FatalErrors + $this.RecoverableErrors) -gt 0
    }

    [string] GetSummary() {
        return "Fatal: $($this.FatalErrors), Recoverable: $($this.RecoverableErrors), Warnings: $($this.Warnings)"
    }
}
#endregion
