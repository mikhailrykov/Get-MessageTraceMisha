<#
.SYNOPSIS
    Get-MessageTraceMisha - Batch Message Trace Query for Exchange Online
    
.DESCRIPTION
    Processes a table of MessageIDs through Get-MessageTraceV2 with throttling and error handling.
    Automatically handles rate-limit errors by pausing and retrying failed queries.
    Supports custom date ranges for trace queries.
    
.PARAMETER InputTable
    Mandatory. PowerShell object or table containing MessageID column.
    
.PARAMETER MessageIDColumn
    The name of the column containing MessageIDs. Default: "MessageID"
    
.PARAMETER StartDate
    Start date for message trace query. Default: 10 days ago
    
.PARAMETER EndDate
    End date for message trace query. Default: Now. Must be after StartDate.
    
.PARAMETER OutputPath
    Path to save the resulting XML file. Default: System temp folder
    
.PARAMETER OutputFileName
    Name for the output XML file. Default: "MessageTrace_[timestamp].xml"
    
.PARAMETER PauseMinutes
    Number of minutes to pause when throttling error occurs. Default: 10
    
.PARAMETER SendReport
    Switch to send the report via email
    
.PARAMETER ReportRecipient
    Email address to send the report to (required if SendReport is used)
    
.PARAMETER SMTPServer
    SMTP server address. Default: "appsmtp.ottawa.ca"
    
.EXAMPLE
    $messageIds = Import-Csv "messages.csv"
    Get-MessageTraceMisha -InputTable $messageIds -MessageIDColumn "MessageID" -OutputPath "C:\Reports" -Verbose
    
.EXAMPLE
    Get-MessageTraceMisha -InputTable $messageIds -StartDate (Get-Date).AddDays(-5) -EndDate (Get-Date) -SendReport -ReportRecipient "admin@company.com" -PauseMinutes 15 -Verbose
    
.NOTES
    Author: Mikhail Rykov
    Requires: Exchange Online PowerShell Module
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [object[]]$InputTable,
    
    [Parameter(Mandatory=$false)]
    [string]$MessageIDColumn = "MessageID",
    
    [Parameter(Mandatory=$false)]
    [datetime]$StartDate = (Get-Date).AddDays(-10),
    
    [Parameter(Mandatory=$false)]
    [datetime]$EndDate = (Get-Date),
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = $env:TEMP,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFileName = "MessageTrace_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml",
    
    [Parameter(Mandatory=$false)]
    [int]$PauseMinutes = 10,
    
    [Parameter(Mandatory=$false)]
    [switch]$SendReport,
    
    [Parameter(Mandatory=$false)]
    [string]$ReportRecipient,
    
    [Parameter(Mandatory=$false)]
    [string]$SMTPServer = "appsmtp.ottawa.ca"
)

# Initialize variables
$results = @()
$failedMessageIds = @()
$successCount = 0
$failureCount = 0
$totalCount = 0
$retryAttempts = @{}
$maxRetries = 3
$scriptStartTime = Get-Date

Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Script execution started"
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Start Date: $($StartDate.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] End Date: $($EndDate.ToString('yyyy-MM-dd HH:mm:ss'))"

# Validate date parameters
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Validating date parameters..."
if ($EndDate -le $StartDate) {
    Write-Error "EndDate ($($EndDate.ToString('yyyy-MM-dd HH:mm:ss'))) must be after StartDate ($($StartDate.ToString('yyyy-MM-dd HH:mm:ss')))"
    return
}
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Date validation passed"

# Get total message count
$totalCount = @($InputTable).Count
Write-Host "Processing $totalCount message(s) for trace..." -ForegroundColor Cyan
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Total messages to process: $totalCount"

# Create output file path
$fullOutputPath = Join-Path -Path $OutputPath -ChildPath $OutputFileName
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Output path: $fullOutputPath"

# Validate input table has MessageID column
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Validating input table structure..."
try {
    $testProperty = $InputTable[0].$MessageIDColumn
    if ($null -eq $testProperty) {
        throw "Column '$MessageIDColumn' not found in input table"
    }
    Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Input table validation passed"
}
catch {
    Write-Error "Error validating input table: $_"
    return
}

# Process each MessageID
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Starting message processing loop..."
for ($i = 0; $i -lt $totalCount; $i++) {
    $rawMessageId = $InputTable[$i].$MessageIDColumn
    
    # Remove angle brackets if present
    $messageId = $rawMessageId -replace '^<|>$', ''
    Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$i] Raw MessageID: '$rawMessageId' | Cleaned MessageID: '$messageId'"
    
    if ($rawMessageId -ne $messageId) {
        Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$i] Angle brackets removed from MessageID"
    }
    
    $currentNumber = $i + 1
    
    # Display progress
    Write-Progress -Activity "Processing Message Traces" `
                   -Status "Message $currentNumber of $totalCount" `
                   -PercentComplete (($currentNumber / $totalCount) * 100) `
                   -CurrentOperation "MessageID: $messageId" `
                   -Id 1
    
    Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$currentNumber/$totalCount] Processing MessageID: $messageId"
    
    # Initialize retry count for this message if not exists
    if (-not $retryAttempts.ContainsKey($messageId)) {
        $retryAttempts[$messageId] = 0
    }
    
    $querySuccessful = $false
    
    while ($retryAttempts[$messageId] -le $maxRetries -and -not $querySuccessful) {
        try {
            Write-Host "[$currentNumber/$totalCount] Querying MessageID: $messageId" -ForegroundColor White
            Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$currentNumber/$totalCount] Executing Get-MessageTraceV2 for MessageID: $messageId (Attempt: $($retryAttempts[$messageId] + 1))"
            
            # Execute Get-MessageTraceV2 with date range
            $trace = Get-MessageTraceV2 -MessageId $messageId `
                                       -StartDate $StartDate `
                                       -EndDate $EndDate `
                                       -ErrorAction Stop
            
            if ($trace) {
                $results += $trace
                $successCount++
                $querySuccessful = $true
                Write-Host "  ✓ Successfully retrieved trace" -ForegroundColor Green
                Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$currentNumber/$totalCount] Successfully retrieved trace for MessageID: $messageId. Total results: $($results.Count)"
            }
            else {
                Write-Warning "  ⚠ No trace data returned for MessageID: $messageId"
                Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$currentNumber/$totalCount] No trace data found for MessageID: $messageId in the specified date range"
                $failureCount++
                $querySuccessful = $true
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$currentNumber/$totalCount] Error occurred: $errorMessage"
            
            # Check for throttling error
            if ($errorMessage -like "*permitted limit*" -or $errorMessage -like "*throttled*") {
                Write-Host "  ⚠ Rate limit exceeded!" -ForegroundColor Yellow
                Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$currentNumber/$totalCount] Throttling error detected. Initiating pause for $PauseMinutes minute(s)"
                Write-Host "  Pausing for $PauseMinutes minute(s) before retry..." -ForegroundColor Yellow
                
                # Countdown timer
                $totalSeconds = $PauseMinutes * 60
                for ($countdown = $totalSeconds; $countdown -gt 0; $countdown--) {
                    $minutes = [math]::Floor($countdown / 60)
                    $seconds = $countdown % 60
                    $percentComplete = (($totalSeconds - $countdown) / $totalSeconds) * 100
                    
                    Write-Progress -Activity "Rate Limit Pause" `
                                   -Status "Waiting to resume queries" `
                                   -SecondsRemaining $countdown `
                                   -PercentComplete $percentComplete `
                                   -CurrentOperation "Time remaining: ${minutes}m ${seconds}s" `
                                   -Id 2
                    Start-Sleep -Seconds 1
                }
                Write-Progress -Activity "Rate Limit Pause" -Completed -Id 2
                Write-Host "  ✓ Resuming queries..." -ForegroundColor Green
                Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$currentNumber/$totalCount] Pause completed. Resuming queries"
                $retryAttempts[$messageId]++
            }
            else {
                # Non-throttling error
                Write-Host "  ✗ Error querying MessageID $messageId : $errorMessage" -ForegroundColor Red
                Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$currentNumber/$totalCount] Non-throttling error for MessageID: $messageId. Adding to failed list"
                $failedMessageIds += $messageId
                $failureCount++
                $querySuccessful = $true
            }
        }
    }
    
    if ($retryAttempts[$messageId] -gt $maxRetries) {
        Write-Host "  ✗ Max retries exceeded for MessageID: $messageId" -ForegroundColor Red
        Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] [$currentNumber/$totalCount] Max retries ($maxRetries) exceeded for MessageID: $messageId"
        if ($messageId -notin $failedMessageIds) {
            $failedMessageIds += $messageId
            $failureCount++
        }
    }
}

Write-Progress -Activity "Processing Message Traces" -Completed -Id 1
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Message processing loop completed"

# Export results to XML
Write-Host "`nExporting results to XML..." -ForegroundColor Cyan
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Starting XML export process"

try {
    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Creating output directory: $OutputPath"
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Building XML document with $($results.Count) records"
    
    # Create XML structure
    $xmlDocument = New-Object System.Xml.XmlDocument
    $xmlDeclaration = $xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", $null)
    $xmlDocument.AppendChild($xmlDeclaration) | Out-Null
    
    $rootElement = $xmlDocument.CreateElement("MessageTraces")
    $rootElement.SetAttribute("ExportDate", (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
    $rootElement.SetAttribute("StartDate", $StartDate.ToString("yyyy-MM-dd HH:mm:ss"))
    $rootElement.SetAttribute("EndDate", $EndDate.ToString("yyyy-MM-dd HH:mm:ss"))
    $rootElement.SetAttribute("TotalRecords", $results.Count)
    $rootElement.SetAttribute("SuccessCount", $successCount)
    $rootElement.SetAttribute("FailureCount", $failureCount)
    $xmlDocument.AppendChild($rootElement) | Out-Null
    
    # Add each result to XML
    Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Adding $($results.Count) records to XML document"
    foreach ($result in $results) {
        $recordElement = $xmlDocument.CreateElement("MessageTrace")
        
        foreach ($property in $result.PSObject.Properties) {
            $propElement = $xmlDocument.CreateElement($property.Name)
            $propElement.InnerText = $property.Value
            $recordElement.AppendChild($propElement) | Out-Null
        }
        
        $rootElement.AppendChild($recordElement) | Out-Null
    }
    
    # Save XML file
    Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Saving XML file to: $fullOutputPath"
    $xmlDocument.Save($fullOutputPath)
    Write-Host "✓ XML file saved: $fullOutputPath" -ForegroundColor Green
    Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] XML file successfully saved"
}
catch {
    Write-Error "Error exporting to XML: $_"
    Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] XML export failed: $_"
    return
}

# Display summary
Write-Host "`n" + ("="*60) -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host ("="*60) -ForegroundColor Cyan
Write-Host "Script Execution Time: $((Get-Date) - $scriptStartTime)" -ForegroundColor White
Write-Host "Query Date Range: $($StartDate.ToString('yyyy-MM-dd HH:mm:ss')) to $($EndDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
Write-Host "Total MessageIDs Processed: $totalCount" -ForegroundColor White
Write-Host "Successful Traces: $successCount" -ForegroundColor Green
Write-Host "Failed Queries: $failureCount" -ForegroundColor Yellow
Write-Host "Resulting Table Rows: $($results.Count)" -ForegroundColor Cyan
Write-Host "Output File: $fullOutputPath" -ForegroundColor White

if ($failedMessageIds.Count -gt 0) {
    Write-Host "`nFailed MessageIDs:" -ForegroundColor Yellow
    $failedMessageIds | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}

Write-Host ("="*60) -ForegroundColor Cyan

Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Displaying summary statistics"
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Script execution time: $((Get-Date) - $scriptStartTime)"

# Send email report if requested
if ($SendReport) {
    if (-not $ReportRecipient) {
        Write-Error "ReportRecipient parameter is required when using SendReport switch"
        Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] SendReport requested but ReportRecipient not provided"
        return
    }
    
    Write-Host "`nSending email report..." -ForegroundColor Cyan
    Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Starting email report process to: $ReportRecipient"
    
    try {
        $emailParams = @{
            To = $ReportRecipient
            Subject = "Message Trace Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            SmtpServer = $SMTPServer
            From = "$([Environment]::UserName)@ottawa.ca"
            BodyAsHtml = $false
            Attachments = @($fullOutputPath)
        }
        
        Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Email parameters configured: From=$($emailParams['From']), To=$($emailParams['To']), SmtpServer=$SMTPServer"
        
        $emailBody = @"
Message Trace Query Report
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

SUMMARY STATISTICS:
- Total MessageIDs Processed: $totalCount
- Successful Traces: $successCount
- Failed Queries: $failureCount
- Resulting Rows in Output: $($results.Count)
- Script Execution Time: $((Get-Date) - $scriptStartTime)

QUERY PARAMETERS:
- Start Date: $($StartDate.ToString('yyyy-MM-dd HH:mm:ss'))
- End Date: $($EndDate.ToString('yyyy-MM-dd HH:mm:ss'))
- Pause Duration on Throttle: $PauseMinutes minutes
- SMTP Server: $SMTPServer

OUTPUT FILE:
- Location: $fullOutputPath
- Format: XML

$(if ($failedMessageIds.Count -gt 0) { 
    "FAILED MessageIDs:`r`n$(($failedMessageIds | ForEach-Object { "  - $_" }) -join "`r`n")"
} else {
    "All queries completed successfully."
})

---
This report was generated by Get-MessageTraceMisha script.
"@
        
        $emailParams['Body'] = $emailBody
        
        Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Sending email via $SMTPServer"
        Send-MailMessage @emailParams
        Write-Host "✓ Email report sent to $ReportRecipient" -ForegroundColor Green
        Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Email report successfully sent to $ReportRecipient"
    }
    catch {
        Write-Error "Error sending email report: $_"
        Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Email send failed: $_"
    }
}

Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Script execution completed successfully"
Write-Verbose "[$(Get-Date -Format 'HH:mm:ss')] Total execution time: $((Get-Date) - $scriptStartTime)"

# Return the results
return $results
