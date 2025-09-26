# === Function to Write Logs ===
function Write-Log {
    param(
        [string] $Email,
        [string] $Computers,
        [string] $Status,
        [string] $ErrorMessage = ""
    )
    try {
        $timestamp = (Get-Date).ToString("s")
        $line = "$timestamp,$Email,""$Computers"",$Status,""$ErrorMessage"""
        $line | Out-File -FilePath $logFile -Append -Encoding UTF8
        Write-Host "[$timestamp] [$Status] – $Email – ($Computers) $ErrorMessage"
    } catch {
        Write-Host "Error writing to log file: $_"
    }
}

# === Function to check if a file exists ===
function CheckFileExistence {
    param ([string]$filePath)
    if (-not (Test-Path $filePath)) {
        throw "Error: The file '$filePath' does not exist."
    }
}

# === Function to handle cleanup in case of an error ===
function CleanupAccess {
    param ($connection)
    try {
        $connection.Close()
    } catch {
        Write-Warning "Failed to close the database connection."
    }
}

# === Function to track email follow-ups in Access Database ===
function TrackEmailFollowUp {
    param (
        [string]$emailAddress,
        [string]$subject
    )

    # Define the Access database file path (change to your actual file path)
    $accessDbPath = "C:\Users\mark.m.s.sumabong\Desktop\PowerShell\email_tracking.accdb"

    # Connect to the Access database using ADO.NET
    $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$accessDbPath;"
    $connection = New-Object -TypeName System.Data.OleDb.OleDbConnection -ArgumentList $connectionString
    $connection.Open()

    # Check if the email address and subject already exist in the table
    $query = "SELECT * FROM EmailTracking WHERE EmailAddress = ? AND Subject = ?"
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $command.Parameters.Add((New-Object Data.OleDb.OleDbParameter('EmailAddress', [System.Data.OleDb.OleDbType]::VarWChar))).Value = $emailAddress
    $command.Parameters.Add((New-Object Data.OleDb.OleDbParameter('Subject', [System.Data.OleDb.OleDbType]::VarWChar))).Value = $subject

    $reader = $command.ExecuteReader()

    if ($reader.HasRows) {
        # If email exists, update follow-up count and date
        $reader.Read()
        $followUpCount = $reader["FollowUpCount"]
        $newFollowUpCount = $followUpCount + 1
        $lastFollowUpDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        # Update follow-up count and last follow-up date
        $updateQuery = "UPDATE EmailTracking SET FollowUpCount = ?, LastFollowUpDate = ? WHERE ID = ?"
        $updateCommand = $connection.CreateCommand()
        $updateCommand.CommandText = $updateQuery
        $updateCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('FollowUpCount', [System.Data.OleDb.OleDbType]::Integer))).Value = $newFollowUpCount
        $updateCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('LastFollowUpDate', [System.Data.OleDb.OleDbType]::Date))).Value = $lastFollowUpDate
        $updateCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('ID', [System.Data.OleDb.OleDbType]::Integer))).Value = $reader["ID"]

        $updateCommand.ExecuteNonQuery()
        Write-Host "Follow-up #$newFollowUpCount updated for $emailAddress"
    } else {
        # If email doesn't exist, insert a new record
        $createdAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $insertQuery = "INSERT INTO EmailTracking (EmailAddress, Subject, Status, FollowUpCount, LastFollowUpDate, CreatedAt) VALUES (?, ?, 'Sent', 1, ?, ?)"
        $insertCommand = $connection.CreateCommand()
        $insertCommand.CommandText = $insertQuery
        $insertCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('EmailAddress', [System.Data.OleDb.OleDbType]::VarWChar))).Value = $emailAddress
        $insertCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('Subject', [System.Data.OleDb.OleDbType]::VarWChar))).Value = $subject
        $insertCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('LastFollowUpDate', [System.Data.OleDb.OleDbType]::Date))).Value = $createdAt
        $insertCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('CreatedAt', [System.Data.OleDb.OleDbType]::Date))).Value = $createdAt

        $insertCommand.ExecuteNonQuery()
        Write-Host "New email entry added for $emailAddress"
    }

    # Clean up
    CleanupAccess -connection $connection
}

# === Configuration ===
$excelPath     = [System.Environment]::GetEnvironmentVariable('EXCEL_PATH')
$logFile       = [System.Environment]::GetEnvironmentVariable('LOG_FILE')
$defaultSender = [System.Environment]::GetEnvironmentVariable('DEFAULT_SENDER')
$templatePath  = "C:\Users\mark.m.s.sumabong\Documents\GitHub\AutoEmail\EmailTemplate.html"

# Validate if necessary environment variables are set
if (-not $excelPath -or -not $logFile -or -not $defaultSender) {
    Write-Host "Error: One or more environment variables are not set."
    exit
}



# Validate template file
if (-not (Test-Path $templatePath)) {
    Write-Host "Error: Email template file not found at $templatePath"
    exit
}
$htmlTemplate = Get-Content -Path $templatePath -Raw





# === Prepare Log File ===
try {
    if (!(Test-Path $logFile)) {
        "Timestamp,Email,Computers,Status,ErrorMessage" |
          Out-File -FilePath $logFile -Encoding UTF8
    }
} catch {
    Write-Host "Error initializing log file: $_"
    exit
}

# === Launch Excel ===
$excel = $null
$workbook = $null
$sheet = $null
try {
    $excel    = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($excelPath)
    $sheet    = $workbook.Sheets.Item("Data")
    $lastRow  = $sheet.UsedRange.Rows.Count
} catch {
    Write-Host "Error opening Excel file: $_"
    exit
}

# === Build email → [computers] map ===
$emailMap = @{ }
for ($row = 2; $row -le $lastRow; $row++) {
    try {
            $email    = $sheet.Cells.Item($row, 5).Text   # Column "To"
            $computer = $sheet.Cells.Item($row, 1).Text   # Column "Computer Name"
            $appName  = $sheet.Cells.Item($row, 4).Text   # Column "App Name"
            $status   = $sheet.Cells.Item($row, 8).Text   # Column "Status"

            # Normalize status text and check for exact match
            if ($status.Trim() -eq "Pending" -and -not [string]::IsNullOrWhiteSpace($email)) {
                if (-not $emailMap.ContainsKey($email)) {
                    $emailMap[$email] = @{
                        Computers = New-Object System.Collections.ArrayList
                        AppNames  = New-Object 'System.Collections.Generic.HashSet[string]'
                    }
                }
                [void]$emailMap[$email].Computers.Add($computer)
                if (-not [string]::IsNullOrWhiteSpace($appName)) {
                    [void]$emailMap[$email].AppNames.Add($appName.Trim())
                }
            }
        } catch {
            Write-Host "Error processing row $row in Excel: $_"
        }
    }


# === Launch Outlook & Pick Sender Account ===
$outlook = $null
$namespace = $null
$account = $null
try {
    $outlook   = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $account   = $namespace.Accounts | Where-Object { $_.SmtpAddress -ieq $defaultSender }
    if (-not $account) {
        Throw "Default sender account '$defaultSender' not found in your Outlook profile."
    }
} catch {
    Write-Host "Error initializing Outlook: $_"
    exit
}

# === Send one email per recipient ===

foreach ($entry in $emailMap.GetEnumerator()) {
    $email      = $entry.Key
    $computers  = $entry.Value.Computers | Sort-Object | Get-Unique
    $compList   = $computers -join "; "

    # Build subject with app names
    $appsSorted    = $entry.Value.AppNames | Sort-Object
    $subjectApps   = ($appsSorted -join ", ")
    if ($subjectApps.Length -gt 80) {
        $subjectApps = ((($appsSorted | Select-Object -First 3) -join ", ") + ", etc.")
    }
    if ([string]::IsNullOrWhiteSpace($subjectApps)) {
        $subjectApps = "Application"
    }
    $subject = "Follow-Up: $subjectApps Assessment on Your Server(s)"

    # Build HTML rows
    $rowsHtml = (
        $computers | ForEach-Object { "<tr><td>$($_)</td></tr>" }
    ) -join "`n"

    # Inject into template (no regex; safe and simple)
    $htmlBody = $htmlTemplate.Replace("{{AppNames}}", [System.Web.HttpUtility]::HtmlEncode($subjectApps)).Replace("{{RowsHtml}}", $rowsHtml)

    try {
        # Create and send
        $mail = $outlook.CreateItem(0)
        $mail.SendUsingAccount = $account
        $mail.Attachments.Add($excelPath)
        $mail.To       = $email
        $mail.Subject  = $subject   # ✅ Use dynamic subject
        $mail.HTMLBody = $htmlBody
        $mail.Send()

        # Track email sent status in Access
        TrackEmailFollowUp -emailAddress $email -subject $subject  # ✅ Use dynamic subject
        Write-Log -Email $email -Computers $compList -Status "Sent"
        Start-Sleep -Seconds 2

    } catch {
        $err = $_.Exception.Message
        Write-Log -Email $email -Computers $compList -Status "Failed" -ErrorMessage $err
    }
}

# === Cleanup ===
try {
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet)    | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)    | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} catch {
    Write-Host "Error during cleanup: $_"
}

Write-Host "`n✅ All done! Log saved to $logFile"

