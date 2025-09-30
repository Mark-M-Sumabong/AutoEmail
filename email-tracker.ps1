<# ===============================
    Email Tracker - Full Script (Old code modernized + .env loader)
    - Loads .env into Process env (optional override)
    - Validates environment variables
    - Reads Excel ("Data" sheet)
    - Groups by recipient where Status == "Pending"
    - Sends Outlook emails; attaches Excel only if computer count >= ATTACH_THRESHOLD
    - Adds a note to the email body when attachment is skipped
    - Tracks follow-ups in Access DB (ACE OLEDB)
    - Logs results and prints a summary
   =============================== #>

[CmdletBinding()]
param(
    # Path to .env file; defaults to "<this script folder>\.env"
    [string]$DotEnvPath = (Join-Path -Path $PSScriptRoot -ChildPath ".env"),
    # If set, values in .env override any existing Process env values
    [switch]$DotEnvOverride
)

# Optional: stricter runtime
# Set-StrictMode -Version Latest

# Ensure HttpUtility is available for HTML encoding (Windows PowerShell)
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# === Helper: Read env var across Process → User → Machine ===
function Get-EnvValue {
    param([Parameter(Mandatory)][string]$Name)
    $val = [System.Environment]::GetEnvironmentVariable($Name, 'Process')
    if ([string]::IsNullOrWhiteSpace($val)) {
        $val = [System.Environment]::GetEnvironmentVariable($Name, 'User')
    }
    if ([string]::IsNullOrWhiteSpace($val)) {
        $val = [System.Environment]::GetEnvironmentVariable($Name, 'Machine')
    }
    return $val
}

# === Helper: Load variables from a .env file into the current Process env ===
function LoadDotEnv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$OverrideExisting
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Host "[INFO] No .env file found at: $Path (skipping)" -ForegroundColor DarkGray
        return
    }

    Write-Host ("[INFO] Loading .env from: {0}" -f $Path) -ForegroundColor Cyan
    $lineNo = 0
    $loaded = 0

    foreach ($rawLine in [System.IO.File]::ReadAllLines($Path)) {
        $lineNo++
        $line = $rawLine.Trim()

        # Skip blanks and comments (# or ;)
        if ($line.Length -eq 0 -or $line.StartsWith('#') -or $line.StartsWith(';')) { continue }

        # Support optional "export KEY=VALUE" prefix
        if ($line -match '^\s*export\s+') {
            $line = $line -replace '^\s*export\s+', ''
        }

        # Parse KEY=VALUE
        if ($line -notmatch '^\s*(?<k>[A-Za-z_][A-Za-z0-9_\.\-]*)\s*=\s*(?<v>.*)\s*$') {
            Write-Host ("[WARN] Skipping invalid line {0}: {1}" -f $lineNo, $rawLine) -ForegroundColor DarkYellow
            continue
        }

        $k = $Matches.k
        $v = $Matches.v

        # Strip a single pair of surrounding quotes if present
        if (($v.StartsWith('"') -and $v.EndsWith('"')) -or ($v.StartsWith("'") -and $v.EndsWith("'"))) {
            $v = $v.Substring(1, $v.Length - 2)
        } else {
            # Remove trailing inline comment if present: KEY=VALUE # comment
            $v = ($v -replace '\s+#.*$', '')
        }

        $v = $v.Trim()

        # Respect precedence unless override requested
        $existing = [System.Environment]::GetEnvironmentVariable($k, 'Process')
        if (-not [string]::IsNullOrWhiteSpace($existing) -and -not $OverrideExisting.IsPresent) {
            continue
        }

        [System.Environment]::SetEnvironmentVariable($k, $v, 'Process')
        $loaded++
    }

    Write-Host ("[OK] Loaded {0} variable(s) from .env" -f $loaded) -ForegroundColor Green
}

# === 0) Load .env (before validation) ===
LoadDotEnv -Path $DotEnvPath -OverrideExisting:$DotEnvOverride

# === 1) Validate Required Environment Variables (Fail Fast) ===
$requiredVars = @(
    "EXCEL_PATH",
    "LOG_FILE",
    "ACCESS_DB_PATH",
    "EMAIL_TEMPLATE_PATH",
    "DEFAULT_SENDER",
    "ATTACH_THRESHOLD",  # optional; we still validate and default later
    "EMAIL_DEADLINE"
)

$missingVars = @()
foreach ($var in $requiredVars) {
    $val = Get-EnvValue -Name $var
    if ([string]::IsNullOrWhiteSpace($val)) {
        $missingVars += $var
    }
}

if ($missingVars.Count -gt 0) {
    Write-Host "`n❌ Missing environment variables:" -ForegroundColor Red
    $missingVars | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    Write-Host "`nSet them with 'setx VAR_NAME VALUE', reopen PowerShell, or provide a .env file." -ForegroundColor Cyan
    Write-Host "Tip: .env path currently: $DotEnvPath"
    exit 1
} else {
    Write-Host "`n✅ All required environment variables are set." -ForegroundColor Green
    foreach ($var in $requiredVars) {
        $val = Get-EnvValue -Name $var
        Write-Host ("{0}: {1}" -f $var, $val)
    }
    Write-Host "Note: ATTACH_THRESHOLD will default to 20 if not a valid number."
}

# === 2) Functions (your originals retained, adapted for env vars) ===

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

    $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$accessDbPath;"
    $connection = New-Object -TypeName System.Data.OleDb.OleDbConnection -ArgumentList $connectionString
    $connection.Open()

    $query = "SELECT * FROM EmailTracking WHERE EmailAddress = ? AND Subject = ?"
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $command.Parameters.Add((New-Object Data.OleDb.OleDbParameter('EmailAddress', [System.Data.OleDb.OleDbType]::VarWChar))).Value = $emailAddress
    $command.Parameters.Add((New-Object Data.OleDb.OleDbParameter('Subject', [System.Data.OleDb.OleDbType]::VarWChar))).Value = $subject

    $reader = $command.ExecuteReader()

    if ($reader.HasRows) {
        $reader.Read()
        $followUpCount = $reader["FollowUpCount"]
        $newFollowUpCount = $followUpCount + 1
        $lastFollowUpDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        $updateQuery = "UPDATE EmailTracking SET FollowUpCount = ?, LastFollowUpDate = ? WHERE ID = ?"
        $updateCommand = $connection.CreateCommand()
        $updateCommand.CommandText = $updateQuery
        $updateCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('FollowUpCount', [System.Data.OleDb.OleDbType]::Integer))).Value = $newFollowUpCount
        $updateCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('LastFollowUpDate', [System.Data.OleDb.OleDbType]::Date))).Value = $lastFollowUpDate
        $updateCommand.Parameters.Add((New-Object Data.OleDb.OleDbParameter('ID', [System.Data.OleDb.OleDbType]::Integer))).Value = $reader["ID"]

        $updateCommand.ExecuteNonQuery()
        Write-Host "Follow-up #$newFollowUpCount updated for $emailAddress"
    } else {
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

    CleanupAccess -connection $connection
}

# === 3) Load Environment Variables into local vars ===
$excelPath        = Get-EnvValue -Name 'EXCEL_PATH'
$logFile          = Get-EnvValue -Name 'LOG_FILE'
$accessDbPath     = Get-EnvValue -Name 'ACCESS_DB_PATH'
$templatePath     = Get-EnvValue -Name 'EMAIL_TEMPLATE_PATH'
$defaultSender    = Get-EnvValue -Name 'DEFAULT_SENDER'
$deadline         = Get-EnvValue -Name 'EMAIL_DEADLINE'  # optional

# Parse ATTACH_THRESHOLD robustly (default to 20 if missing/invalid or < 1)
[int]$attachThreshold = 0
$attachRaw = Get-EnvValue -Name 'ATTACH_THRESHOLD'
if (-not [int]::TryParse($attachRaw, [ref]$attachThreshold) -or $attachThreshold -lt 1) {
    $attachThreshold = 20
}

# === 4) Validate Files Exist ===
CheckFileExistence -filePath $excelPath
CheckFileExistence -filePath $templatePath
CheckFileExistence -filePath $accessDbPath

# Initialize log file (create dir if needed)
try {
    $logDir = Split-Path -Path $logFile -Parent
    if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    if (!(Test-Path $logFile)) {
        "Timestamp,Email,Computers,Status,ErrorMessage" | Out-File -FilePath $logFile -Encoding UTF8
    }
} catch {
    Write-Host "Error initializing log file: $_"
    exit
}

# Load HTML template
$htmlTemplate = Get-Content -Path $templatePath -Raw

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

# === Build email → [computers, CC, AppNames] map (only Pending) ===
$emailMap = @{ }
for ($row = 2; $row -le $lastRow; $row++) {
    try {
        $email    = $sheet.Cells.Item($row, 5).Text   # Column "To"
        $computer = $sheet.Cells.Item($row, 1).Text   # Column "Computer Name"
        $appName  = $sheet.Cells.Item($row, 4).Text   # Column "App Name"
        $status   = $sheet.Cells.Item($row, 8).Text   # Column "Status"
        $ccValue  = $sheet.Cells.Item($row, 6).Text   # Column "CC"

        if ($status.Trim() -eq "Pending" -and -not [string]::IsNullOrWhiteSpace($email)) {
            if (-not $emailMap.ContainsKey($email)) {
                $emailMap[$email] = @{
                    Computers = New-Object System.Collections.ArrayList
                    AppNames  = New-Object 'System.Collections.Generic.HashSet[string]'
                    CC        = New-Object 'System.Collections.Generic.HashSet[string]'
                }
            }
            [void]$emailMap[$email].Computers.Add($computer)
            if (-not [string]::IsNullOrWhiteSpace($appName)) {
                [void]$emailMap[$email].AppNames.Add($appName.Trim())
            }
            if (-not [string]::IsNullOrWhiteSpace($ccValue)) {
                # If CC has multiple addresses separated by ';', split them:
                # $ccValue.Split(';') | ForEach-Object { if (-not [string]::IsNullOrWhiteSpace($_)) { [void]$emailMap[$email].CC.Add($_.Trim()) } }
                [void]$emailMap[$email].CC.Add($ccValue.Trim())
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
$sentCount = 0
$failedCount = 0
$withAttachment = 0
$withoutAttachment = 0

foreach ($entry in $emailMap.GetEnumerator()) {
    $email      = $entry.Key
    $computers  = $entry.Value.Computers | Sort-Object | Get-Unique
    $computerCount = $computers.Count
    $compList   = $computers -join "; "

    $appsSorted    = $entry.Value.AppNames | Sort-Object
    $subjectApps   = ($appsSorted -join ", ")
    if ($subjectApps.Length -gt 80) {
        $subjectApps = ((($appsSorted | Select-Object -First 3) -join ", ") + ", etc.")
    }
    if ([string]::IsNullOrWhiteSpace($subjectApps)) {
        $subjectApps = "Application"
    }
    $subject = "Follow-Up: $subjectApps Assessment on Your Server(s)"

    # Build rows for the template (use REAL HTML tags)
    $rowsHtml = (
        $computers | ForEach-Object { "<tr><td>$($_)</td></tr>" }
    ) -join "`n"

    # Compose HTML body
    $htmlBody = $htmlTemplate.
        Replace("{{AppNames}}", [System.Web.HttpUtility]::HtmlEncode($subjectApps)).
        Replace("{{RowsHtml}}", $rowsHtml)

    # Optional: inject deadline if your template supports {{Deadline}}
    if (-not [string]::IsNullOrWhiteSpace($deadline)) {
        $htmlBody = $htmlBody.Replace("{{Deadline}}", [System.Web.HttpUtility]::HtmlEncode($deadline))
    }

    $ccList = ($entry.Value.CC | Sort-Object | Get-Unique) -join "; "

    # Attachment policy + note in body if skipped
    $attachmentStatus = "Included"
    if ($computerCount -lt $attachThreshold) {
        $htmlBody += "<p><em>Note: Attachment omitted because the number of servers ($computerCount) is below the threshold ($attachThreshold).</em></p>"
        $attachmentStatus = "Skipped"
    }

    try {
        $mail = $outlook.CreateItem(0)
        $mail.SendUsingAccount = $account
        $mail.To       = $email
        if (-not [string]::IsNullOrWhiteSpace($ccList)) {
            $mail.CC = $ccList
        }
        $mail.Subject  = $subject
        $mail.HTMLBody = $htmlBody

        if ($attachmentStatus -eq "Included") {
            $mail.Attachments.Add($excelPath) | Out-Null
            $withAttachment++
        } else {
            $withoutAttachment++
        }

        $mail.Send()

        TrackEmailFollowUp -emailAddress $email -subject $subject
        Write-Log -Email $email -Computers $compList -Status "Sent (Attachment: $attachmentStatus)"
        $sentCount++
        Start-Sleep -Seconds 2
    } catch {
        $err = $_.Exception.Message
        Write-Log -Email $email -Computers $compList -Status "Failed" -ErrorMessage $err
        $failedCount++
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

# === Summary ===
Write-Host "`n✅ All done! Log saved to $logFile"
Write-Host "Summary:"
Write-Host "  Emails Sent: $sentCount"
Write-Host "  Failed: $failedCount"
Write-Host "  With Attachment: $withAttachment"
Write-Host "  Without Attachment: $withoutAttachment"