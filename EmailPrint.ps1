Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$ConfigFile = "$PSScriptRoot\mailer_config.json"
$CredFile = "$PSScriptRoot\mailer_secure_string.txt"

# --- Reset Logic ---
if ($args -contains "--reset") {
    if (Test-Path $ConfigFile) { Remove-Item $ConfigFile -Force }
    if (Test-Path $CredFile) { Remove-Item $CredFile -Force }
    [System.Windows.Forms.MessageBox]::Show("All saved settings and passwords have been deleted.", "Reset Complete", "OK", "Information")
    exit
}

# --- Initial Setup ---
if (-not (Test-Path $ConfigFile) -or -not (Test-Path $CredFile)) {
    [System.Windows.Forms.MessageBox]::Show("No saved settings found. Let's set up your printer automation.", "Initial Setup", "OK", "Information")
    
    $sender = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Sender Gmail Address:", "Setup")
    if (-not $sender) { exit }

    $password = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your 16-character App Password:", "Setup")
    if (-not $password) { exit }

    $receiver = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Printer Email Address:", "Setup")
    if (-not $receiver) { exit }

    [System.Windows.Forms.MessageBox]::Show("Select the folder containing the files to print.", "Setup", "OK", "Information")
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit }
    $folder = $folderBrowser.SelectedPath

    $password | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File $CredFile
    
    $config = @{ sender_email = $sender; receiver_email = $receiver; folder_path = $folder }
    $config | ConvertTo-Json | Out-File $ConfigFile
}

# --- Load Config ---
$config = Get-Content $ConfigFile | ConvertFrom-Json
$senderEmail = $config.sender_email
$receiverEmail = $config.receiver_email
$folderPath = $config.folder_path

try {
    $secureString = Get-Content $CredFile | ConvertTo-SecureString
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
    $appPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to decrypt password. Please run with --reset.", "Error", "OK", "Error")
    exit
}

if (-not (Test-Path $folderPath)) {
    [System.Windows.Forms.MessageBox]::Show("The folder '$folderPath' no longer exists.", "Error", "OK", "Error")
    exit
}

# --- Process Files ---
$sentFolder = Join-Path $folderPath "Sent_Files"
if (-not (Test-Path $sentFolder)) { New-Item -ItemType Directory -Path $sentFolder | Out-Null }

$allFiles = Get-ChildItem -Path $folderPath -File
if ($allFiles.Count -eq 0) {
    # Optional: You can comment out this warning if you want it to be completely silent when empty
    # [System.Windows.Forms.MessageBox]::Show("No files found to print.", "Warning", "OK", "Warning")
    exit
}

$batchSize = 10
$totalEmailsSent = 0
$totalFilesAttached = 0
$totalFilesMoved = 0

try {
    $smtp = New-Object System.Net.Mail.SmtpClient("smtp.gmail.com", 587)
    $smtp.EnableSsl = $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential($senderEmail, $appPassword)

    for ($i = 0; $i -lt $allFiles.Count; $i += $batchSize) {
        $batch = $allFiles | Select-Object -Skip $i -First $batchSize
        
        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From = $senderEmail
        $msg.To.Add($receiverEmail)
        
        # --- BLANK SUBJECT AND BODY FOR PRINTER ---
        $msg.Subject = "" 
        $msg.Body = ""
        # ------------------------------------------

        $successfullyAttached = @()

        foreach ($file in $batch) {
            try {
                $attachment = New-Object System.Net.Mail.Attachment($file.FullName)
                $msg.Attachments.Add($attachment)
                $successfullyAttached += $file
                $totalFilesAttached++
            } catch {
                # Skip locked files
            }
        }

        if ($successfullyAttached.Count -gt 0) {
            $smtp.Send($msg)
            $totalEmailsSent++
            
            # Clean up attachments to unlock files
            $msg.Attachments.Dispose()

            # Move Files
            foreach ($file in $successfullyAttached) {
                $baseName = $file.BaseName
                $extension = $file.Extension
                $counter = 1
                $newFileName = $file.Name
                $destPath = Join-Path $sentFolder $newFileName

                while (Test-Path $destPath) {
                    $newFileName = "$baseName ($counter)$extension"
                    $destPath = Join-Path $sentFolder $newFileName
                    $counter++
                }
                Move-Item -Path $file.FullName -Destination $destPath -Force
                $totalFilesMoved++
            }
        }
    }

    $summary = "Transfer Complete!`n`nSender: $senderEmail`nReceiver: $receiverEmail`nTotal Attachments: $totalFilesAttached`nTotal Emails Sent: $totalEmailsSent`nFiles Moved to Archive: $totalFilesMoved"
    [System.Windows.Forms.MessageBox]::Show($summary, "Success", "OK", "Information")

} catch {
    [System.Windows.Forms.MessageBox]::Show("Error sending to printer: $($_.Exception.Message)", "Error", "OK", "Error")
} finally {
    if ($smtp) { $smtp.Dispose() }
}