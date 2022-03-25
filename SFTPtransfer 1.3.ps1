# Made by: Kvoo
# Date: 25.03.2022
# Version: 1.3

# Start Log
$VerbosePreference = "Continue"
$dt = get-date -format dd-MM-yyyy-hh-mm
Start-Transcript -Path "$PSScriptRoot\Logs\log-$dt.txt" -IncludeInvocationHeader -NoClobber

Import-Module Transferetto

$SFTPServer = 'sftp.contoso.com'
$SFTPUser = 'sftp_user'
$SFTPPort = '20322'
$SFTPPrivKey = "$PSScriptRoot\privkey.priv"
$LocalDownloadPath = "$PSScriptRoot\Files\Downloaded"
$LocalSentPath = "$PSScriptRoot\Files\Sent"

$MailAdmin = @{
    to = @("sysadmin@contoso.com")
    from = "SFTPTransfer@contoso.com"
    smtpserver = "mail.contoso.com"
    BodyAsHtml = $True
}

# Delete sent files older than 1 month
Write-Output "Deleting sent files older than 1 month..."
$limit = (Get-Date).AddDays(-30)
Get-ChildItem -Path $LocalSentPath -Recurse -Force -Verbose | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force

# Initializes SFTP Session
try {
    Write-Output "Inintializing SFTP Session"
    $SFTPSession = Connect-SFTP -PrivateKey $SFTPPrivKey -Server $SFTPServer -Username $SFTPUser -Port $SFTPPort -ErrorAction Stop
}
catch {
    Send-MailMessage @MailAdmin -Body "SFTP: Error initializing SFTP Session. Script Stopped. Check Log for further details" -Subject "SVFTP01: SFTP Script Error"
    Write-Output $Error
    Exit
}

# Create variable containing all file objects.
$SFTPFileList = Get-SFTPList -SftpClient $SFTPSession -Path /bzc/juris/out
Write-Output "Files to download: SFTPFileList"

# Loops through each path except /sftppath/path/path/. and /sftppath/path/path/.. and downloads the files, if there are any
# Then checks if the file is readable
$LocalFiles = @()
$SFTPFiles = @()
if ($SFTPFileList.Length -ne 2){
    foreach ($File in $SFTPFileList) {
        $SFTPFilePath = $File | Select-Object -ExpandProperty FullName
        $FileName = Split-Path $SFTPFilePath -Leaf
        $LocalFullPath = "$LocalDownloadPath\$FileName"
            if (($SFTPFilePath -ne '/sftppath/path/path/.') -and ($SFTPFilePath -ne '/sftppath/path/path/..')) {
                $SFTPFiles = $SFTPFiles + $SFTPFilePath
                try{
                    Write-Output "Downloading file:"
                    Receive-SFTPFile -SftpClient $SFTPSession -LocalPath $LocalFullPath -RemotePath $SFTPFilePath
                    Write-Output "Downloaded File: "
                    Get-Item "$LocalDownloadPath\$FileName" -ErrorAction Stop
                    $LocalFiles = $LocalFiles + $LocalFullPath
                }
                catch {
                    Send-MailMessage @MailAdmin -Body "SFTP: Error downloading $FileName. Script Stopped. Check Log for further details" -Subject "SVFTP01: SFTP Script Error"
                    Write-Output $Error
                    Exit
                }
            }
    }
}

#E-Mail Var hashtable
$Mail = @{
    to = "recipient@contoso.com"
    cc = "admin@contoso.com"
    from = "Do-NOT-Reply@contoso.com"
    Body = "YOu'll find todays XML-Files attached to this E-Mails. This email was generated automatically."
    subject = "XML Data Transfer"
    smtpserver = "mail.contoso.com"
    BodyAsHtml = $True
    Attachments = $LocalFiles
}

# Send E-Mail and move files to sent path
if ($LocalFiles) {
    try {
        Send-MailMessage @Mail -ErrorAction Stop
        Write-Output "Sent E-Mail"
    }
    catch{
        Send-MailMessage @MailAdmin -Body "SFTP: Error sending E-Mail. Check Log for further details" -Subject "SVFTP01: SFTP Script Error"
        Write-Output $Error
    }
    foreach ($LocalFile in $LocalFiles) {
            try {
                Move-Item -Path $LocalFile -Destination $LocalSentPath -ErrorAction Stop
                Write-Output "Moved $LocalFile to sent path"
            }
            catch {
                Send-MailMessage @MailAdmin -Body "SFTP: Error copying file. Check Log for further details" -Subject "SVFTP01: SFTP Script Error"
                Write-Output $Error
            }
    }
}

# Delete downloaded Files on SFTP Server
if ($SFTPFiles){
    foreach ($File in $SFTPFiles) {
        Remove-SFTPFile -SftpClient $SFTPSession -RemotePath $File
        Write-Output "Deleted $File on SFTP Server"
    }
}

# Close SFTP Session
Disconnect-SFTP -SftpClient $SFTPSession
Write-Output "Disconnected SFTP Session"

Stop-Transcript
