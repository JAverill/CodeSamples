## Mass Application Deployment
## Version 2.0 - Author: John Averill, IBM

## Load Assemblies.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Net
Add-Type -AssemblyName System.Management.Automation

## Ensure PSExec is Present Locally.
$Exec = Test-Path "C:\Windows\System32\PSTools"
IF($Exec -eq $false){RoboCopy "Client Server Location\pstools" "C:\Windows\System32\PSTools" /mir /mt:32}

## Ensure a Temp Directory Exists for Use.
$temp = Test-Path "C:\Temp"
IF(!$temp){New-Item -Path "C:\" -ItemType Directory -Name "Temp" -Force}

## Declare Global Functions.
Function Read-InputBoxDialog([string]$Message, [string]$WindowTitle)
    {return [Microsoft.VisualBasic.Interaction]::InputBox($Message, $WindowTitle)}

Function Message-Box([string]$Message,[string]$WindowTitle,[string]$ButtonStyle)
    {return [System.Windows.MessageBox]::Show($Message,$WindowTitle,$ButtonStyle)}

$ErrorActionPreference = "silentlycontinue"

## Create Form (GUI).
$installform = New-Object System.Windows.Forms.Form
$installform.Text = 'Available Applications.'
$installform.Size = New-Object System.Drawing.Size(520,700)
$installform.StartPosition = 'CenterScreen'

## Display available Software on backup share.
$BackupListBox = New-Object System.Windows.Forms.ListBox
$BackupListBox.Location = New-Object System.Drawing.Point(10,60)
$BackupListBox.Size = New-Object System.Drawing.Size(340,600)
$BackupListBox.Height = 700

## Get Software Share Folders, and show only folders that have applications that can be installed remotely.
$directory = "Client Server Location\software"
CD $directory
$allfolders = dir | Select-Object -ExpandProperty Name | Sort-Object
$outputapps = @()
ForEach ($folder in $allfolders){
    [Void]$BackupListBox.Items.Add($folder)
}
$BackupListBox.add_SelectedIndexChanged($SelectedFile)
$installform.Controls.Add($BackupListBox)

## Hosts Label.
$hostlabel = New-Object System.Windows.Forms.Label
$hostlabel.Location = New-Object System.Drawing.Point(10,10)
$hostlabel.Size = New-Object System.Drawing.Size(210,20)
$hostlabel.Text = "Enter Path to text file containing hosts:"
$installform.controls.add($hostlabel)

## Hostname File location input.
$machineinput = New-Object System.Windows.Forms.TextBox
$machineinput.Location = New-Object System.Drawing.Point(220,10)
$machineinput.Size = New-Object System.Drawing.Size(210,20)
$machineinput.TextAlign = 'Center'
$machineinput.BackColor = 'Red'
$machineinput.ReadOnly = $false
$installform.Controls.Add($machineinput)

## Install Button.
$InstallButton = New-Object System.Windows.Forms.Button
$InstallButton.Location = New-Object System.Drawing.Point(360,100)
$InstallButton.Size = New-Object System.Drawing.Size(150,46)
$InstallButton.Text = "Install Application"
$InstallButton.BackColor = 'Red'
$InstallButton.Add_Click({

    ## Set Hostnames variable.
    $hostsfile = $machineinput.text
    $installform.refresh
    $hosts = Get-Content "$hostsfile"

    ## Determine Selected Item.
    $appname = $BackupListBox.SelectedItem

    ## Create Arrays to classify online machines vs. offline machines.
    $onlinemachines = @()
    $offlinemachines = @()

    ## Create Text Array to create .bat file which will create central logging and execute deployment.
    $array = @()
    $array += '@echo off'
    $array += 'MKDIR "Client Server Location\automation\MassDeployment\Files\Logs\%COMPUTERNAME%"'
    $array += 'echo %COMPUTERNAME%' + " " + "$appname from $directory" + " " + '%DATE% %TIME% >> "Client Server Location\automation\MassDeployment\Files\Logs\%COMPUTERNAME%\DeploymentStarted.txt"'
    $array += "C:\Flags\Temp\$appname\Deploy-Application.exe"
    $array += 'echo %COMPUTERNAME%' + " " + "$appname from $directory" + " " + '%DATE% %TIME% >> "Client Server Location\automation\MassDeployment\Files\Logs\%COMPUTERNAME%\DeploymentCompleted.txt"'
    $array += 'exit'

    ## Create .bat file to install application.
    $array | Out-File "C:\Temp\install$appname.bat" -Encoding ASCII -Force

    ## Determine which machines on list are currently available.
    ForEach($machine in $hosts){
        $online = Test-Connection $machine -Count 1 -Quiet
        IF($online -eq $True){$onlinemachines += $machine}
        ELSE{$offlinemachines += $machine}}
    
    ## Output online list to user for clarification.
    Message-Box -Message "The Following Machines were found online:`r`n$(($onlinemachines) -join ',')`r`nCopying installation script to those machines..." -WindowTitle 'Online Machines' -ButtonStyle 'OKCancel'
    
    ## Copy Installation Media, and installer/logging .bat script to each online machine.
    ForEach($computer in $onlinemachines){
        RoboCopy "$directory\$appname" "\\$computer\C$\Flags\Temp\$appname" /mir /mt:32
        Copy-Item -Path "C:\Temp\install$appname.bat" -Destination "\\$computer\C$\Flags\Temp" -Force}
    ## Inform user of such.
    Message-Box -Message "Installation Script Copied! Initiating Download / Install on:`r`n$(($onlinemachines) -join ',')" -WindowTitle 'Script Copied' -ButtonStyle 'OKCancel'
    
    ## Execute installer/logging .bat script on each online machine.
    start "C:\Windows\System32\pstools\psexec.exe" -ArgumentList "-s","-d","\\$(($onlinemachines) -join ',')","C:\Flags\Temp\Install$appname.bat"
    ## Notify user of successful execution, and location of log files.
    Message-Box -Message "Deployment started. Installation Logs will be created at \\LOCALHOSTNAME\C$\Flags. Start/End logs will be created at Client Server Location\automation\MassDeployment\Files\Logs\LOCALHOSTNAME\ for each host as each host starts and finishes the install." -WindowTitle 'Installation Started' -ButtonStyle 'OKCancel'
    
    ## Log Offline Machines.
    $offlinemachines | Out-File "C:\Temp\$appname-offline.txt"
    Message-Box -Message "Offline machine list output to C:\Temp\$appname-offline.txt for future batches." -WindowTitle 'Offline Logs' -ButtonStyle 'OKCancel'
})
$installform.Controls.Add($InstallButton) 

## Exit Button.
$exitbutton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(360,156)
$exitButton.Size = New-Object System.Drawing.Size(150,46)
$exitButton.BackColor = 'Green'
$exitButton.Text = 'Exit'
$exitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$installform.CancelButton = $exitButton
$installform.Controls.Add($exitButton)

## Show form (GUI).
$installform.AutoSize = $true
$result = $installform.ShowDialog()