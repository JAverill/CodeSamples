## Advanced Remote Troubleshooting Tool
## Version 2.0 (LIVE Production V1.0) - Author: John Averill, IBM
## Description - A GUI tool to remotely execute many different system management functions on remote computers.

## Import Assemblies and Declare Global Variables
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Net
Add-Type -AssemblyName System.Management.Automation

## Declare Global Functions
Function Read-InputBoxDialog([string]$Message, [string]$WindowTitle)
{return [Microsoft.VisualBasic.Interaction]::InputBox($Message, $WindowTitle)}
$ErrorActionPreference = "silentlycontinue"

Function Call-ApplicationManagement{
    # Create Form
    $application = New-Object System.Windows.Forms.Form
    $application.Text = 'Application Management Window'
    $application.Size = New-Object System.Drawing.Size(700,610)
    $application.Position = 'CenterScreen'

    # Create Context Items and Buttons
    $windowlabel = New-Object System.Windows.Forms.Label
    $windowlabel.Location = New-Object System.Drawing.Point(10,10)
    $windowlabel.Size = New-Object System.Drawing.Size(360,20)
    $windowlabel.Text = 'Application Management Window:'
    $application.Controls.Add($windowlabel)

    $hostnamelabel = New-Object System.Windows.Forms.TextBox
    $hostnamelabel.Location = New-Object System.Drawing.Point(365,30)
    $hostnamelabel.Size = New-Object System.Drawing.Size(500,60)
    $hostnamelabel.Text = 'Host Name:'
    $hostnamelabel.TextAlign = 'Center'
    $hostnamelabel.BackColor = 'LightBlue'
    $hostnamelabel.ReadOnly = $true
    $application.Controls.Add($hostnamelabel)

    $hostnameinput = New-Object System.Windows.Forms.TextBox
    $hostnameinput.Location = New-Object System.Drawing.Point(365,60)
    $hostnameinput.Size = New-Object System.Drawing.Size(500,60)
    $hostnameinput.TextAlign = 'Center'
    $hostnameinput.BackColor = 'Red'
    $hostnameinput.ReadOnly = $false
    IF($hostname){$hostnameinput.Text = $hostname}
    $application.Controls.Add($hostnameinput)

    $RefreshButton = New-Object System.Windows.Forms.Button
    $RefreshButton.Location = New-Object System.Drawing.Point(365,100)
    $RefreshButton.Size = New-Object System.Drawing.Size(150,46)
    $RefreshButton.Text = 'Refresh'
    $RefreshButton.BackColor = 'Yellow'
    $RefreshButton.Add_Click({
        $hostname = $hostnameinput.Text
        ForEach($app in $apps){[void] $AppBox.Items.Remove($app)}
        $application.refresh
        $apps = Get-WmiObject -ComputerName $hostname -Class Win32_Product | Select-Object -ExpandProperty Name | Sort-Object
        ForEach($app in $apps){[void] $AppBox.Items.Add($app)}
        $application.refresh
    })
    $application.Controls.Add($RefreshButton)

    $RemoveButton = New-Object System.Windows.Forms.Button
    $RemoveButton.Location = New-Object System.Drawing.Point(540,100)
    $RemoveButton.Size = New-Object System.Drawing.Size(150,46)
    $RemoveButton.Text = 'Uninstall Selected'
    $RemoveButton.BackColor = 'Red'
    $RemoveButton.Add_Click({
            $function = 'RemoveApp'
            $hostnameinput.Text = $hostname
            IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
            $app = $AppBox.SelectedItem
            $output = @()
            $output = "`r`nAttempting to remove $app from $hostname."
            $OutputBox.Text += $output
            $application.refresh
            $WMIPackages = Get-WMIObject -Class win32_product -EA 'SilentlyContinue' -ComputerName $hostname
            $WMIs = $WMIPackages | Where-Object {$_.Name -eq $app}
            IF (!$WMIs) 
            {
                $output =  "`r`nERROR: Could not find $app on remote machine."
                $OutputBox.Text += $output
                $application.refresh
            }
            ELSE 
            {
                ForEach($WMI in $WMIs)
                {
                    $ProductID = $WMI | Select -ExpandProperty IdentifyingNumber
                    $AppName = $WMI | Select -ExpandProperty Name
                    $proceed = Read-InputBoxDialog -Message "$AppName Identified! Remove?" -WindowTitle 'Application'    
                    IF($proceed -eq 'Yes'){Invoke-Command -ComputerName $hostname -ArgumentList $ProductID -EA 'SilentlyContinue' -ScriptBlock{
                        param($ProductID)               
                        Start-Process -FilePath msiexec.exe -ArgumentList /uninstall,$ProductID,/passive,/norestart -Wait}}    
                    ELSEIF($proceed -eq 'Y'){Invoke-Command -ComputerName $hostname -ArgumentList $ProductID -EA 'SilentlyContinue' -ScriptBlock{
                        param($ProductID)               
                        Start-Process -FilePath msiexec.exe -ArgumentList /uninstall,$ProductID,/passive,/norestart -Wait}} 
                    ELSEIF($proceed -eq 'END'){$output ="User elected to end search for matching applications.";$OutputBox.Text += "`r`n$output";$application.refresh;return}
                    ELSE{$output = "`r`nUser elected not to install $AppName. Moving on.";$OutputBox.Text += "`r`n$output";$application.refresh;SilentlyContinue}  
                    $post = Get-WMIObject -Class Win32_product -EA 'SilentlyContinue' -ComputerName $hostname | Where-Object {$_.Name -eq $app}
                    IF (!$post) 
                    {
                        $output = "`r`nSuccess!`r`n `r`n$AppName with Product ID: $ProductID uninstalled from $hostname!"
                        $OutputBox.Text += "$output"
                        $application.refresh
                    }
                    ELSE 
                    {
                        $output = "`r`n$AppName with Product ID: $ProductID failed to uninstall! It may have a removal tool that needs to be run, or may be corrupted. Attempt manual removal."
                        $OutputBox.Text += "`r`n$output"
                        $form.refresh
                    }
                }
            }
        })
    $application.Controls.Add($RemoveButton)

    $InstallButton = New-Object System.Windows.Forms.Button
    $InstallButton.Location = New-Object System.Drawing.Point(715,100)
    $InstallButton.Size = New-Object System.Drawing.Size(150,46)
    $InstallButton.Text = 'Install Application'
    $InstallButton.BackColor = 'Red'
    $InstallButton.Add_Click({
        $function = 'InstallApp'
        $hostnameinput.Text = $hostname
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
        $showlist = Read-InputBoxDialog -Message 'Do you wish to see a list of applications in the backup directory? NOTE: Populating the list takes a while, but without it you must know the exact name of the folder you are installing from.' -WindowTitle 'Show List?'
        IF($showlist -like 'No') {$appname = Read-InputBoxDialog -Message 'Enter Application Folder Name exactly as it appears in the Software Share:' -WindowTitle 'Application'}
        IF($showlist -like 'Yes') {
            $installform = New-Object System.Windows.Forms.Form
            $installform.Text = 'Available Applications.'
            $installform.Size = New-Object System.Drawing.Size(520,700)
            $installform.StartPosition = 'CenterScreen'

            $BackupListBox = New-Object System.Windows.Forms.ListBox
            $BackupListBox.Location = New-Object System.Drawing.Point(10,60)
            $BackupListBox.Size = New-Object System.Drawing.Size(340,600)
            $BackupListBox.Height = 700
            # Get Software Share Folders, and show only folders that have applications that can be installed remotely.
            $directory = "CLIENT NAS LOCATION"
            CD $directory
            $allfolders = dir | Select-Object -ExpandProperty Name | Sort-Object
            $outputapps = @()
            ForEach ($folder in $allfolders){
                [Void]$BackupListBox.Items.Add($folder)
            }
            $BackupListBox.add_SelectedIndexChanged($SelectedFile)
            $installform.Controls.Add($BackupListBox)

            $2ndInstallButton = New-Object System.Windows.Forms.Button
            $2ndInstallButton.Location = New-Object System.Drawing.Point(360,100)
            $2ndInstallButton.Size = New-Object System.Drawing.Size(150,46)
            $2ndInstallButton.Text = "Install on`r`n$hostname"
            $2ndInstallButton.BackColor = 'Red'
            $2ndInstallButton.Add_Click({  
                $appname = $BackupListBox.SelectedItem
                $installform.refresh
            })
            $2ndInstallButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $installform.AcceptButton = $2ndInstallButton
            $installform.Controls.Add($2ndInstallButton) 

            $exitbutton = New-Object System.Windows.Forms.Button
            $exitButton.Location = New-Object System.Drawing.Point(360,156)
            $exitButton.Size = New-Object System.Drawing.Size(150,46)
            $exitButton.BackColor = 'Green'
            $exitButton.Text = 'Exit'
            $exitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $installform.CancelButton = $exitButton
            $installform.Controls.Add($exitButton)

            $installform.AutoSize = $true
            $result = $installform.ShowDialog()
        }
        IF($result -eq [System.Windows.Forms.DialogResult]::OK){$appname = $BackupListBox.SelectedItem}
        $OutputBox.Text += "`r`nInstalling $appname on $hostname."
        $application.refresh
        $appdirectory = "CLIENT NAS LOCATION"
        $OutputBox.Text += "`r`nCopying application to cache for local installation."
        $application.refresh
        $testforfile = Test-Path "\\$hostname\C$\Flags\Temp\$appname"
        IF($testforfile -eq $False){Copy-Item -Path $appdirectory\$appname -Destination "\\$hostname\C$\Flags\Temp" -Recurse -Force}
        IF(Test-Path "\\$hostname\C$\Flags\Temp\$appname\Deploy-Application.ps1"){
            IF(Invoke-Command -ComputerName $hostname -ArgumentList $appname -ScriptBlock{
                param($appname)
                powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -WindowStyle Hidden -File "C:\Flags\Temp\$appname\Deploy-Application.ps1"})
                {
                    $output = "`r`nSuccess! Installation of $appname executed on $hostname."
                    $OutputBox.Text += "$output"
                    $application.refresh
                }
            ELSE
            {
                $output = "`r`nFailure! Installation of $appname failed on $hostname. Attempt Manual installation. See log in \\$hostname\C$\Flags\$appname.log for more information. `r`n `r`n"
                $OutputBox.Text += "$output"
                $application.refresh
            }
        }
        ELSE{
            $alldir = GCI "\\$hostname\C$\Flags\Temp\$appname" | Select-Object -ExpandProperty Name
            IF($alldir -contains "*.msi"){
                $msis = $alldir | Select-Object {Where $._ -Contains ".msi"}
                $msideploy = Read-InputBoxDialog -Message 'No Deployment Package identified, but there is an .MSI package. Attempt MSI install?' -WindowTitle 'Deploy vs MSI'
                IF($msideploy -like 'Yes'){
                    ForEach ($msipackage in $msis){
                        $installorno = Read-InputBoxDialog -Message "$msipackage identified. Install?"
                        IF($installorno -like 'Yes'){
                            IF(Invoke-Command -ComputerName $hostname -ArgumentList $appname,$msipackage -ScriptBlock{
                            param($appname)
                            param($msipackage)
                            Start-Process -FilePath msiexec.exe -ArgumentList /i,"C:\Flags\Temp\$appname\$msipackage",/passive,/norestart -Wait
                            }){
                                $output = "`r`nSuccess! Installation of $appname executed on $hostname ."
                                $OutputBox.Text += "$output"
                                $application.refresh()
                            }
                            ELSE{
                                $output = "`r`nFailure! Installation of $appname via $msipackage failed on $hostname. Attempt Manual installation. See log in \\$hostname\C$\Flags\$appname.log for more information. `r`n `r`n"
                                $OutputBox.Text += "$output"
                                $application.refresh()
                            }
                        }
                        ELSE{$output += "User declined MSI installation of $appname via $msipackge"}
                    }   
                }
                ELSE{$output += "User declined MSI installation and Deployment package is not present."}
            }
        }
    })
    $application.Controls.Add($InstallButton)

    $DoneButton = New-Object System.Windows.Forms.Button
    $DoneButton.Location = New-Object System.Drawing.Point(540,710)
    $DoneButton.Size = New-Object System.Drawing.Size(150,46)
    $DoneButton.Text = 'Done/Return'
    $DoneButton.BackColor = 'Green'
    $DoneButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $application.AcceptButton = $DoneButton
    $application.Controls.Add($DoneButton)

    $AppBox = New-Object System.Windows.Forms.ListBox
    $AppBox.Location = New-Object System.Drawing.Point(10,60)
    $AppBox.Size = New-Object System.Drawing.Size(340,600)
    $AppBox.Height = 700
        # Get Installed Applications if Hostname already declared.
        IF($hostname){$apps = Get-WmiObject -ComputerName $hostname -Class Win32_Product | Select-Object -ExpandProperty Name | Sort-Object
        ForEach($app in $apps){[void] $AppBox.Items.Add($app)}}
    $AppBox.add_SelectedIndexChanged($SelectedFile)
    $application.Controls.Add($AppBox)

    $OutputLabel = New-Object System.Windows.Forms.TextBox
    $OutputLabel.Location = New-Object System.Drawing.Point(365,170)
    $OutputLabel.Size = New-Object System.Drawing.Size(500,60)
    $OutputLabel.Text = 'Output :'
    $OutputLabel.TextAlign = 'Center'
    $OutputLabel.BackColor = 'LightBlue'
    $OutputLabel.ReadOnly = $true
    $application.Controls.Add($OutputLabel)

    $OutputBox = New-Object System.Windows.Forms.TextBox
    $OutputBox.Location = New-Object System.Drawing.Point(365,200)
    $OutputBox.Size = New-Object System.Drawing.Size(500,600)
    $OutputBox.Height = 500
    $OutputBox.Multiline = $True
    $OutputBox.ReadOnly = $True
    $OutputBox.WordWrap = $True
    $OutputBox.Scrollbars = 'Vertical'
    $application.Controls.Add($OutputBox)

    # Create Events

    # Display Form
    $application.AutoSize = $true
    $result = $application.ShowDialog()

    # End
    IF($result -eq [System.Windows.Forms.DialogResult]::OK){
        $subformoutput = $OutputBox.Text
        $ListBox2.Text += "`r`n-------Application Management-------`r`n$subformoutput`r`n-------End Application Management-------"
        $form.refresh}
    }

    Function C-Share{
        & explorer.exe "\\$hostname\C$"
    }

    Function Service-Status{
        IF(!$hostname){$hostname = Read-InputBoxDialog -Prompt 'Enter Hostname'}
        $service = Read-InputBoxDialog -message 'Enter Name of Service' -windowtitle 'Service Name'
        $online = Test-Connection $hostname -Quiet -Count 1
        IF(!$online){Write-Host "$hostname offline or not responding."}
        $status = Get-Service -ComputerName $hostname -Name $service | Select -ExpandProperty Status
        $starttype = Get-Service -ComputerName $hostname -Name $service | Select -ExpandProperty StartType
        IF(!$status){$result = "Could not find $service, check spelling and formatting.";$listbox2.text += "`r`n$result";$form.refresh;Return}
        $result = "The Service $service is $status on $hostname and is set to Startup Type $starttype."
        $listbox2.text += "`r`n$result"
        $form.refresh}

    Function Get-User{
        $function = 'GetUser'
        
        $form.refresh
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
        ForEach ($serverline in @(query user /server:$hostname) -split "\n") 
            {
                $parsed_server = $serverline -split '\s'
            }
        $boottime = Get-CIMInstance -ComputerName $hostname -ClassName win32_operatingsystem | select -ExpandProperty lastbootuptime
        $user = $parsed_server[1]
        $state = $parsed_server[31]
        $logindate = $parsed_server[40]
        $logintime = $parsed_server[41]
        $logintime2 = $parsed_server[42]
        $output = "`r`n$hostname `r`n `r`nBoot Time: $boottime `r`nUser: $user `r`nUser State: $state `r`nLogon Time: $logindate $logintime $logintime2"
        $listBox2.Text += "$output`r`n**********"
        $form.refresh}
    
    Function View-HDSpace{
        $function = 'ViewHDSpace'
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
        $form.refresh
        $ListBox2.Text += "`r`nCalculating HD Space on $hostname ..."
        $form.refresh
        $result1 = Get-WmiObject -ComputerName $hostname -Class Win32_logicalDisk -Filter "DeviceID='C:'"
        $Space = $Result1.size / 1gb -as [int]
        $FreeSpace = $Result1.Freespace / 1gb -as [int]
        $output = "`r`n Disk Space = $Space GB `r`n Free Space = $FreeSpace GB`r`n**********"
        $listBox2.Text += $output
        $form.refresh}
    
    Function Reinstall-SCCM {
        $function = 'ReinstallSCCM'
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
        $form.refresh
        $ListBox2.Text += "`r`nUninstalling Software Center... `r`n"
        $form.refresh
        $directory = Test-Path "\\$hostname\C$\Flags\Temp\Client"
        IF($directory -eq $False){Copy-Item -Path "\\VA10P51090.US.AD.WELLPOINT.COM\Client" -Destination "\\$hostname\C$\Flags\Temp" -Force -Recurse}
        IF (Invoke-Command -ComputerName $hostname -ScriptBlock {Start-Process -FilePath "C:\Flags\Temp\Client\ccmsetup.exe" -ArgumentList /uninstall -Wait -PassThru})
        {
            $output = 'CCMSETUP.EXE Uninstall Successfully Executed on Remote Machine. Uninstall log located in CCMSETUP logs folder on remote machine.'
            $ListBox2.Text += "`r`n$output"
            $form.refresh
            Start-Sleep -Seconds 60.0
            $ListBox2.Text += "`r`nInstalling Software Center..."
            IF (Invoke-Command -ComputerName $hostname -EA SilentlyContinue -ScriptBlock {Start-Process -FilePath "C:\Flags\Temp\Client\Install.bat" -Wait -PassThru})
                 {
                     $ListBox2.Text += "`r`nCCM Client Install Started. Allow 20 minutes for client update.`r`n**********"
                     $form.refresh
                 }
               Else {
                     $ListBox2.Text += "`r`nReinstall failed to execute. Reboot machine and try again or perform manual reinstall.`r`n**********"
                     $form.refresh
                    }
        }
        Else 
        {
            $output = "`r`nCCMSETUP.EXE failed to Start. Manual execution required.`r`n**********"
            $listBox2.Text += "`r`n$output`r`n**********"
            $form.refresh
        }}
    
    Function View-CCMCache{
        $function = 'ViewCCMCache'
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
        $form.refresh
        $ListBox2.Text += "`r`nChecking CCM Cache Allocation on $hostname`r`n"
        $form.refresh
        $Space2 = Get-WmiObject -ComputerName $hostname -Namespace ROOT\CCM\SoftMgmtAgent -Query "Select size from CacheConfig" | select -ExpandProperty size
        $output = "`r`n$hostname`r`nCCMCache Space $Space2 MB"
        $listBox2.Text += "`r`n$output`r`n**********"
        $form.refresh}
    
    Function Increase-CCMCache{
        $function = 'IncreaseCCMCache'
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
        $form.refresh
        $ListBox2.Text += "`r`nSetting CCM Cache Allocation to 50GB on $hostname ...`r`n"
        $form.refresh
        Invoke-Command -ComputerName $hostname -EA SilentlyContinue -ScriptBlock{
          $Cache = Get-WmiObject -NameSpace Root\CCM\SoftMgmtAgent -Class CacheConfig
          $Cache.Size = '51200'
          $Cache.Put()
          Restart-Service -Name CcmExec
          }
        $Space3 = Get-WmiObject -ComputerName $hostname -Namespace ROOT\CCM\SoftMgmtAgent -Query "Select size from CacheConfig" | select -expandproperty size
        $output = "`r`nFunction Complete.`r`n$hostname`r`n$Space3 MB"
        $listBox2.Text += "`r`n$output`r`n**********"
        $form.refresh}

    Function Clear-CCMCache{
        $function = 'ClearCCMCache'
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
        $form.refresh
        $ListBox2.Text += "`r`nClearing CCM Cache on $hostname ..."
        $form.refresh
        Invoke-Command -ComputerName $hostname -EA SilentlyContinue -ScriptBlock{
          $resman= New-Object -ComObject "UIResource.UIResourceMgr"
          $cacheinfo= $resman.GetCacheInfo()
          $cacheinfo.GetCacheElements() | foreach {$cacheinfo.DeleteCacheElement($_.CacheElementID)}
          }
        $Space4 = Get-WmiObject -ComputerName $hostname -Namespace ROOT\CCM\SoftMgmtAgent -Query "Select size from CacheConfig" | Select-Object -ExpandProperty size
        $output = "`r`nFunction Complete.`r`n$hostname`r`n$Space4 MB Available in CCM Cache."
        $listBox2.Text += "`r`n$output`r`n**********"
        $form.refresh}
    
    Function Rename-Citrix{
        $function = 'RenameCitrix'
        $User = Read-InputBoxDialog -message 'Please Enter User Domain ID.' -WindowTitle 'Domain ID'
        $form.refresh
        $ListBox2.Text += "`r`nExporting list of Citrix Servers..."
        $form.refresh
        $directory = Test-Path "C:\Tools\citrix_profile_paths.txt"
        IF($directory -eq $false){Copy-Item -Path '\\VDAASW1015915\C$\tools\citrix_profile_paths.txt' -Destination 'C:\tools\citrix_profile_paths.txt' -Force}
        $Paths = Get-Content "C:\tools\citrix_profile_paths.txt"
        $ListBox2.Text += "`r`nResetting all Citrix Profiles for User: $User ..."
        $form.refresh
        ForEach ($path in $paths){
          $pathtest = Test-Path $path\$user
          $pathtestus = Test-Path "$path\$user.us"
          If ($pathtest = $True){
                $newname = "$user" + "_" + "CTX"
                Rename-Item "$path\$user" -NewName $newname -Force
                $output = "`r`nProfile Renamed to $newname on server location $path."
                $listbox2.text += "$output"
                $form.refresh
            }
          ElseIf ($pathtestus = $True) {
                $newname = "$user" + ".us" + "_" + "CTX"
                Rename-Item "$path\$user.us" -NewName $newname -Force
                $output = "`r`nProfile renamed to $user _CTX on server location $path."
                $listbox2.text = "$output"
                $form.refresh
            }
          Else {
               $output = "`r`nProfile Not Found on server location $path.`r`n"
               $listbox2.Text += "$output"
               $form.refresh
               }
            }
        $form.refresh}

    Function ScanRepair-OS{  
        $function = 'ScanRepairOS'
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
        $form.refresh
        $listBox2.Text += "`r`nProcessing. Note: Full System Scan / Repair takes a while!"
        Invoke-Command -ComputerName $hostname -ScriptBlock{
              $output = @()
              $CBS = Start-Process -FilePath "${env:Windir}\System32\SFC.EXE" -ArgumentList /scannow -Wait -passthru
              IF($CBS){
                   $output += "`r`nSuccess! System verification scan complete. Reboot to apply any changes, view log file at %windir%\Logs\CBS\CBS.log."
                   }
              ELSE{
                   $output += "`r`nSFC /Scannow did not execute. Significant local issues likely."
                   }
            }
        $listBox2.Text += "`r`n$output`r`n**********"
        $form.refresh}
    
    Function Enable-UAC{
        $function = 'EnableUAC'
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
        $form.refresh
        $ListBox2.Text += "`r`nEnabling UAC on $hostname ..."
        $form.refresh
        $enablelua = Invoke-Command -ComputerName $hostname -ScriptBlock{Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableLUA}
        IF(!$enablelua)
            {
            $post5 = Invoke-Command -ComputerName $hostname -ScriptBlock{New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableLUA -PropertyType DWORD -Value 1 -Force}
            IF($post5)
                {
                    $output = "`r`nOperation Completed Successfully!`r`nAdded Registry Key EnableLUA.`r`n**********"
                    $listBox2.Text += $output
                    $form.refresh
                }
            ELSE
                {
                    $output = "`r`nCannot Complete Operation. `r`nUnable to Invoke Commands on remote machine $hostname`r`n**********"
                    $listBox2.Text += $output
                    $form.refresh
                }
            }
        ELSEIF($enablelua -eq 0)
            {
                Invoke-Command -ComputerName $hostname -ScriptBlock{Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableLUA -Value 1 -Force}
                $output = "`r`nOperation Completed Successfully! `r`n Edited Registry Key and enabled LUA.`r`n**********"
                $listBox2.Text += $output
                $form.refresh
            }    
        ELSEIF($enablelua -eq 1)
            {
                $output = "`r`nCannot complete Operation!`r`nUAC Already Enabled!`r`n**********"
                $listBox2.Text += $output
                $form.refresh
            }
        ELSE
            {
                $output = "`r`nCannot complete Operation!`r`nCannot locate registry key. Attempt manually!`r`n**********"
                $listBox2.Text += $output
                $form.refresh
            }}
        
    Function System-Cleanup{
        $function = 'SystemCleanup' 
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'} 
        $form.refresh
        $listBox2.Text += "`r`nProcessing System Cleanup on $hostname..."
        $form.refresh
        $result3 = Invoke-Command -ComputerName $hostname -EA 'SilentlyContinue' -ScriptBlock{
        $VolumeCachesRegDir = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
        $CacheDirItemNames = Get-ItemProperty "$VolumeCachesRegDir\*" | select -ExpandProperty PSChildName
        $CacheDirItemNames | 
            %{
                $exists = Get-ItemProperty -Path "$VolumeCachesRegDir\$_" -Name "StateFlags6553"
                If (($exists -ne $null) -and ($exists.Length -ne 0))
                    {
                        Set-ItemProperty -Path "$VolumeCachesRegDir\$_" -Name StateFlags6553 -Value 2
                    }
                Else
                    {
                        New-ItemProperty -Path "$VolumeCachesRegDir\$_" -Name StateFlags6553 -Value 0 -PropertyType DWord
                    }
            }
        Start-Sleep -Seconds 3
        Start-Process -FilePath CleanMgr.exe -ArgumentList '/sagerun:65535' -WindowStyle Hidden -PassThru }
        $output = "`r`nSet State Flags for System Cleanup Tool on $hostname."
        $listBox2.Text += $output
        $form.refresh
        IF($result3)
            {
                $output = "`r`nFull System Cleanup Executed.`r`n**********"
                $listBox2.Text += $output
                $form.refresh
            }
        ELSE
            {
                $output = "`r`nRemote execution of cleanmgr.exe failed. Attempt manual cleanup.`r`n**********"
                $listBox2.Text += $output
                $form.refresh
            }}
    
    Function Clear-OfficeCache{
        $function = 'ClearOfficeCache'
        IF(!$hostname){$hostname = Read-InputBoxDialog -message 'Please Enter Asset Host Name.' -WindowTitle 'Host Name'}
            $username = Read-InputBoxDialog -Message 'Enter UserName' -WindowTitle 'Username'
            $listBox2.Text += "`r`nProcessing Office Cache Clear on $hostname..."
            $form.refresh
            $officecache = Invoke-Command -ComputerName $hostname -EA 'SilentlyContinue' -ArgumentList $username -ScriptBlock{
                param($username)
                Remove-Item -Path "C:\Users\$username\AppData\Local\Microsoft\Office\15.0\OfficeFileCache" -Force -Recurse
                New-Item -ItemType Directory -Path "C:\users\$username\AppData\Local\Microsoft\Office\15.0\" -Name OfficeFileCache -Force
            }
        IF($officecache)
        {
            $output = "`r`nOffice Document Cache Cleared!`r`n**********"
            $listBox2.Text += $output
            $form.refresh
        }
        ELSE
        {
            $output = "`r`nOffice Document Cache Not Cleared! Could not locate directory remotely.`r`n**********"
            $listBox2.Text += $output
            $form.refresh
        }}

    Function Check-ImageDate{
        $function = 'CheckImageDate'
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Host Name'}
        $form.refresh
        $listBox2.Text += "`r`nChecking Image Date..."
        $form.refresh
        $output = Invoke-Command -ComputerName $hostname {
            systeminfo | findstr /i "original"}
        $listBox2.Text += "`r`n$output`r`n**********"
        $form.refresh}

    Function Get-ADGroups{
        $function = 'GetADGroups'
        $username = Read-InputBoxDialog -Message 'Enter Active Directory Username' -WindowTitle 'Username'
        $form.refresh
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        $result10 = Get-ADPrincipalGroupMembership $username | Sort-Object -Property samaccountname
        $form.refresh
        $output = Write-Output $result10 | Select-Object -ExpandProperty samaccountname
        FOREACH($item in $output){
            $ListBox2.Text += "`r`n$item"}
        $listbox2.Text += "`r`nGet AD Groups Function Complete.`r`n**********"
        $form.refresh}
    
    Function Unlock-Reset{
        $function = 'UnlockReset'
        $username = Read-InputBoxDialog -Message 'Enter Active Directory Username' -WindowTitle 'Username'
        $form.refresh
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        $result11 = Unlock-ADAccount -Identity $username -PassThru
        IF($result11){$output = "`r`nAccount $username unlocked in Active Directory. `r`n `r`n"}
        ELSE{$output = "`r`nAccount $username not found, attempt manual unlock. `r`n `r`n"}
        $form.refresh
        $listBox2.Text += $output
        $form.refresh
        $reset = Read-InputBoxDialog -Message 'Do you wish to reset the password on this account?' -WindowTitle 'Reset?'
        IF($reset -like 'Yes')
        {
            $SecurePW = Read-InputBoxDialog -Message 'Enter New Password' -WindowTitle 'Password'
            Set-ADAccountPassword -Identity $username -NewPassword $SecurePW
            Set-ADUser $username -ChangePasswordAtLogon $True
            $output = "`r`nPassword reset to $SecurePW and user must change password on next logon."
        }
        ELSE{$output = "`r`nUser elected not to reset password, or user response not understood."}
        $listBox2.Text += "$Output`r`n**********"
        $form.refresh}
    
    Function View-OSVersion{
        $function = 'ViewOSVersion'
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
        $form.refresh
        $listBox2.Text += "`r`nProcessing... "
        $form.refresh
        $result13 = Get-ADComputer -Properties * -identity $hostname
        $computername = $result13 | Select -ExpandProperty Name
        $OS = $result13 | Select -ExpandProperty OperatingSystem
        $OSBuild = $result13 | Select -ExpandProperty OperatingSystemVersion
        IF($OSBuild -eq 14393){$OSVersion = 1607}
        ELSEIF($OSBuild -eq '10.0 (15063)'){$OSVersion = 1703}
        ELSEIF($OSBuild -eq '10.0 (16299)'){$OSVersion = 1709}
        ELSEIF($OSBuild -eq '10.0 (17134)'){$OSVersion = 1803}
        ELSEIF($OSBuild -eq '10.0 (17763)'){$OSVersion = 1809}
        ELSE{$OSVersion = 'Unable to determine OS Version'}
        $output = "`r`nComputer ... $computername `r`nOS ... $OS `r`nOS Build ... $OSBuild `r`nOS Version ... $OSVersion"
        $form.refresh
        $listBox2.Text += Write-Output "$output`r`n**********"
        $form.refresh}
    
    Function Set-NetProfile{
        $function = 'SetNetProfile'
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        $currentprofile = Invoke-Command -ComputerName $hostname {Get-NetConnectionProfile}
        $currentprofilename = $currentprofile | Select -ExpandProperty Name
        $currentprofiletype = $currentprofile | Select -ExpandProperty NetworkCategory
        $ListBox2.Text += "`r`nGetting profile information for network adapters on $hostname ... `r`n `r`n"
        $form.refresh
        $output = "`r`nCurrent Profile is $currentprofilename .`r`nCurrent Profile is listed as $currentprofiletype `r`n `r`n"
        $listBox2.Text += $output
        $form.refresh
        Invoke-Command -ComputerName $hostname -EA SilentlyContinue -ArgumentList $currentprofilename {param($currentprofilename) Set-NetConnectionProfile -Name $currentprofilename -NetworkCategory Private -EA SilentlyContinue}
        $newprofiletype = Invoke-Command -ComputerName $hostname -ArgumentList $currentprofilename {param($currentprofilename) Get-NetConnectionProfile -Name $currentprofilename} | Select -ExpandProperty NetworkCategory
        $output = "`r`nProfile type reset from $currentprofiletype to $newprofiletype ."
        $listBox2.Text += $output
        $form.refresh}

    Function Compare-Directories{
        $function = 'CompareDirectories'
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        $source = Read-InputBoxDialog -Message 'Enter Source File Path' -WindowTitle 'Source directory'
        $destination = Read-InputBoxDialog -Message 'Enter Destination File Path' -WindowTitle 'Destination directory'
        $sourceresult = GCI -Recurse $source
        $destinationresult = GCI -Recurse $destination
        $output = "`r`nSource: $source `r`nDestination: $destination `r`n `r`n"
        $listBox2.Text += $output
        $form.refresh
        $results = Compare -ReferenceObject $sourceresult -DifferenceObject $destinationresult -Property Name -PassThru | ? {$_.sideindicator -eq "<="}
        ForEach ($result in $results){
            $output = "`r`nFile $result is not present on $destination .`r`n"
            $listBox2.Text += $output}
        IF($listbox2.Text -eq "Source: $source `r`nDestination: $destination `r`n `r`n"){$output = 'No differences found between the two directories';$listBox2.Text += $output; $form.refresh}}

    Function Copy-Directory{
        $function = 'CopyDirectory'
        $form.refresh
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        $source = Read-InputBoxDialog -Message 'Enter path of directory to be copied.' -WindowTitle 'Source'
        $destination = Read-InputBoxDialog -Message 'Enter path of directory to copy Source to.' -Window Title 'Destination'
        $output = "`r`nCopying files...`r`n `r`nSource: $source `r`nDestination: $destination `r`n `r`n"
        $listBox2.Text += $output
        $form.refresh
        $output = robocopy "$source" "$destination" /mir /mt:32
        IF($output -contains "ERROR"){$ListBox2.Text += "`r`nCopy appears to have failed! Log below. `r`n `r`n `r`n"}
        ELSE{$listBox2.Text += 'Full copy succeeded! ROBOCOPY results below.'}
        $listBox2.Text += $output
        $form.refresh}

    Function Force-Action{
        $function = 'ForceActions'
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
        $form.refresh
        $ListBox2.Text += "`r`nForcing Software Actions on $hostname ..."
        $form.refresh
        IF(Invoke-Command -ComputerName $hostname -EA 'SilentlyContinue' -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"})
        {
            $output = "`r`nRequest Machine Assignments action triggered on $hostname"
            $listBox2.Text += $output
            $form.refresh
        }
        ELSE{$output = "`r`nCould not reach SMS Client on $hostname to trigger Request Machine Assignments action.";$listBox2.Text += $output;$form.refresh}
        IF(Invoke-Command -ComputerName $hostname -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000121}"})
        {
            $output = "`r`nApplication manager policy action triggered on $hostname"
            $listBox2.Text += $output
            $form.refresh
        }
        ELSE{$output = "`r`nCould not reach SMS Client on $hostname to trigger Application manager policy action.";$listBox2.Text += $output;$form.refresh}
        IF(Invoke-Command -ComputerName $hostname -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000113}"})
        {
            $output = "`r`nScan by Update Source Action triggered on $hostname"
            $listBox2.Text += $output
            $form.refresh
        }
        ELSE{$output = "`r`nCould not reach SMS Client on $hostname to trigger Scan by Update Source action.";$listBox2.Text += $output;$form.refresh}
        IF(Invoke-Command -ComputerName $hostname -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000108}"})
        {
            $output = "`r`nSoftware Updates Assignments Evaluation Cycle triggered on $hostname"
            $listBox2.Text += $output
            $form.refresh
        }
        ELSE{$output = "`r`nCould not reach SMS Client on $hostname to trigger Software Updates Assignments Evaluation Cycle.";$listBox2.Text += $output;$form.refresh}
        $output = "`r`nForce Actions Function Complete!"
        $listBox2.Text += "$output`r`n**********"
        $form.refresh}

    Function Font-Fix{
        $function = 'FontFix'
        $form.refresh
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
        $ListBox2.Text += "`r`nSetting Font Permissions on $hostname ..."
        $form.refresh
        IF(Invoke-Command -ComputerName $hostname -EA SilentlyContinue -ScriptBlock{Get-ACL C:\Windows\Fonts\Arial.ttf | Set-ACL C:\Windows\Fonts\*.*})
        {
            $output = "`r`nPermissions for font files on $hostname updated."
            $listBox2.Text += "$output`r`n**********"
            $form.refresh
        }
        ELSE{$ListBox2.Text += "`r`nCould not update font permissions on $hostname. Attempt locally.`r`n**********";$form.refresh}
        IF(Invoke-Command -ComputerName $hostname -EA SilentlyContinue -ScriptBlock{Get-ACL C:\Windows\Fonts\Arial.ttf | Set-ACL C:\Windows\Fonts})
        {
            $output = "`r`nPermissions for font folder on $hostname updated.`r`n**********"
            $listBox2.Text += $output
            $form.refresh
        }
        ELSE{$ListBox2.Text += "`r`nCould not update font permissions on $hostname. Attempt locally.`r`n**********";$form.refresh}}

    Function Get-Java{
        $function = 'GetJava'
        $form.refresh
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
        $result20 = Invoke-Command -ComputerName $hostname {(get-childitem "HKLM:\SOFTWAARE\wow6432node\JavaSoft\Java Runtime Environment" -ErrorAction 'SilentlyContinue' | Select-Object -ExpandProperty pschildname[1])}
        $output = "`r`nJava Versions Detected: $result20"
        $listBox2.Text += "$output`r`n**********"
        $form.refresh}

    Function Side-By-Side{
        $function = 'SideBySide'
        $form.refresh
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
        Invoke-Command -ComputerName $hostname -ScriptBlock
        {
            $alreadydeleted = Test-Path "C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe.Config"
            IF($alreadydeleted -eq $true)
            {
                Remove-Item "C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe.Config" -Force
            }
        }
        $diditdelete = Test-Path "\\$hostname\C$\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe.Config"
        IF($diditdelete -eq $false)
        {
            $output = "`r`nPowershell Config file removed on remote machine $hostname. `r`nSide-By-Side Error should be resolved."
            $listBox2.Text += "$output`r`n**********"
            $form.refresh
        }
        ELSE
        {
            $output = "`r`nFailed to remove Config file, attempt manually."
            $listBox2.Text += "$output`r`n**********"
            $form.refresh
        }}

    Function Activate-Office{
        $function = 'ActivateOffice'
        $form.refresh
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
        IF((Test-Path "\\$hostname\C$\Program Files (x86)\Microsoft Office\Office15\ospp.vbs") -eq $True)
        {$output = Invoke-Command -ComputerName $hostname -ScriptBlock{ 
                cd "C:\Program Files (x86)\Microsoft Office\Office15"
                cscript .\ospp.vbs /act
                }
        }
        ELSEIF((Test-Path "\\$hostname\C$\Program Files (x86)\Microsoft Office\Office16\ospp.vbs") -eq $True)
        {$output = Invoke-Command -ComputerName $hostname -ScriptBlock{
                cd "C:\Program Files (x86)\Microsoft Office\Office16"
                cscript .\ospp.vbs /act
                }
        }
        ELSE
        {
            $output = "`r`nCould not find installation of Microsoft Office on $hostname"
        }
        
        $form.refresh
        $listBox2.Text += "$output`r`n**********"
        $form.refresh}

    Function Activate-Windows{
        $function - 'ActivateWindows'
        $form.refresh
        $listBox2.Text += "`r`nProcessing..."
        $form.refresh
        IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
        $output = Invoke-Command -ComputerName $hostname -ScriptBlock{ 
                cd "C:\Windows\System32"
                cscript .\slmgr.vbs /ato
                }
        ELSE
        {
            $output = "`r`nCould not find installation of Microsoft Windows on $hostname. Connectivity issues likely, attempt manually."
        }
        $form.refresh
        $listBox2.Text += "$output`r`n**********"
        $form.refresh}

    Function PS-GUI{
        $form.Refresh
        $powershellform = New-Object System.Windows.Forms.Form
        $powershellform.Size = New-Object System.Drawing.Size(500,500)
        $powershellform.Text = 'PowerShell Interactive Console'
        $consolebox = New-Object System.Windows.Forms.TextBox
        $consolebox.Size = New-Object System.Drawing.Size(470,30)
        $consolebox.Location = New-Object System.Drawing.Point(10,10)
        $consolebox.MultiLine = $True
        $consolebox.WordWrap = $True
        $powershellform.Controls.Add($consolebox)
        $PSoutputbox = New-Object System.Windows.Forms.TextBox
        $PSoutputbox.Height = '300'
        $PSoutputbox.AutoSize = $True
        $PSoutputBox.Size = New-Object System.Drawing.Size(460,300)
        $PSoutputBox.Location = New-Object System.Drawing.Point(10,90)
        $PSoutputBox.WordWrap = $True
        $PSoutputbox.Multiline = $True
        $powershellform.Controls.Add($PSoutputbox)
        $runcode = New-Object System.Windows.Forms.Button
        $runcode.Text = 'Run'
        $runcode.Size = New-Object System.Drawing.Size(50,30)
        $runcode.Location = New-Object System.Drawing.Point(10,50)
        $runcode.BackColor = "Green"
        $runcode.Add_Click({
            $powershellinput = $consolebox.Text
            $powershelloutput = powershell.exe -executionpolicy unrestricted $powershellinput
            ForEach ($thing in $powershelloutput){$PSOutputbox.Text += "`r`n$thing"}
            $powershellform.Refresh
        })
        $powershellform.Controls.Add($runcode)
        $closebutton = New-Object System.Windows.Forms.Button
        $closebutton.Text = 'Exit'
        $closebutton.Size = New-Object System.Drawing.Size(50,30)
        $closebutton.Location = New-Object System.Drawing.Point(10,400)
        $closebutton.BackColor = "Red"
        $powershellform.CancelButton = $closebutton
        $closebutton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $powershellform.Controls.Add($closebutton)
        $powershellform.ShowDialog()}
    
    Function EDPA{
        ## EDPA Fix GUI

        ## Load Assemblies
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName Microsoft.VisualBasic
        Add-Type -AssemblyName System.Net
        Add-Type -AssemblyName System.Management.Automation

        ## Declare Global Functions
        Function Read-InputBoxDialog([string]$Message, [string]$WindowTitle)
        {return [Microsoft.VisualBasic.Interaction]::InputBox($Message, $WindowTitle)}

        Function Message-Box([string]$Message,[string]$WindowTitle,[string]$ButtonStyle)
        {return [System.Windows.MessageBox]::Show($Message,$WindowTitle,$ButtonStyle)}

        $ErrorActionPreference = "silentlycontinue"

        ## Create Form (GUI)
        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Symantec Fix GUI'
        $form.Size = New-Object System.Drawing.Size(280,300)
        $form.Autosize = $False
        $form.AutoSizeMode = "GrowAndShrink"
        $form.StartPosition = 'CenterScreen'

        ## Tool Name Label
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,10)
        $label.Size = New-Object System.Drawing.Size(110,20)
        $label.Text = "Symantec Fix GUI"
        $label.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $form.Controls.Add($label)

        ## Hostname Label
        $hostlabel = New-Object System.Windows.Forms.Label
        $hostlabel.Location = New-Object System.Drawing.Point(10,40)
        $hostlabel.Size = New-Object System.Drawing.Size(110,20)
        $hostlabel.Text = "Enter Hostname:"
        $hostlabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $form.controls.add($hostlabel)

        ## Ping Host Button
        $pinghost = New-Object System.Windows.Forms.Button
        $pinghost.Location = New-Object System.Drawing.Point(120,10)
        $pinghost.Size = New-Object System.Drawing.Point(110,20)
        $pinghost.Text = 'Ping Host'
        $pinghost.BackColor = 'LightBlue'
        $pinghost.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $pinghost.Add_Click({
            $hostname = $hostnameinput.Text
            $x = 0
            DO {$ping = Test-Connection $hostname -Count 1 -Quiet;$x = ($x + 1)}
            Until ($ping -contains "True" -or $x -eq 10)
            IF($x -eq 10){Message-Box -Message "Extended Ping to $hostname timed out. Machine still offline." -WindowTitle 'Timed Out' -ButtonStyle 'OKCancel'}
            ELSE{Message-Box -Message "$hostname now online." -WindowTitle 'Online!' -ButtonStyle 'OKCancel'}
        })
        $form.controls.add($pinghost)

        ## Hostname Input
        $hostnameinput = New-Object System.Windows.Forms.TextBox
        $hostnameinput.Location = New-Object System.Drawing.Point(120,40)
        $hostnameinput.Size = New-Object System.Drawing.Size(110,20)
        $hostnameinput.TextAlign = 'Center'
        $hostnameinput.BackColor = 'Red'
        $hostnameinput.ReadOnly = $false
        $form.Controls.Add($hostnameinput)
        IF($hostname){$hostnameinput.text = $hostname}

        ## Kill Tasks Button
        $Button1 = New-Object System.Windows.Forms.Button
        $Button1.Location = New-Object System.Drawing.Point(95,130)
        $Button1.Size = New-Object System.Drawing.Size(75,50)
        $Button1.Text = 'Kill Tasks'
        $Button1.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $Button1.BackColor = 'Yellow'
        $Button1.Add_Click({ 
            $form.refresh
            $hostname = $hostnameinput.Text
            Try{
                start "CLIENT NAS LOCATION\automation\EDPAFix\Files\PSTools\PSExec.exe" -ArgumentList "-S","\\$hostname","CLIENT NAS LOCATION\automation\EDPAFix\Files\taskkilledpa.bat"
                Message-Box -Message "EDPA.exe and WDP.exe Killed on $hostname." -WindowTitle "Success!" -ButtonStyle 'OKCancel'}
            Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}
        })
        $form.Controls.Add($Button1)

        ## Deploy DLP HotFix Button
        $Button2 = New-Object System.Windows.Forms.Button
        $Button2.Location = New-Object System.Drawing.Point(95,70)
        $Button2.Size = New-Object System.Drawing.Size(75,50)
        $Button2.Text = 'Deploy DLP HotFix'
        $Button2.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $Button2.BackColor = 'Yellow'
        $Button2.Add_Click({
            $form.refresh
            $hostname = $hostnameinput.Text
            Try{
                RoboCopy "CLIENT NAS LOCATION\Symantec_DataLossPreventionforEndpoint-Hotfix_15-5-0208-0_11" "\\$hostname\C$\Flags\Temp\Symantec_DataLossPreventionforEndpoint-Hotfix_15-5-0208-0_11" /mir /mt:32
                start "CLIENT NAS LOCATION\automation\edpafix\files\pstools\psexec.exe" -ArgumentList "-s","\\$hostname","C:\Flags\Temp\Symantec_DataLossPreventionforEndpoint-Hotfix_15-5-0208-0_11\Deploy-Application.exe"
                Message-Box -Message "Symantec DLP for Endpoint Hotfix 15.5.0208.0_11 installed on $hostname" -WindowTitle "Success!" -ButtonStyle 'OKCancel'}
            Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}
        })
        $form.Controls.Add($Button2)

        ## Restore VDAAS Config Button
        $Button3 = New-Object System.Windows.Forms.Button
        $Button3.Location = New-Object System.Drawing.Point(10,130)
        $Button3.Size = New-Object System.Drawing.Size(75,50)
        $Button3.Text = 'VDAAS Config Replace'
        $Button3.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $Button3.BackColor = 'Yellow'
        $Button3.Add_Click({
            $form.refresh
            $hostname = $hostnameinput.Text
            Try{
                start "CLIENT NAS LOCATION\automation\EDPAFix\Files\PSTools\PSExec.exe" -ArgumentList "-S","\\$hostname","C:\Windows\Options\Scripts\reconfigxd.cmd"
                Message-Box -Message "Configuration file replaced, $hostname rebooting." -WindowTitle "Success!" -ButtonStyle 'OKCancel'}
            Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}
        })
        $form.Controls.Add($Button3)

        ## Check DLP Version Button
        $Button4 = New-Object System.Windows.Forms.Button
        $Button4.Location = New-Object System.Drawing.Point(180,70)
        $Button4.Size = New-Object System.Drawing.Size(75,50)
        $Button4.Text = 'Check DLP Version'
        $Button4.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $Button4.BackColor = 'Yellow'
        $Button4.Add_Click({
            $form.refresh
            $hostname = $hostnameinput.Text
            Try{
                $version = Invoke-Command -ComputerName $hostname {[System.Diagnostics.FileVersionInfo]::GetVersionInfo("C:\Program Files\Symantec\Endpoint Agent\edpa.exe").FileVersion}
                IF($version -eq "15.5.0208.01004"){
                    Message-Box -Message "EDPA Agent Version: $version`r`nThis version is up to date." -WindowTitle "Result" -ButtonStyle 'OKCancel'}
                ELSEIF($version -and ($version -ne "15.5.0208.01004")){
                    Message-Box -Message "EDPA Agent Version: $version`r`nThis version is out of date. Update now." -WindowTitle "Result" -ButtonStyle 'OKCancel'}
                ELSE{
                    Message-Box -Message "Could not get EDPA Agent Version Info. Please check connectivity to VPN." -WindowTitle "Result" -ButtonStyle 'OKCancel'}
                }
            Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}
        })
        $form.Controls.Add($Button4)

        ## Force SCCM Retrieval Button
        $Button5 = New-Object System.Windows.Forms.Button
        $Button5.Location = New-Object System.Drawing.Point(10,70)
        $Button5.Size = New-Object System.Drawing.Size(75,50)
        $Button5.Text = 'Force SCCM Actions'
        $Button5.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $Button5.BackColor = 'Yellow'
        $Button5.Add_Click({
            $form.refresh
            $hostname = $hostnameinput.Text
            IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
                $form.refresh
                Message-Box -Message "`r`nForcing Software Actions on $hostname ..." -WindowTitle 'Starting' -ButtonStyle 'OKCancel'
                $form.refresh
                Try{Invoke-Command -ComputerName $hostname -EA 'SilentlyContinue' -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"}
                    $output = "Request Machine Assignments action triggered on $hostname"
                    Message-Box -Message $output -WindowTitle 'Result' -ButtonStyle 'OKCancel'
                    $form.refresh
                }
                Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}

                Try{Invoke-Command -ComputerName $hostname -EA 'SilentlyContinue' -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000022}"}
                    $output = "Machine Policy retrieval triggered on $hostname"
                    Message-Box -Message $output -WindowTitle 'Result' -ButtonStyle 'OKCancel'
                    $form.refresh
                }   
                Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}

                Try{Invoke-Command -ComputerName $hostname -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000121}"}
                    $output = "Application manager policy action triggered on $hostname"
                    Message-Box -Message $output -WindowTitle 'Result' -ButtonStyle 'OKCancel'
                    $form.refresh
                }
                Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}

                Try{Invoke-Command -ComputerName $hostname -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000113}"}
                    $output = "Scan by Update Source Action triggered on $hostname"
                    Message-Box -Message $output -WindowTitle 'Result' -ButtonStyle 'OKCancel'
                    $form.refresh
                }
                Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}

                Try{Invoke-Command -ComputerName $hostname -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000108}"}
                    $output = "Software Updates Assignments Evaluation Cycle triggered on $hostname"
                    Message-Box -Message $output -WindowTitle 'Result' -ButtonStyle 'OKCancel'
                    $form.refresh
                }
                Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}
                $output = "Force Actions Function Complete!"
                Message-Box -Message $output -WindowTitle 'Overall Result' -ButtonStyle 'OK'
                $form.refresh
        })
        $form.Controls.Add($Button5)

        ## Install DLP Fix Button
        $Button6 = New-Object System.Windows.Forms.Button
        $Button6.Location = New-Object System.Drawing.Point(180,130)
        $Button6.Size = New-Object System.Drawing.Size(75,50)
        $Button6.Text = 'Install DLP Fix'
        $Button6.BackColor = 'Yellow'
        $Button6.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $Button6.Add_Click({
            $form.refresh
            $hostname = $hostnameinput.Text
            Try{
                RoboCopy "CLIENT NAS LOCATION\Symantec_DLPFix_1-0-0-0_11" "\\$hostname\C$\Flags\Temp\Symantec_DLPFix_1-0-0-0_11" /mir /mt:32
                & "CLIENT NAS LOCATION\automation\edpafix\files\pstools\psexec.exe" -ArgumentList "-S","\\$hostname","C:\Flags\Temp\Symantec_DLPFix_1-0-0-0_11\Deploy-Application.exe"
                Message-Box -Message "Symantec_DLPFix_1-0-0-0_11 installed on $hostname" -WindowTitle "Success!" -ButtonStyle 'OKCancel'}
            Catch{
                Message-Box -Message "[$env:COMPUTERNAME] ERROR: $_" -WindowTitle 'Error' -ButtonStyle 'OK'}
        })
        $form.Controls.Add($Button6)

        ## WinRM Button
        $Button7 = New-Object System.Windows.Forms.Button
        $Button7.Location = New-Object System.Drawing.Point(10,190)
        $Button7.Size = New-Object System.Drawing.Size(75,50)
        $Button7.Text = 'Check WinRM Service'
        $Button7.BackColor = 'Yellow'
        $Button7.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $Button7.Add_Click({
            $form.refresh
            $hostname = $hostnameinput.text
            $winrm = Get-Service -Name WinRM -ComputerName $hostname | Select-Object -ExpandProperty Status
            IF($winrm -and ($winrm -ne 'Running')){Get-Service -Name WinRM -ComputerName $hostname | Set-Service -Status 'Running'
                Message-Box -Message "WinRM Service started on $hostname" -WindowTitle 'Service Started' -ButtonStyle 'OK'}
            ELSEIF(!$winrm){Message-Box -Message "Could not reach $hostname. Check connectivity." -WindowTitle 'Inaccessible Host' -ButtonStyle 'OK'}
            ELSE{Message-Box -Message "WinRM service running on $hostname already." -WindowTitle 'Service Running' -ButtonStyle 'OK'}
        })
        $form.controls.Add($Button7)

        ## Cancel Button
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(180,190)
        $CancelButton.Size = New-Object System.Drawing.Size(75,50)
        $CancelButton.Text = 'Done'
        $CancelButton.BackColor = 'Green'
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $CancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
        $form.CancelButton = $CancelButton
        $form.Controls.Add($CancelButton)

        ## Display Form
        $result = $form.ShowDialog()

        ## Output
        IF($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        $subformoutput = "Symantec Fix Applied to $hostname."
        $ListBox2.Text += "`r`n-------Symantec Fix-------`r`n$subformoutput`r`n-------End Symantec Fix-------"
        $form.refresh}
    }

Function Enable-WinRM{
    IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
    $ListBox2.Text += "`r`nAttempting to enable Windows Remote Management Service on $hostname."
    $form.refresh
    $winrm = Get-Service -Name WinRM -ComputerName $hostname | Select-Object -ExpandProperty Status
    IF($winrm -and ($winrm -ne 'Running')){
        Get-Service -Name WinRM -ComputerName $hostname | Set-Service -Status 'Running'
        $ListBox2.Text += "`r`nWinRM Service started on $hostname"}
    ELSEIF(!$winrm){
        $ListBox2.Text += "`r`nCould not reach $hostname to check WinRM Service. Check connectivity.";$form.refresh}
    ELSE{$ListBox2.Text += "`r`nWinRM service running on $hostname already.";$form.refresh}
}

Function Remove-Citrix{
    IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
    $ListBox2.Text += "`r`nAttempting to Uninstall Citrix Receiver on $hostname."
    $form.refresh
    $psexec = Test-Path "C:\Windows\System32\PSTools\psexec.exe"
    IF(!$psexec){
        RoboCopy "CLIENT NAS LOCATION\automation\edpafix\files\pstools" "C:\Windows\System32\PSTools" /mir /mt:32
        $ListBox2.Text += "`r`nPSTools not present. Copying.
        $form.refresh"
        }
    $receivercleanup = Test-Path "\\$hostname\C$\Flags\Temp\ReceiverCleanupUtility.exe"
    IF(!$receivercleanup){
        Copy-Item -Path "CLIENT NAS LOCATION\citrix\ReceiverCleanupUtility.exe" -Destination "\\$hostname\C$\Flags\Temp" -Force
        $ListBox2.Text += "`r`nCopying Receiver Cleanup Utility to $hostname."
        $form.refresh}
    $ListBox2.Text += "`r`nExecuting Receiver Cleanup Utility."
    $form.refresh
    start "C:\Windows\System32\PSTools\psexec.exe" -ArgumentList "-s","-i","\\$hostname","C:\Flags\Temp\ReceiverCleanupUtility.exe"
    $ListBox2.Text += "`r`nExecution of Receiver Cleanup Utility on $hostname complete!"
}

Function Office-Telemetry{
    IF(!$hostname){$hostname = Read-InputBoxDialog -Message 'Enter Host Name' -WindowTitle 'Hostname'}
    $ListBox2.Text += "`r`nRemoving Office Telemetry Debugger Key..."
    $form.refresh
    Invoke-Command -ComputerName $hostname {Start-ScheduledTask -TaskName 'MS_Office2016Telemetry_16-0-4266-1001_EN_11_BlockedApps'}
}
    
## Create Form (GUI)
$form = New-Object System.Windows.Forms.Form
$form.Text = 'A.R.T.T.'
$form.Size = New-Object System.Drawing.Size(910,770)
$form.Autosize = True
$form.AutoSizeMode = "GrowAndShrink"
$form.StartPosition = 'CenterScreen'

## Tool Name Label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,10)
$label.Size = New-Object System.Drawing.Size(360,20)
$label.Text = 'Advanced Remote Troubleshooting Tool (ARTT):'
$label.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($label)

## Pick a Function Label
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,30)
$textBox.Size = New-Object System.Drawing.Size(340,20)
$textBox.Text = 'Pick a Function'
$textBox.TextAlign = 'Center'
$textBox.BackColor = 'LightBlue'
$textBox.ReadOnly = $true
$textBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($textBox)

## HostName Label
$textBox6 = New-Object System.Windows.Forms.TextBox
$textBox6.Location = New-Object System.Drawing.Point(365,30)
$textBox6.Size = New-Object System.Drawing.Size(500,60)
$textBox6.Text = 'Host Name:'
$textBox6.TextAlign = 'Center'
$textBox6.BackColor = 'LightBlue'
$textBox6.ReadOnly = $true
$textBox6.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($textBox6)

## HostName Input Box
$textBox7 = New-Object System.Windows.Forms.TextBox
$textBox7.Location = New-Object System.Drawing.Point(365,62)
$textBox7.Size = New-Object System.Drawing.Size(500,60)
$textBox7.TextAlign = 'Center'
$textBox7.BackColor = 'Red'
$textBox7.ReadOnly = $false
$textBox7.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($textBox7)

## Text Box Function Description Label
$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(365,140)
$textBox3.Size = New-Object System.Drawing.Size(500,60)
$textBox3.Text = 'Function Description'
$textBox3.TextAlign = 'Center'
$textBox3.BackColor = 'LightBlue'
$textBox3.ReadOnly = $true
$textBox3.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($textBox3)

## Output Text Box Label
$textBox5 = New-Object System.Windows.Forms.TextBox
$textBox5.Location = New-Object System.Drawing.Point(365,368)
$textBox5.Size = New-Object System.Drawing.Size(500,60)
$textBox5.Text = 'Output Box'
$textBox5.TextAlign = 'Center'
$textBox5.BackColor = 'LightBlue'
$textBox5.ReadOnly = $true
$textBox5.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($textBox5)

## Function List Box
$ListBox = New-Object System.Windows.Forms.ListBox
$ListBox.Location = New-Object System.Drawing.Point(10,60)
$ListBox.Size = New-Object System.Drawing.Size(340,600)
$ListBox.Height = 600
[void] $ListBox.Items.Add('APPLICATIONS: Show List of All Installed Applications')
[void] $ListBox.Items.Add('APPLICATIONS: Remove Application')
[void] $ListBox.Items.Add('APPLICATIONS: Install Application from Backup')
[void] $ListBox.Items.Add('APPLICATIONS: Get Java Version')
[void] $ListBox.Items.Add('APPLICATIONS: Clear Office Document Cache')
[void] $ListBox.Items.Add('APPLICATIONS: Launch Symantec Fix GUI')
[void] $ListBox.Items.Add('APPLICATIONS: Remove Citrix Receiver')
[void] $ListBox.Items.Add(' ')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Enable WinRM Service')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: View Available HD Space')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Full Scan / Repair OS')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Full System Cleanup')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Enable UAC')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Side-by-Side')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: View OS Version')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Check Image Date')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Register Font Permissions')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Set Network Connection Profile')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Activate Office')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: Activate Windows')
[void] $ListBox.Items.Add('SYSTEM FUNCTIONS: View User and Uptime')
[void] $ListBox.Items.Add(' ')
[void] $ListBox.Items.Add('SOFTWARE CENTER: Force Actions')
[void] $ListBox.Items.Add('SOFTWARE CENTER: View CCM Cache Allocation.')
[void] $ListBox.Items.Add('SOFTWARE CENTER: Increase CCM Cache Allocation.')
[void] $ListBox.Items.Add('SOFTWARE CENTER: Clear CCM Cache.')
[void] $ListBox.Items.Add('SOFTWARE CENTER: Reinstall Software Center Client.')
[void] $ListBox.Items.Add(' ')
[void] $ListBox.Items.Add('ACCOUNT FUNCTIONS: Rebuild Citrix Profile')
[void] $ListBox.Items.Add('ACCOUNT FUNCTIONS: Get-ADGroups')
[void] $ListBox.Items.Add('ACCOUNT FUNCTIONS: Unlock/Reset-Password')
[void] $ListBox.Items.Add(' ')
[void] $ListBox.Items.Add('FILE FUNCTIONS: Copy a directory')
[void] $ListBox.Items.Add('FILE FUNCTIONS: Compare two directories')
[void] $ListBox.Items.Add('FILE FUNCTIONS: Open C$ Share')
[void] $ListBox.Items.Add(' ')
[void] $ListBox.Items.Add('CONSOLE FUNCTIONS: Powershell Console')
[void] $ListBox.Items.Add('CONSOLE FUNCTIONS: Check Service Status')
$listBox.add_SelectedIndexChanged($SelectedFile)
$listBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($ListBox)

## Function Description Output Box
$textBox4 = New-Object System.Windows.Forms.TextBox
$textBox4.Location = New-Object System.Drawing.Point(365,170)
$textBox4.Size = New-Object System.Drawing.Size(500,75)
$textBox4.Height = 90
$textBox4.Multiline = $True
$textBox4.ReadOnly = $True
$textBox4.WordWrap = $True
$textBox4.ScrollBars = 'Vertical'
$textBox4.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($textBox4)

## Function Output Box
$listBox2 = New-Object System.Windows.Forms.TextBox
$listBox2.Location = New-Object System.Drawing.Point(365,397)
$listBox2.Size = New-Object System.Drawing.Size(500,600)
$listBox2.Height = 252
$listBox2.Multiline = $True
$listBox2.ReadOnly = $True
$listBox2.WordWrap = $True
$listBox2.Scrollbars = 'Vertical'
$listBox2.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($listBox2)

## Description Button
$DescriptionButton = New-Object System.Windows.Forms.Button
$DescriptionButton.Location = New-Object System.Drawing.Point(340,680)
$DescriptionButton.Size = New-Object System.Drawing.Size(150,46)
$DescriptionButton.Text = 'Description'
$DescriptionButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$DescriptionButton.BackColor = 'Yellow'

## Description Button Functions
$DescriptionButton.Add_Click({
    IF ($ListBox.SelectedItem -eq 'APPLICATIONS: Show List of All Installed Applications'){
        $textBox4.Text = "Show List of All Installed Applications.`r`n `r`nShows full sorted list of all applications installed on remote machine."}
    IF ($ListBox.SelectedItem -eq 'SYSTEM FUNCTIONS: View Available HD Space'){
        $textBox4.Text = "View Available HD Space `r`n `r`nShows sorted HD Space on remote machine."}
    IF ($ListBox.SelectedItem -eq 'SOFTWARE CENTER: Reinstall Software Center Client.'){
        $textBox4.Text = "Reinstall Software Center Client. `r`n `r`nRemoves and reinstalls the Automated Software Delivery Client on a Remote Machine."}
    IF ($ListBox.SelectedItem -eq 'SOFTWARE CENTER: View CCM Cache Allocation.'){
        $textBox4.Text = "View CCM Cache Allocation. `r`n `r`nShows current Automated Software Delivery Cache Size."}
    IF ($ListBox.SelectedItem -eq 'SOFTWARE CENTER: Increase CCM Cache Allocation.'){
        $textBox4.Text = "Set CCM Cache Allocation to 50GB. `r`n `r`nSets the Automated Software Delivery Cache Size to 50GB (twice the default). This is best used for large deployments and Upgrades in Place."}
    IF ($ListBox.SelectedItem -eq 'SOFTWARE CENTER: Clear CCM Cache.'){
        $textBox4.Text = "Clear CCM Cache. `r`n `r`nClears the Automated Software Delivery Cache. *Use only in very specific circumstances!* Should only be used for Upgrades in Place, and as a last ditch effort for space management."}
    IF ($ListBox.SelectedItem -eq 'ACCOUNT FUNCTIONS: Rebuild Citrix Profile'){
        $textBox4.Text = "Rebuild Citrix Profile. `r`n `r`nRenames Citrix User Directory on Citrix Servers to force Profile Recreation."}
    IF ($ListBox.SelectedItem -eq 'APPLICATIONS: Remove Application'){
        $textBox4.Text = "Remove Application. `r`n `r`nRemoves Application from Remote Machine matching the name entered. *Use cautiously, as it will remove all applications with that name.*`r`nDuring software removal it will ask permission for each removal.`r`nNOTE: Entering End at the prompt will end the script with no further action."}
    IF ($ListBox.SelectedItem -eq 'APPLICATIONS: Install Application from Backup'){
        $textBox4.Text = "Install Application from Backup. `r`n `r`nInstalls Application listed from ATS Backup Share on Remote Machine. Usage instructions: FileName must be exact as it appears in folder 'CLIENT NAS LOCATION'. Will only work with applications that have a Deploy-Application package. Installation log is not displayed on this GUI, but is saved to the remote machines Flags folder."}
    IF ($ListBox.SelectedItem -eq 'APPLICATIONS: Launch Symantec Fix GUI'){
        $textBox4.Text = "Launch Symantec Fix GUI. `r`n`r`nLaunches custom Symantec GUI with options for single workstation fixes for DLP hotfix problems."}
    IF ($ListBox.SelectedItem -eq 'SYSTEM FUNCTIONS: Full Scan / Repair OS'){
        $textBox4.Text = "Full Scan / Repair OS. `r`n `r`nRuns SFC /scannow on a remote machine. NOTE: This is running a lengthy process over the network - it is *not* fast. Results will appear after it has finished executing, but often take quite some time. Full log of the scan located at '%windir%\Logs\CBS\CBS.log' on remote machine."}
    IF ($ListBox.SelectedItem -eq 'SYSTEM FUNCTIONS: Enable UAC'){
        $textBox4.Text = "Enable UAC. `r`n `r`nEnables User Account Controls on remote machine. NOTE: Requires reboot to complete."}
    IF ($ListBox.SelectedItem -eq 'SYSTEM FUNCTIONS: Full System Cleanup'){
        $textBox4.Text = "Full System Cleanup. `r`n `r`nExecutes the System Cleanup Tool on a Remote Machine, with admin privileges."}
    IF ($ListBox.SelectedItem -eq 'APPLICATIONS: Clear Office Document Cache'){
        $textBox4.Text = "Clear Office Document Cache. `r`n `r`nClears the Office Document Cache on a Remote Machine. Note: Requires Reboot."}
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Check Image Date'){
        $textBox4.Text = "Check Image Date. `r`n `r`nDetermines the original date at which machine was imaged. Can often lead to determining if a machine is due for refresh."}
    IF ($ListBox.SelectedItem -Eq 'ACCOUNT FUNCTIONS: Get-ADGroups'){
        $textBox4.Text = "Get AD Groups. `r`n `r`nShows full list of all Active Directory Security Groups of which an indicated user account is a member."}
    IF ($ListBox.SelectedItem -Eq 'ACCOUNT FUNCTIONS: Unlock/Reset-Password'){
        $textBox4.Text = "Unlock / Reset AD Password. `r`n `r`nAllows the simple unlock, or full unlock and reset of a user's Active Directory Password. Reset will prompt for new password."}
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: View OS Version'){
        $textBox4.Text = "View OS Versioning. `r`n `r`nShows Operating System, Version, and Build."}
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Set Network Connection Profile'){
        $textBox4.Text = "Set Network Connection Profile. `r`n `r`nSets current network connection profile type to Private. `r`n `r`nNOTE: Domain Authenticated profiles (ad.wellpoint.com direct connections) cannot be changed and will just display 'Domain Authenticated'."}
    IF ($ListBox.SelectedItem -Eq 'FILE FUNCTIONS: Compare two directories'){
        $textBox4.Text = "Compare two directories. `r`n `r`nCompares two designated file locations and displays each file that is in the first and missing from the second."}
    IF ($ListBox.SelectedItem -Eq 'FILE FUNCTIONS: Copy a directory'){
        $textBox4.Text = "Copy a directory. `r`n `r`nCopies a directory from a source location to a destination location. Uses multi-threading to copy the file, so this method is faster than many."}
    IF ($ListBox.SelectedItem -Eq 'SOFTWARE CENTER: Force Actions'){
        $textBox4.Text = "Force SCCM Actions. `r`n `r`nForces all software inventory and retrieval actions on a remote machine. Best used for immediately installing recently pushed software, and updating applications and OS."}
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Register Font Permissions'){
        $textBox4.Text = "Register Font Permissions.`r`n `r`nCollects file permissions associated with standard fonts and applies it to all fonts installed. This resolves a particular issue caused by upgrades in places from Windows 7 to Windows 10 on multiple machines."}
    IF ($ListBox.SelectedItem -Eq 'APPLICATIONS: Get Java Version'){
        $textBox4.Text = "Get Java Version. `r`n `r`nShows the versions of all copies of the Java Runtime Environment currently installed on a remote machine."}
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: View User and Uptime'){
        $textBox4.Text = "View User and Uptime. `r`n `r`nShows currently logged on user, current state of user, current session logon time, and the time of last system boot."}
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Side-by-Side'){
        $textBox4.Text = "Side-by-Side. `r`n `r`nRemoves the Powershell Configuration file on a remote machine and forces the creation of a new one, for the purpose of resolving side-by-side errors in Powershell and SCCM primarily."}
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Activate Office'){
        $textBox4.Text = "Activate Office. `r`n `r`nRegisters and activates an Enterprise copy of Microsoft Office (works with 2013 and 2016)."}
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Activate Windows'){
        $textBox4.Text = "Activate Windows. `r`n `r`nRegisters and activates an Enterprise copy of Microsoft Windows (works with 7 and 10)."}
    IF ($ListBox.SelectedItem -Eq 'CONSOLE FUNCTIONS: Powershell Console'){
        $textBox4.Text = "Powershell Console. `r`n `r`nOpens a GUI to allow manual input of PowerShell commands."}
    IF ($ListBox.SelectedItem -Eq 'FILE FUNCTIONS: Open C$ Share'){
        $textBox4.Text = "Open C$ Share.`r`n`r`nOpens C$ share on host to view backend files."}
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Enable WinRM Service'){
        $textBox4.Text = "Enable WinRM Service.`r`n`r`nCheck Windows Remote Management Service on remote machine, and if not running enable it."}
    IF ($ListBox.SelectedItem -Eq 'APPLICATIONS: Remove Citrix Receiver'){
        $textBox4.Text = "Remove Citrix Receiver.`r`n`r`nTransfers and executes the Citrix Receiver Removal Tool on the remote machine interactively. Can be reinstalled from Application Management Window."}
    $form.refresh
        })
$form.Controls.Add($DescriptionButton)

## Execute Button
$ExecuteButton = New-Object System.Windows.Forms.Button
$ExecuteButton.Location = New-Object System.Drawing.Point(75,680)
$ExecuteButton.Size = New-Object System.Drawing.Size(150,46)
$ExecuteButton.Text = 'Execute'
$ExecuteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$ExecuteButton.BackColor = 'Red'

## Execute Button Functions
$ExecuteButton.Add_Click({
     $curruser = (Get-WMIObject -Class win32_Process | where {$_.processname -match "explorer.exe"}).getowner().user
     $form.refresh
     $hostname = $textbox7.Text
## View Available HD Space
     IF ($ListBox.SelectedItem -eq 'SYSTEM FUNCTIONS: View Available HD Space')
     {
         View-HDSpace
         $form.refresh
     }
## Show List of All Installed Applications
     IF ($ListBox.SelectedItem -eq 'APPLICATIONS: Show List of All Installed Applications')
     {
        Call-ApplicationManagement
        $form.refresh
     }
## Reinstall Software Center Client
     IF ($ListBox.SelectedItem -eq 'SOFTWARE CENTER: Reinstall Software Center Client.')
     {
         Reinstall-SCCM
         $form.refresh
      } 
## View CCM Cache Allocation              
     IF ($ListBox.SelectedItem -eq 'SOFTWARE CENTER: View CCM Cache Allocation.')
     {
         View-CCMCache
         $form.refresh
     }
## Set CCM Cache Allocation to 50GB
     IF ($ListBox.SelectedItem -eq 'SOFTWARE CENTER: Increase CCM Cache Allocation.')
     {
         Increase-CCMCache
         $form.refresh
     }
## Clear CCM Cache
     IF ($ListBox.SelectedItem -eq 'SOFTWARE CENTER: Clear CCM Cache.')
     {
         Clear-CCMCache
         $form.refresh
     }
## Rebuild Citrix Profile
     IF ($ListBox.SelectedItem -eq 'ACCOUNT FUNCTIONS: Rebuild Citrix Profile')
     {
         Rename-Citrix
         $form.refresh
     }
## Remove Application
     IF ($ListBox.SelectedItem -eq 'APPLICATIONS: Remove Application')
    {
        Call-ApplicationManagement
        $form.refresh
    }
## Install Application from Backup
     IF ($ListBox.SelectedItem -eq 'APPLICATIONS: Install Application from Backup')
     {
        Call-ApplicationManagement
        $form.refresh
     }
## Full System Scan / Repair OS
     IF ($ListBox.SelectedItem -eq 'SYSTEM FUNCTIONS: Full Scan / Repair OS')
     {
        ScanRepair-OS
        $form.refresh
     }
## Enable UAC
     IF ($ListBox.SelectedItem -eq 'SYSTEM FUNCTIONS: Enable UAC')
     {
        Enable-UAC
        $form.refresh
     } 
## Full System Cleanup
     IF ($ListBox.SelectedItem -eq 'SYSTEM FUNCTIONS: Full System Cleanup')
     {
        System-Cleanup
        $form.refresh
     }
## Clear Office Document Cache
     IF ($ListBox.SelectedItem -eq 'APPLICATIONS: Clear Office Document Cache')
     {
        Clear-OfficeCache
        $form.refresh 
     } 
## Check Image Date 
     IF ($ListBox.SelectedItem -eq 'SYSTEM FUNCTIONS: Check Image Date')
     {
        Check-ImageDate
        $form.refresh
     }
## Get-ADGroup Memberships
    IF ($ListBox.SelectedItem -eq 'ACCOUNT FUNCTIONS: Get-ADGroups')
    {
        Get-ADGroups
        $form.refresh
    }
## Unlock/Reset AD Password
    IF ($ListBox.SelectedItem -eq 'ACCOUNT FUNCTIONS: Unlock/Reset-Password')
    {
        Unlock-Reset
        $form.refresh
    }
## View OS Version
    IF ($listBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: View OS Version')
    {
        View-OSVersion
        $form.refresh
    }
## Set Network Connection Profile for current adapter to Private
    IF ($listbox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Set Network Connection Profile')
    {
        Set-NetProfile
        $form.refresh
    }
## Compare two directories
    IF ($listbox.SelectedItem -Eq 'FILE FUNCTIONS: Compare two directories')
    {
        Compare-directories
        $form.refresh
    }
## Copy a directory
    IF ($ListBox.SelectedItem -Eq 'FILE FUNCTIONS: Copy a directory')
    {
        Copy-Directory
        $form.refresh
    }
## Run Software Center Actions
    IF ($listbox.SelectedItem -Eq 'SOFTWARE CENTER: Force Actions')
    {
        Force-Action
        $form.refresh
    }
## Re-Register Font Permissions
    IF ($listBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Register Font Permissions')
    {
        Font-Fix
        $form.refresh
    }
## View currently installed Java REs
    IF ($listBox.SelectedItem -Eq 'APPLICATIONS: Get Java Version')
    {
        Get-Java
        $form.refresh
    }
## View Logged on User and Uptime
    IF ($listBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: View User and Uptime')
    {
        Get-User
        $form.refresh
    }
## Side-By-Side
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Side-by-Side')
    {
        Side-By-Side
        $form.refresh
    }
## Activate Office
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Activate Office')
    {
        Activate-Office
        $form.refresh
    }
## Activate Windows
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Activate Windows')
    {
        Activate-Windows
        $form.refresh
    }
    IF($ListBox.SelectedItem -Eq 'CONSOLE FUNCTIONS: Powershell Console'){
        PS-GUI
        $form.refresh
    }
    IF($ListBox.SelectedItem -Eq 'CONSOLE FUNCTIONS: Check Service Status'){
        Service-Status
        $form.refresh
    }
## Open C$ Share
    IF ($ListBox.SelectedItem -Eq 'FILE FUNCTIONS: Open C$ Share')
    {
        C-Share
        $form.refresh
    }
## Launch Symantec Fix GUI
    IF ($ListBox.SelectedItem -Eq 'APPLICATIONS: Launch Symantec Fix GUI'){
        EDPA
        $form.refresh
    }
## Enable WinRM Service
    IF ($ListBox.SelectedItem -Eq 'SYSTEM FUNCTIONS: Enable WinRM Service'){
        Enable-WinRM
        $form.refresh
    }
## Remove Citrix Receiver
    IF ($ListBox.SelectedItem -Eq 'APPLICATIONS: Remove Citrix Receiver'){
        Remove-Citrix
        $form.refresh
    }
## Logging
    $form.refresh
    $date = Get-Date
    $form.refresh
    $logging = New-Object PSObject
    $logging | Add-Member NoteProperty Date $date
    $logging | Add-Member NoteProperty Function $function
    $logging | Add-Member NoteProperty Hostname $hostname
    $logging | Add-Member NoteProperty Output $output
    $logging | Export-CSV "C:\Flags\Temp\ARTTLog.csv" -NoTypeInformation -Append
    IF($output){$output = 'Function Completed Successfully.'}
    ELSE{$output = 'Function did not produce output.'}
    $toollogging = New-Object PSObject
    $toollogging | Add-Member NoteProperty User $curruser
    $toollogging | Add-Member NoteProperty Date $date
    $toollogging | Add-Member NoteProperty Hostname $hostname
    $toollogging | Add-Member NoteProperty Output $output
    $toollogging | Export-CSV "\\VDAASW1015915\C$\ARTT\Files\Logging\$function.csv" -NoTypeInformation -Append
})
$form.Controls.Add($ExecuteButton)

## Cancel Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(600,680)
$CancelButton.Size = New-Object System.Drawing.Size(150,46)
$CancelButton.Text = 'Done'
$CancelButton.BackColor = 'Green'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$CancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

## Display Form
$result = $form.ShowDialog()