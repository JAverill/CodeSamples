## Encryption Verification / New Asset Checklist Script
## Version 1.5 - Author: John Averill, IBM
## .5 Update: Tooltips and navigation assistance added.

## Overview: A script utilized by all Client Field Support Imaging Technicians on newly imaged assets to ensure all encryption is in place, software is installed, and the image process is complete. It is intended to be used on groups of machines that have been freshly imaged. Field Support estimates a total of 120 hours are spent performing this process each month, and 60 hours of savings (50% manhour reduction) is the project goal, which has been confirmed met by this solution.

## Load Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Net
Add-Type -AssemblyName System.Management.Automation

## Declare Global Functions and variables

## Display Data Functions
Function Read-InputBoxDialog([string]$Message, [string]$WindowTitle)
    {return [Microsoft.VisualBasic.Interaction]::InputBox($Message, $WindowTitle)}

Function Message-Box([string]$Message,[string]$WindowTitle,[string]$ButtonStyle)
    {return [System.Windows.MessageBox]::Show($Message,$WindowTitle,$ButtonStyle)}

## Set Error Action Preference to avoid interruption to GUI
$ErrorActionPreference = "silentlycontinue"

## Force SCCM Updates Function
Function Force-Actions
{
    $global:sccmoutput = [system.collections.arraylist]@{}
    IF(Invoke-Command -ComputerName $global:hostname -EA 'SilentlyContinue' -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"})
    {
        $global:sccmoutput += "Request Machine Assignments action triggered on $global:hostname."
    }
    ELSE
    {
        $global:sccmoutput += "Could not reach SMS Client on $global:hostname to trigger Request Machine Assignments action."
    }
    IF(Invoke-Command -ComputerName $global:hostname -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000121}"})
    {
        $global:sccmoutput += "Application manager policy action triggered on $global:hostname."
    }
    ELSE
    {
        $global:sccmoutput += "Could not reach SMS Client on $global:hostname to trigger Application manager policy action."
    }
    IF(Invoke-Command -ComputerName $global:hostname -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000113}"})
    {
        $global:sccmoutput += "Scan by Update Source Action triggered on $hostname."
    }
    ELSE
    {
        $global:sccmoutput += "Could not reach SMS Client on $hostname to trigger Scan by Update Source action."
    }
    IF(Invoke-Command -ComputerName $hostname -ScriptBlock{Invoke-WMIMethod -NameSpace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000108}"})
    {
        $global:sccmoutput += "Software Updates Assignments Evaluation Cycle triggered on $hostname."
    }
    ELSE
    {
        $global:sccmoutput += "Could not reach SMS Client on $hostname to trigger Software Updates Assignments Evaluation Cycle."
    }
    $global:sccmsuccess = ($global:sccmoutput | Where-Object {$_ -like "*triggered*"}).count
    $global:sccmfailure = ($global:sccmoutput | Where-Object {$_ -like "Could not*"}).count
}

## Get Info:
Function Get-Info
{
    ## Confirm WinRM Running and enable if not
    $winrm = Get-Service -Name WinRM -ComputerName $global:hostname | Select-Object -ExpandProperty Status
    IF($winrm -and ($winrm -ne 'Running'))
    {
        Get-Service -Name WinRM -ComputerName $global:hostname | Set-Service -Status 'Running'
    }
    
    ## Check for McAfee Products
    $data = Get-WMIOBject -Class win32_product -ComputerName $global:hostname

    ## Endpoint Encryption
    $global:mcafeeproducts = [system.collections.arraylist]@{}
    $global:mcafeeproducts = $data | Where-Object {$_.Name -like "*McAfee*"} | Select-Object -ExpandProperty Name
    $global:mcafeecount = $global:mcafeeproducts.count

    ## Check for Office
    $global:office = $data | Where-Object {$_.Name -like "Microsoft Office Professional*"} | Select-Object -ExpandProperty Name
    $global:officever = $data | Where-Object {$_.Name -like "Microsoft Office Professional*"} | Select-Object -ExpandProperty Version

    ## Check for Pulse Secure
    $global:pulse = $data | Where-Object {$_.Name -like "Pulse Secure"} | Select-Object -ExpandProperty Name
    $global:psversion = $data | Where-Object {$_.Name -like "Pulse Secure"} | Select-Object -ExpandProperty Version

    ## Gpupdate / Check Machine Cert
    $global:policyupdate = Invoke-Command -ComputerName $global:hostname {gpupdate /target:computer /force}
    $global:certs = Invoke-Command -ComputerName $global:hostname {Get-ChildItem -path cert:\LocalMachine\My}
    
    ## Force SCCM Updates
    Force-Actions

    ## Check for Bitlocker Encryption
    $bde = Invoke-Command -ComputerName $global:hostname {Manage-BDE -Status C:}
    $status = $bde | Select-String -Pattern "Conversion Status:"
    IF($status -like "*Fully Encrypted*")
    {
        $global:bitlocker = "Encrypted"
    }
    ELSE
    {
        $global:bitlocker = "NOT Encrypted"
    }

    ## Check Image Version
    $key = Invoke-Command -ComputerName $global:hostname {Get-ItemProperty -Path "HKLM:\SYSTEM\AssociateTechnologyManagement" -Name 'EnterpriseImageVersion'}
    $global:imageversion = $key.EnterpriseImageVersion
    IF(!$global:imageversion){$global:imageversion = 'Registry Inaccessible'}

    ## Check for System Info Icon - Flag for Reimage if not present
    $global:sysinfo = Test-Path "\\$global:hostname\C$\Users\Public\Desktop\System Info.lnk"
}

## Log to CSV
Function Log-ForCSV
{
    $logging = New-Object PSObject
    $logging | Add-Member NoteProperty Hostname $global:hostname
    $logging | Add-Member NoteProperty Microsoft_Office "$global:office - $global:officever"
    $logging | Add-Member NoteProperty Pulse_Secure "$global:pulse - $global:psversion"
    $logging | Add-Member NoteProperty Updates_Forced "$global:sccmsuccess Update cycles Succeeded"
    $logging | Add-Member NoteProperty BitLocker_Status $global:bitlocker
    $logging | Add-Member NoteProperty Image_Version $global:imageversion
    $logging | Add-Member NoteProperty Icons_Created $global:sysinfo
    IF($global:mcafeeproducts){
    $z = 7
    $a = 0
    ForEach($item in $global:mcafeeproducts)
    {
        $z = $z + 1
        $a = $a + 1
        $itemname = $item.name
        $logging | Add-Member NoteProperty McAfee_Product_$a $itemname
    }
}
    $global:loggingarray += $logging
}
Function Log-ToCSV
{
    $global:loggingarray | Export-CSV -Path $global:exportpath -NoTypeInformation -Append
}

## Log to DataGrid
Function Log-DataGrid
{
    ## Content Boolean Switches
    IF($global:office){$global:ms = 'âˆš Confirmed'}
    ELSE{$global:ms = 'X FAIL'}
    IF($global:pulse){$global:ps = 'âˆš Confirmed'}
    ELSE{$global:ps = 'X FAIL'}
    IF($global:policyupdate -like "*successfully*" -and $global:certs -like "*us.ad.wellpoint.com*"){$global:gp = 'âˆš Confirmed'}
    ELSE{$global:gp = 'X FAIL'}
    IF($global:sccmsuccess -eq 4){$global:sc = 'âˆš Confirmed'}
    ELSE{$global:sc = 'X FAIL'}
    IF($global:bitlocker -eq 'Encrypted'){$global:bl = 'âˆš Confirmed'}
    ELSE{$global:bl = 'X FAIL'}
    IF($global:sysinfo = $True){$global:si = 'âˆš Confirmed'}
    ELSE{$global:si = 'X FAIL'}
    $global:row = @($global:hostname,$global:ms,$global:ps,$global:gp,$global:sc,$global:bl,$global:imageversion,$global:si)
    IF($global:mcafeeproducts){
    $z = 7
    $a = 0
    ForEach($item in $global:mcafeeproducts)
    {
        $z = $z + 1
        $global:row += $item
    }
}
    $outputbox.Rows.Add($global:row)
    $form.refresh()
}

## Create ToolTips
Function ToolTip-Cell{
$tooltip = New-Object System.Windows.Forms.ToolTip
$showhelp ={
    Switch ($this.name){
    "SCCM" {$tip = "Open SCCM Collection Manager to install software."}
    "export" {$tip = "Export CSV log of script run to chosen location."}
    "CheckMachines" {$tip = "Execute script on all listed machine.`r`nNOTE: The batch will take approximately 1 - 2 minutes per machine on list."}
    "hostlabel" {$tip = "Select method of entering list of hostnames to execute script on."}
    }
    $tooltip.SetToolTip($this,$tip)}
}

## Create GUI
## Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'A.I.C.'
$form.Size = New-Object System.Drawing.Size(680,500)
$form.Autosize = $False
$form.AutoSizeMode = "GrowAndShrink"
$form.StartPosition = 'CenterScreen'

## Tool Name Label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,10)
$label.Size = New-Object System.Drawing.Size(250,20)
$label.Text = "Asset Imaging Checklist"
$label.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($label)

## Filename Label
$hostlabel = New-Object System.Windows.Forms.Label
$hostlabel.Location = New-Object System.Drawing.Point(10,40)
$hostlabel.Size = New-Object System.Drawing.Size(150,20)
$hostlabel.Text = "Select Method of Input:"
$hostlabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.controls.add($hostlabel)

## Hostname Input
$optionsarray = @("Single Hostname","List of Hostnames","Comma-Separated List","List from CSV Path")
$machineinput = New-Object System.Windows.Forms.ComboBox
$machineinput.Location = New-Object System.Drawing.Point(370,40)
$machineinput.Name = "hostlabel"
$machineinput.add_MouseHover($ShowHelp)
$machineinput.Size = New-Object System.Drawing.Size(110,20)
$machineinput.DropDownHeight = 110
$machineinput.BackColor = 'Red'
$machineinput.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
ForEach($option in $optionsarray)
    {
        $machineinput.Items.Add($option)
    }
$form.Controls.Add($machineinput)

## Single Host Entry
$hostnameinput = New-Object System.Windows.Forms.TextBox
$hostnameinput.Location = New-Object System.Drawing.Point(490,40)
$hostnameinput.Size = New-Object System.Drawing.Size(110,20)
$hostnameinput.TextAlign = 'Center'
$hostnameinput.BackColor = 'Red'
$hostnameinput.ReadOnly = $false
$form.Controls.Add($hostnameinput)

## Machine Output box
$outputbox = New-Object System.Windows.Forms.DataGridView
$outputbox.Location = New-Object System.Drawing.Point(10,70)
$outputbox.Size = New-Object System.Drawing.Size(650,300)
$outputbox.BackColor = 'LightBlue'
$outputbox.Name = "outputbox"
$outputbox.add_MouseHover($ShowHelp)
$outputbox.ScrollAlwaysVisible = $true
$outputbox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$y = (8 + $global:mcafeecount)
$outputbox.ColumnCount = $y
$outputbox.ColumnHeadersVisible = $true
$outputbox.Columns[0].Name = "Hostname"
$hostnamecolumn = $outputbox.Columns[0]
$hostnameheader = $hostnamecolumn.HeaderCell
$hostnameheader.ToolTipText = "Asset Host (Computer) Name."
$outputbox.Columns[1].Name = "Microsoft_Office"
$officecolumn = $outputbox.Columns[1]
$officeheader = $officecolumn.HeaderCell
$officeheader.ToolTipText = "CONFIRMED: Microsoft Office Professional is installed on this asset.`r`nFAIL: Microsoft Office Professional is not installed on this asset."
$outputbox.Columns[2].Name = "Pulse_Secure"
$pulsecolumn = $outputbox.Columns[2]
$pulseheader = $pulsecolumn.HeaderCell
$pulseheader.ToolTipText = "CONFIRMED: Pulse Secure is installed on this asset.`r`nFAIL: Pulse Secure is not installed on this asset."
$outputbox.Columns[3].Name = "GP_Update_Performed"
$gpupdatecolumn = $outputbox.Columns[3]
$gpupdateheader = $gpupdatecolumn.HeaderCell
$gpupdateheader.ToolTipText = "CONFIRMED: Local Machine Policy Update was performed successfully, and the Local Machine certificate's presence and recent update has been confirmed.`r`nFAIL: Either the Local Machine Policy update failed, or the cert was not updated."
$outputbox.Columns[4].Name = "Updates_Forced"
$updatescolumn = $outputbox.Columns[4]
$updatesheader = $updatescolumn.HeaderCell
$updatesheader.ToolTipText = "CONFIRMED: SCCM Updates were forced on the asset and succeeded.`r`nFAIL: One or more components of Software Center (SCCM Client) is not responding on the remote machine. SCCM Troubleshooting may be required."
$outputbox.Columns[5].Name = "BitLocker_Encryption"
$bitlockercolumn = $outputbox.Columns[5]
$bitlockerheader = $bitlockercolumn.HeaderCell
$bitlockerheader.ToolTipText = "CONFIRMED: Bitlocker is enabled and HD is fully enabled on this asset.`r`nFAIL: Bitlocker is disabled, or HD is not fully encrypted."
$outputbox.Columns[6].Name = "Image_Version"
$imagecolumn = $outputbox.Columns[6]
$imageheader = $imagecolumn.HeaderCell
$imageheader.ToolTipText = "Shows current image version.`r`nRegistry Inaccessible/BLANK: Indicates the remote registry on this asset is not responding for some reason which may indicate a number of different problems."
$outputbox.Columns[7].Name = "Complete_Image"
$incompletecolumn = $outputbox.Columns[7]
$incompleteheader = $incompletecolumn.HeaderCell
$incompleteheader.ToolTipText = "CONFIRMED: Presence of Public shortcuts (final step in imaging) confirmed, imaging complete.`r`nFAIL: Imaging incomplete."
IF($mcafeeproducts){
    $z = 7
    $a = 0
    ForEach($item in $mcafeeproducts)
    {
        $a = $a + 1
        $z = $z + 1
        $outputbox.Columns[$z].Name = "McAfee_Product_$a"
        $mcafeecolumn = $outputbox.columns[$z]
        $mcafeecolumn.HeaderCell
        $mcafeecolumn.ToolTipText = "McAfee Product Detected and confirmed installed. Number $a in list of $global:mcafeecount."
    }
}
$form.controls.add($outputbox)

## Check Machines Button
$Checkmachines = New-Object System.Windows.Forms.Button
$Checkmachines.Location = New-Object System.Drawing.Point(500,390)
$Checkmachines.Size = New-Object System.Drawing.Size(50,50)
$Checkmachines.Text = 'Check!'
$Checkmachines.Name = "CheckMachines"
$Checkmachines.add_MouseHover($ShowHelp)
$Checkmachines.BackColor = 'LightBlue'
$Checkmachines.Add_Click({
    $global:loggingarray = [system.collections.arraylist]@{}
    $form.refresh
    ## Clear variables for use.
    $global:hostname = $null
    $global:ip = $null

    ## Single Hostname Switch
    IF($machineinput.SelectedItem -eq "Single Hostname")
    {
        $global:hostname = $hostnameinput.Text
        Message-Box -Message "Checking $global:hostname!" -WindowTitle 'Checking!' -ButtonStyle 'OK'
        Get-Info
        Log-DataGrid
        Log-ForCSV
        Message-Box -Message "Complete and written to Data Grid!" -WindowTitle 'Done!' -ButtonStyle 'OK'
    }

    ## Multiple Hostname text file Switch
    IF($machineinput.SelectedItem -eq "List of Hostnames")
    {
        $path = $hostnameinput.Text
        $global:hostnames = Get-Content $path
        Message-Box -Message "Checking all machines! Please wait." -WindowTitle 'Checking!' -ButtonStyle 'OK'
        ForEach($global:hostname in $global:hostnames)
        {
            Get-Info
            Log-DataGrid
            Log-ForCSV
            $global:hostname = $null
            $global:ip = $null
            $global:imageversion = $null
        }
        Message-Box -Message "Complete and written to Data Grid!" -WindowTitle 'Done!' -ButtonStyle 'OK'
    }

    ## Multiple Hostname Comma-Delineated List
    IF($machineinput.SelectedItem -eq "Comma-Separated List")
    {
        $itemsarray = $hostnameinput.Text
        $splitarray = $itemsarray.split(',')
        $global:hostnames = [system.collections.arraylist]@{}
        ForEach($item in $splitarray){$global:hostnames += $item}
        Message-Box -Message "Checking all machines! Please wait." -WindowTitle 'Checking!' -ButtonStyle 'OK'
        ForEach($global:hostname in $global:hostnames)
        {
            Get-Info
            Log-DataGrid
            Log-ForCSV
            $global:hostname = $null
            $global:ip = $null
            $global:imageversion = $null
        }
        Message-Box -Message "Complete and written to Data Grid!" -WindowTitle 'Done!' -ButtonStyle 'OK'
    }

    ## Multiple Hostname CSV
    IF($machineinput.SelectedItem -eq "List from CSV Path")
    {
        $path = $hostnameinput.Text
        $csv = Import-CSV -Path $path -Delimiter "," -Header @('hostname') | Select "hostname"
        $first = ($csv | Select-Object -ExpandProperty hostname)[0]
        IF($first -like "L*"){$global:hostnames = $csv | Select-Object -ExpandProperty "Hostname"}
        ELSEIF($first -like "D*"){$global:hostnames = $csv | Select-Object -ExpandProperty "Hostname"}
        ELSEIF($first -like "V*"){$global:hostnames = $csv | Select-Object -ExpandProperty "Hostname"}
        ELSE{$global:hostnames = $csv | Select-Object -ExpandProperty "Hostname" -Skip 1}
        ForEach($global:hostname in $global:hostnames)
        {
            Get-Info
            Log-DataGrid
            Log-ForCSV
            $global:hostname = $null
            $global:ip = $null
            $global:imageversion = $null
        }
        Message-Box -Message "Complete and written to Data Grid!" -WindowTitle 'Done!' -ButtonStyle 'OK'
    }
})
$form.controls.add($Checkmachines)

## Export Log Button
$export = New-Object System.Windows.Forms.Button
$export.Location = New-Object System.Drawing.Point(10,390)
$export.Size = New-Object System.Drawing.Size(50,50)
$export.Name = "export"
$export.add_MouseHover($ShowHelp)
$export.Text = 'Export Log'
$export.BackColor = 'LightBlue'
$export.Add_Click({
    $global:exportpath = Read-InputBoxDialog -Message "Enter full path and filename to export to.`r`nExample: C:\Temp\Results.csv`r`nNote: If csv already exists, new rows will be added." -WindowTitle 'Enter Path'
    Log-ToCSV
    Message-Box -Message "Export Complete!" -WindowTitle 'Results' -ButtonStyle 'OK'
})
$form.controls.add($export)

## Launch SCCM Collection Manager Button
$sccm = New-Object System.Windows.Forms.Button
$sccm.Location = New-Object System.Drawing.Point(70,390)
$sccm.Size = New-Object System.Drawing.Size(70,50)
$sccm.Text = 'Manage Collections'
$sccm.name = "SCCM"
$sccm.Add_MouseHover($ShowHelp)
$sccm.BackColor = 'LightBlue'
$sccm.Add_Click({
    & "\\CLIENT NAS STORAGE LOCATION\automation\newassetchecklist\files\SCCM Collection Manager.exe"
})
$form.controls.add($sccm)

## Cancel Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(575,390)
$CancelButton.Size = New-Object System.Drawing.Size(50,50)
$CancelButton.Text = 'Done'
$CancelButton.BackColor = 'Green'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$CancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$form.showdialog()

## Restore Error Action Preference
$ErrorActionPreference = "continue"
