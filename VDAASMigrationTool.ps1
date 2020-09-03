## VDAAS Migration Tool Version 2.0
## Author: John Averill, IBM

## Overview: A script to migrate Windows 7 Virtual Machines to Windows 10 Virtual Machines on the same Domain.

## Declare Directories and Locations.
$migdirectory = "\\CLIENT NAS LOCATION\CMTools\usmt"
$source = Read-Host -Prompt 'Source Host Name'
$destination = Read-Host -Prompt 'Destination Host Name'

## Capture Technician Information for Log Files.
Write-Host 'Capturing Technician Information for Log Files'
$current = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]$current
$principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
$curruser = $current.name
$user = $curruser.TrimStart("CLIENT DOMAIN NAME\")
$computer = $env:COMPUTERNAME 

## Test Connection to both machines.
$sourceonline = Test-Connection -ComputerName $source -Count 1 -Quiet
$destinationonline = Test-Connection -ComputerName $destination -Count 1 -Quiet
IF(!$sourceonline){Write-Host 'Source Machine Offline'; Pause; Return}
IF(!$destinationonline){Write-Host 'Destination Machine Offline'; Pause; Return}

## Copy USMT directories to both machines.
Write-Host "Copy Migration USMT folders to source and destination machines"
RoboCopy "$migdirectory\Folder to be placed on source machine" "\\$source\C$\Folder to be placed on source machine" /mir /mt:64
RoboCopy "$migdirectory\Folder to be placed on destination machine" "\\$destination\C$\Folder to be placed on destination machine" /mir /mt:64

## Get Credentials for Admin user.
Write-Host "Get Credentials"
$ManualDirectory = "\\CLIENT NAS LOCATION\USMT_Manual_Storage"
$scanstatecontent = Get-Content "$ManualDirectory\ScanstateT.bat"
$credentials = Get-Credential
$username = $credentials.UserName
$securepassword = $credentials.Password
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securepassword)
$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

## Edit ScanState.
Write-Host "Edit ScanState."
IF($scanstatecontent){Remove-Item -Path "\\$source\C$\USMT\SourceMachine\Scanstate.bat" -EA 'SilentlyContinue' -Force}
$newscanstate = "@echo off`r`nnet use z: \\CLIENT NAS LOCATION\USMT_Manual_Storage /user:$username $password`r`nC:\USMT\SourceMachine\USMT-1702\scanstate.exe Z:\USMT\%ComputerName% /o /localonly /encrypt:AES_256 /key:refkey /efs:copyraw /c /l:c:\Flags\ScanState_Manual.log /uel:90 /ue:%ComputerName%\* /i:C:\USMT\SourceMachine\USMT-1702\AICIApp.xml /i:C:\USMT\SourceMachine\USMT-1702\AICIUser.xml`r`nnet use z: /delete"
IF(New-Item -Path "\\$source\C$\USMT\SourceMachine\" -ItemType File -Name 'Scanstate.bat' -Force){
    $newscanstate | Set-Content -Path "\\$source\C$\USMT\SourceMachine\Scanstate.bat" -Force
    Write-Host 'Scanstate Modified.'}

## Execute ScanState.bat on Source Machine to create .MIG file.
$result1 = Invoke-Command -ComputerName $source -EA 'SilentlyContinue' -ScriptBlock{
   & "C:\USMT\SourceMachine\Scanstate.bat" -wait
}
$result1

## Failure State Check and Logging for ScanState.bat.
IF(!$result1){
     Write-Host 'Failed to run scanstate.bat!'
     $result4 = @()
     $logging3 = @{
          Date = Get-Date
          Source = "$source"
          Destination = "$destination"
}
$result4 += New-Object PSObject -Property $logging3
$result4 | Export-CSV -Path "\\CLIENT LOGS NAS DRIVE\OSDLogs\USMT_Data_Transfer_Logs\Failures\FailureLogs-Source.csv" -NoTypeInformation -Append -Force
Return}
ELSE{Write-Host 'Success! scanstate.bat executed!'}

## Edit LoadState.
$ManualDirectory = "\\CLIENT NAS DRIVE\USMT_Manual_Storage"
$loadstatecontent = Get-Content "$ManualDirectory\LoadState.bat"
IF($loadstatecontent){Remove-Item -Path "\\$destination\C$\USMT\DestionationMachine\Loadstate.bat" -EA 'SilentlyContinue' -Force}
$newloadstate = "@echo off`r`nnet use z: \\CLIENT NAS LOCATION\USMT_Manual_Storage /user:$username $password `r`nC:\USMT\DestionationMachine\USMT-1702\loadstate.exe Z:\USMT\$source /c /ue:%ComputerName%\* /decrypt:AES_256 /key:refkey /v:5 /i:C:\USMT\DestionationMachine\USMT-1702\AICIApp.xml /i:C:\USMT\DestionationMachine\USMT-1702\AICIUser.xml /l:c:\Flags\LoadState_Manual.log`r`nnet use Z: /delete`r`npause"
IF(New-Item -Path "\\$destination\C$\USMT\DestionationMachine\" -ItemType File -Name 'LoadState.bat' -Force){
    $newloadstate | Set-Content -Path "\\$destination\C$\USMT\DestionationMachine\LoadState.bat" -Force
    Write-Host 'LoadState Modified.'}

## Execute LoadState.bat on destination machine to load .MIG file.
Write-Host "Cleaning up ScanState.bat!"
$deletescanstateedit = Remove-Item -Path "\\$source\C$\USMT\SourceMachine\Scanstate.bat" -Force
IF($deletescanstateedit){Write-Host 'Scanstate.bat deleted.'}
Write-Host "Executing LoadState.bat"
$result2 = Invoke-Command -ComputerName $destination -EA 'SilentlyContinue' -ScriptBlock{
    cmd /c "C:\USMT\DestionationMachine\LoadState.bat"}
$result2
$loadstatelogging = "\\$destination\C$\Flags\LoadState_Manual.log"

## Failure State Check for LoadState.bat.
IF(!$result2){
     Write-Host 'Failed to run LoadState.bat!'
     $result5 = @()
     $logging4 = @{
     Date = Get-Date
     Source = "$source"
     Destination = "$destination"
          }
     $result5 += New-Object PSObject -Property $logging4
     $result5 | Export-CSV -Path "\\CLIENT NAS LOCATION\OSDLogs\USMT_Data_Transfer_Logs\Failures\FailureLogs-Destination.csv" -NoTypeInformation -Append -Force
     Return}
ELSEIF(Select-String -Path $loadstatelogging -Pattern 'Failed.'){
     Write-Host "Log File shows failures on copying specific files with LoadState.bat! Migration ran, but may only be partial. See \\CLIENT NAS LOCATION\OSDLogs\USMT_Data_Transfer_Logs\Failures\LoadStateLogs\$user-$source-$destination.log for more information."
     $result6 = @()
     $logging5 = @{
     Date = Get-Date
     Source = "$source"
     Destination = "$destination"
          }
     $result6 += New-Object PSObject -Property $logging5
     $result6 | Export-CSV -Path "\\CLIENT NAS LOCATION\OSDLogs\USMT_Data_Transfer_Logs\Failures\FailureLogs-LoadStateErrorsDetected.csv" -NoTypeInformation -Append -Force
     Copy-Item -Path "$loadstatelogging" -Destination "\\CLIENT NAS LOCATION\OSDLogs\USMT_Data_Transfer_Logs\Failures\LoadStateLogs\$user-$source-$destination.log" -Force -EA 'SilentlyContinue' -Recurse
     }
ELSE{
     Write-Host 'Success! LoadState.bat executed with no errors!'
     $result3 = @()
     $logging2 = @{
          Date = Get-Date
          Source = "$source"
          Destination = "$destination"
     }
     $result3 += New-Object PSObject -Property $logging2
     $result3 | Export-CSV -Path "\\CLIENT NAS LOCATION\OSDLogs\USMT_Data_Transfer_Logs\Success\SuccessLogs.csv" -NoTypeInformation -Append -Force
     }

## Copy Sticky Notes, because apparently that's a problem for our users, and the Desktop Engineering team does not wish to add it to the ScanState.
$users = Invoke-Command -ComputerName $source {Get-ChildItem C:\Users} | Select-Object -ExpandProperty Name
ForEach ($user in $users){
    $stickynotes = Test-Path "\\$source\C$\Users\$user\AppData\Roaming\Microsoft\Sticky Notes\StickyNotes.snt"
    IF($stickynotes -eq $True){
        $Legacy = Test-Path "\\$destination\C$\Users\$user\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\Legacy\"
        IF($Legacy -eq $False)
            {New-Item -Path "\\$destination\C$\Users\$user\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState" -ItemType Directory -Name Legacy -Force}
        IF(Copy-Item -Path "\\$source\C$\Users\$user\AppData\Roaming\Microsoft\Sticky Notes\StickyNotes.snt" -Destination "\\$destination\C$\Users\$user\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\Legacy\" -Force -Recurse -Verbose){Write-Host 'StickyNotes Copied'}
        IF(Rename-Item -Path "\\$destination\C$\Users\$user\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\Legacy\StickyNotes.snt" -NewName "ThresholdNotes.snt" -EA 'SilentlyContinue' -Force){Write-Host 'Threshold Notes created'}
            Write-Host "$user Sticky Notes Data Copied!"}
        IF(Remove-Item -Path "\\$destination\C$\Users\$user\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\Settings" -EA 'SilentlyContinue' -Force -Recurse){Write-Host 'Default StickyNotes Settings Deleted.'}
        ELSE{Write-Host "No Sticky Notes Data found for user $user"}}
    Write-Host 'Sticky Notes Transfer Complete'

## Logging. 
Write-Host 'Writing Log Files'
$scanstate = Copy-Item -Path "\\$source\C$\Flags\ScanState_Manual.log" -Destination "\\CLIENT NAS LOCATION\OSDLogs\USMT_Data_Transfer_Logs\Success\ScanStateLogs\$user-ScanState_Manual_$source.log" -Force -EA 'SilentlyContinue' -Recurse
$loadstate = Copy-Item -Path "\\$destination\C$\Flags\LoadState_Manual.log" -Destination "\\CLIENT NAS LOCATION\OSDLogs\USMT_Data_Transfer_Logs\Success\LoadStateLogs\$user-$source-$destination.log" -Force -EA 'SilentlyContinue' -Recurse
IF($scanstate) {Write-Host "Logs state successful Scan State Execution on $source, from $computer, by $user"}
IF($loadstate) {Write-Host "Logs state successful Load State Execution on $destination, from $computer, by $user"}

## Cleanup of Migration Data and all credential information.
$deleteloadstateedit = Remove-Item -Path "\\$destination\C$\USMT\DestionationMachine\LoadState.bat" -Force
IF($deleteloadstateedit){Write-Host 'LoadState.bat deleted.'}
$deletescanstate = Remove-Item -Path "\\$source\C$\USMT\SourceMachine\Scanstate.bat"
IF($deletescanstate){Write-Host 'ScanState.bat deleted.'}
IF($scanstate){
    IF($loadstate){
         Invoke-Command -ComputerName $destination -EA 'SilentlyContinue' -ScriptBlock {
              Remove-Item "C:\USMT" -Force -recurse}
         Write-Host 'Directory Deleted!'}
    ELSE{Write-Host "Errors found in executing LoadState.bat on $destination. USMT directory left for manual attempt."}}
ELSE{Write-Host "Errors found in executing ScanState.bat on $source. USMT directory left for manual attempt."}

## End.
