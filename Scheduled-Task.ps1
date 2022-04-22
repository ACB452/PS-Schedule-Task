
$computers = "hbzmb42w-vm", "CLDW10058", "CAIV017FH60R2L", "CAIV01CFH60R2L", "CAHM82D2J60R2L", "COCS11B8ST1R2L", "CAIV0121FL2R2L", "CLDW10086"

### Create a Session ###
function Open-Session {
    $computers = "CLDW10086"
    New-PSSession -ComputerName $computers
    $session = Get-PSSession
}

foreach ($computer in $computers) {
    
}




# COPY ADFDI locally
Invoke-Command -Session $session -ScriptBlock {Copy-item -path "\\az25desktop\dsl$\DICE\Packages\OracleADFDIplugin\OracleADFDIplugin\lib\ADFDI\" -destination "\\$computers\c$\temp" -recurse}

## Install ADFDE .exe ##
#Invoke-Command -Session $session -ScriptBlock {Start-Process '\\az25desktop\dsl$\DICE\Packages\OracleADFDIplugin\OracleADFDIplugin\lib\ADFDI\adfdi-excel-addin-installer.exe' -ArgumentList '/quiet /log C:\temp\inst.log' -Wait }
Invoke-Command -Session $session -ScriptBlock {Start-Process 'C:\temp\ADFDI\adfdi-excel-addin-installer.exe' -ArgumentList '/quiet /log C:\temp\inst.log' -Wait }
# Check if installed
Invoke-Command -Session $session -ScriptBlock {get-wmiobject Win32_Product | Where-Object {$_.Name -match "ADF"} | Format-Table Name }

# -----------------------------
## Add to change to current logged in user
Invoke-Command -Session $session -ScriptBlock {$User = (Get-WmiObject -Class Win32_Process -Filter 'Name="explorer.exe"').GetOwner().User}
# Run VSTO
Invoke-Command -Session $session -ScriptBlock {Start-Process 'C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe' -ArgumentList '/I "C:\Users\$User\AppData\Local\Oracle\Oracle ADF 11g Desktop Integration Add-In for Excel\adfdi-excel-addin.vsto" /S' -Wait }
# -----------------------------
#foreach ($computer in $computers) {

#}
# Run VSTO Installer
#Invoke-Command -ComputerName "hbzmb42w-vm" -ScriptBlock {Start-Process '\\hbzmb42w-vm\c$\temp\ADFDI\adfdi-excel-addin-installer.exe' -ArgumentList '/quiet /log C:\temp\inst.log' -Wait }

# New-ScheduledTaskTrigger
# If not exist - add
# Trigger - expiration?
$task.Settings.Hidden = $true
#Invoke-Command -Session $session -ScriptBlock {$Time = New-ScheduledTaskTrigger -At 1:00PM -Once}
Invoke-Command -Session $session -ScriptBlock {$Time = New-ScheduledTaskTrigger -AtLogOn}
# Action
Invoke-Command -Session $session -ScriptBlock {$Action=New-ScheduledTaskAction -Execute 'C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe' -WorkingDirectory 'C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe' -Argument '/I C:\Users\<UserName>\ "\\$computers\c$\Users\Abraham.Baquilod.su\AppData\Local\Oracle\Oracle ADF 11g Desktop Integration Add-In for Excel\adfdi-excel-addin.vsto" /Silent'}
Invoke-Command -Session $session -ScriptBlock {$Action=New-ScheduledTaskAction -Execute 'C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe' -WorkingDirectory 'C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe' -Argument '/I "C:\Users\Abraham.Baquilod.su\AppData\Local\Oracle\Oracle ADF 11g Desktop Integration Add-In for Excel\adfdi-excel-addin.vsto"'}
#Invoke-Command -Session $session -ScriptBlock {$Action=New-ScheduledTaskAction -Execute "C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe" -WorkingDirectory "C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe" -Argument "/I C:\Users\<UserName>\ '\\$computers\c$\Users\Abraham.Baquilod.su\AppData\Local\Oracle\Oracle ADF 11g Desktop Integration Add-In for Excel\adfdi-excel-addin.vsto' /Silent"}
# Schedule task
Invoke-Command -Session $session -ScriptBlock {Register-ScheduledTask -TaskName "Run VSTO Installer" -Trigger $Time -Action $Action -RunLevel Highest}
Invoke-Command -Session $session -ScriptBlock {Start-ScheduledTask -TaskName "Run VSTO Installer"}
## CHECK scheduled task ##
Invoke-Command -Session $session -ScriptBlock {Get-ScheduledTask -TaskName "Run VSTO Installer" -Verbose}
## Add wait time to fully install
# remove scheduled task after completion
Invoke-Command -Session $session -ScriptBlock {Unregister-ScheduledTask -TaskName "Run VSTO Installer"}



#restart app
Stop-Process -Name "EXCEL.EXE" -force
Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE" -Verb RunAs
Restart-Computer -ComputerName $computers



# end session - fixes issue is state:broken
Disconnect-PSSession -Session $session
Get-PSSession | Disconnect-PSSession
Get-PSSession | Remove-PSSession
Get-PSSession
