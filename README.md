# MECMOfflineUpdateGUI

To make updates easier I have prepared a PowerShell script that has a GUI to perform the updates. As with the Service Connection Tool, this needs to be run as Administrator.
The GUI has three tabs. The two tabs that have a red background are run on the Site System Server with  Service Connection Point role installed. The green tab is run on the internet connected computer.
When you first run the script, it looks for the Service Connection Tool, if it cannot find it, it will ask you to locate the folder. First it looks to see if the $env:SMS_LOG_PATH\..\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool\ServiceConnectionTool.exe exists, if it does not it looks to see if the  “$PSScriptRoot\ServiceConnectionTool\ ServiceConnectionTool.exe” file exists ($PSScriptRoot is the location the script is run from). If it finds the ServiceConnectionTool.exe file it will use that folder otherwise it will prompt for it.
