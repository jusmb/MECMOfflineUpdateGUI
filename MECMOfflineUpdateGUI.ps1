# -------------------------------------------------------------------------------------------------
#  <copyright file="MECMOfflineUpdateGUI.ps1" Company="Justin Mnatsakanyan-Barbalace,Sr CSA-E, Microsoft, 2022">
#      All rights reserved.
#  </copyright>
#
#  Description: This script prepares the host machine for various Azure Migrate Scenarios.

#  Version: 1.0.3

#  Requirements: Run as Administrator
#  
<#
LEGAL DISCLAIMER
This Sample Code is provided for the purpose of illustration only and is not
intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
nonexclusive, royalty-free right to use and modify the Sample Code and to
reproduce and distribute the object code form of the Sample Code, provided
that You agree: (i) to not use Our name, logo, or trademarks to market Your
software product in which the Sample Code is embedded; (ii) to include a valid
copyright notice on Your software product in which the Sample Code is embedded;
and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
against any claims or lawsuits, including attorneys’ fees, that arise or result
from the use or distribution of the Sample Code.
 
This script is provided "AS IS" with no warranties, and confers no rights. Use
of included script samples are subject to the terms specified
at http://www.microsoft.com/info/cpyright.htm.
#>
$ScriptFullPath=$myInvocation.InvocationName 
$ScriptName=$myInvocation.MyCommand

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)

#-------------------------------------------------------------#
#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#

Add-Type -AssemblyName PresentationCore, PresentationFramework, System.Windows.Forms

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="700" Height="450">
<Grid>
 <TabControl Margin="15,15,15,15">
     <TabItem  Name="ExTab" Header="Export Usage Data Cab">
         <Grid Background="#ff0000">
         
             <Label Name="ExDescription" HorizontalAlignment="Left" VerticalAlignment="Top" Content="This utility should be run on the server with the Service Connection Point role." Margin="10,4,0,0" Foreground="#ffffff" FontWeight="Bold" Width="637" Height="25" FontStyle="Italic"/>
         
             <TextBox Name="ExSCTLocation" HorizontalAlignment="Left" VerticalAlignment="Top" Height="22" Width="314" Margin="231,30,0,0" FontWeight="Medium"/>
             <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="ServiceConnectionTool.exe Location" Margin="10,30,0,0" Foreground="#ffffff" FontWeight="Medium" Width="214" Height="27"/>
             <Button Name="ExSCTBrowse" Content="Browse" HorizontalAlignment="Left" VerticalAlignment="Top" Width="76" Margin="555,30,0,0" Height="23"/>

             <Label Name="ExDescription1" HorizontalAlignment="Left" VerticalAlignment="Top" Content="This utility will out put several files that need to be moved to a internet connected device. Please select a folder " Margin="10,75,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>
             <Label Name="ExDescription2" HorizontalAlignment="Left" VerticalAlignment="Top" Content="to out put to And make sure the ServiceConnectionTool.exe is the current version in the CD.Latest folder. By default " Margin="10,95,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>
             <Label Name="ExDescription3" HorizontalAlignment="Left" VerticalAlignment="Top" Content="the location this script is run from is where the package is created If tis is not where you want it created select " Margin="10,115,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>
             <Label Name="ExDescription4" HorizontalAlignment="Left" VerticalAlignment="Top" Content="a new location to create the package. " Margin="10,135,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>

             <TextBox Name="ExPackageLocation" HorizontalAlignment="Left" VerticalAlignment="Top" Height="22" Width="314" Margin="231,175,0,0" FontWeight="Medium"/>
             <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="Location to save the transfer package" Margin="10,175,0,0" Foreground="#ffffff" FontWeight="Medium" Width="214" Height="27"/>
             <Button Name="ExPackageBrowse" Content="Browse" HorizontalAlignment="Left" VerticalAlignment="Top" Width="76" Margin="555,175,0,0" Height="23"/>

             <Button Content="Export Usage" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="545,295,0,0" Name="ExportUsageData"/>
         </Grid>
    </TabItem>
    <TabItem  Name="DLTab" Header="Download Updates">
         <Grid Background="#008000">
             <Label Name="DLDescription" HorizontalAlignment="Left" VerticalAlignment="Top" Content="This utility should be run on the server with the Service Connection Point role." Margin="10,4,0,0" Foreground="#ffffff" FontWeight="Bold" Width="637" Height="25" FontStyle="Italic"/>
         
             <TextBox Name="DLSCTLocation" HorizontalAlignment="Left" VerticalAlignment="Top" Height="22" Width="314" Margin="231,30,0,0" FontWeight="Medium"/>
             <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="ServiceConnectionTool.exe Location" Margin="10,30,0,0" Foreground="#ffffff" FontWeight="Medium" Width="214" Height="27"/>
             <Button Name="DLSCTBrowse" Content="Browse" HorizontalAlignment="Left" VerticalAlignment="Top" Width="76" Margin="555,30,0,0" Height="23"/>

             <Label Name="DLDescription1" HorizontalAlignment="Left" VerticalAlignment="Top" Content="The computer this is downloaded from needs to be able to connect to the internet." Margin="10,75,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>
             <Label Name="DLDescription2" HorizontalAlignment="Left" VerticalAlignment="Top" Content="" Margin="10,95,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>
             <Label Name="DLDescription3" HorizontalAlignment="Left" VerticalAlignment="Top" Content="" Margin="10,115,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>
             <Label Name="DLDescription4" HorizontalAlignment="Left" VerticalAlignment="Top" Content="If the location of the Usagedata.cab folder is not automatically selected, please select it. " Margin="10,135,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>

             <TextBox Name="DLPackageLocation" HorizontalAlignment="Left" VerticalAlignment="Top" Height="22" Width="314" Margin="231,175,0,0" FontWeight="Medium"/>
             <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="Root location where files will be saved" Margin="10,175,0,0" Foreground="#ffffff" FontWeight="Medium" Width="230" Height="27"/>
             <Button Name="DLPackageBrowse" Content="Browse" HorizontalAlignment="Left" VerticalAlignment="Top" Width="76" Margin="555,175,0,0" Height="23"/>
             <Grid HorizontalAlignment="Left" VerticalAlignment="Top" Width="600" Height="164" Margin="10,200,0,0">
                <RadioButton Name="DownloadAll" GroupName="DownloadType" Foreground="#ffffff" HorizontalAlignment="Left" VerticalAlignment="Top" Content="Download everything, including updates and hotfixes." Margin="16,10,0,0" Width="370" Height="38"/>
                <RadioButton Name="DownloadHotfix" GroupName="DownloadType" Foreground="#ffffff" HorizontalAlignment="Left" VerticalAlignment="Top" Content="Only download all hotfixes" Margin="16,40,0,0"  Width="370" />
                <RadioButton Name="DownloadSiteVersion" GroupName="DownloadType" Foreground="#ffffff" IsChecked="True"  HorizontalAlignment="Left" VerticalAlignment="Top" Content="Only download updates and hotfixes that have a newer version than the version of your site" Margin="16,70,0,0"  Width="600" />
             </Grid>
             <Button Name="DownloadUpdates" Content="Download Updates" HorizontalAlignment="Left" VerticalAlignment="Top" Width="110" Margin="530,310,0,0" />
         </Grid>
     </TabItem>
     <TabItem Name="ULTab" Header="Send Updates To CM">
         <Grid Background="#b22222">
             <Label Name="ULDescription" HorizontalAlignment="Left" VerticalAlignment="Top" Content="This utility should be run on the server with the Service Connection Point role." Margin="10,4,0,0" Foreground="#ffffff" FontWeight="Bold" Width="637" Height="25" FontStyle="Italic"/>
         
             <TextBox Name="ULSCTLocation" HorizontalAlignment="Left" VerticalAlignment="Top" Height="22" Width="314" Margin="231,30,0,0" FontWeight="Medium"/>
             <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="ServiceConnectionTool.exe Location" Margin="10,30,0,0" Foreground="#ffffff" FontWeight="Medium" Width="214" Height="27"/>
             <Button Name="ULSCTBrowse" Content="Browse" HorizontalAlignment="Left" VerticalAlignment="Top" Width="76" Margin="555,30,0,0" Height="23"/>

             <Label Name="ULDescription1" HorizontalAlignment="Left" VerticalAlignment="Top" Content="The UpdatePack needs to be run on the computer with the Service Connection Point MECM role installed on it." Margin="10,75,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>
             <Label Name="ULDescription2" HorizontalAlignment="Left" VerticalAlignment="Top" Content="" Margin="10,95,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>
             <Label Name="ULDescription3" HorizontalAlignment="Left" VerticalAlignment="Top" Content="" Margin="10,115,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>
             <Label Name="ULDescription4" HorizontalAlignment="Left" VerticalAlignment="Top" Content="If the location of the UpdatePack folder is not automatically selected, please select it. " Margin="10,135,0,0" Foreground="#ffffff" FontWeight="Medium" Width="637" Height="24"/>

             <TextBox Name="ULPackageLocation" HorizontalAlignment="Left" VerticalAlignment="Top" Height="22" Width="314" Margin="231,175,0,0" FontWeight="Medium"/>
             <Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="Location of the UpdatePacks folder" Margin="10,175,0,0" Foreground="#ffffff" FontWeight="Medium" Width="214" Height="27"/>
             <Button Name="ULPackageBrowse" Content="Browse" HorizontalAlignment="Left" VerticalAlignment="Top" Width="76" Margin="555,175,0,0" Height="23"/>

             <Button Name="UploadUpdates" Content="Upload Updates" HorizontalAlignment="Left" VerticalAlignment="Top" Width="110" Margin="530,310,0,0" />
         </Grid>
     </TabItem>
 </TabControl>
</Grid></Window>
"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#


#region Logic
#Functions
#get-childitem $PSScriptRoot -Recurse | where{$_.Name -eq "ServiceConnectionTool.exe"}


Function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton)
{
$AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
$Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
$OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
$OpenFileDialog.AddExtension = $false
$OpenFileDialog.CheckFileExists = $false
$OpenFileDialog.DereferenceLinks = $true
$OpenFileDialog.Filter = "Folders|`n"
$OpenFileDialog.Multiselect = $false
if($Message){$OpenFileDialog.Title =$Message}else{$OpenFileDialog.Title = "Select folder"}
$OpenFileDialog.InitialDirectory=$InitialDirectory
$OpenFileDialogType = $OpenFileDialog.GetType()
$FileDialogInterfaceType = $Assembly.GetType('System.Windows.Forms.FileDialogNative+IFileDialog')
$IFileDialog = $OpenFileDialogType.GetMethod('CreateVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null)
$OpenFileDialogType.GetMethod('OnBeforeVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$IFileDialog)
[uint32]$PickFoldersOption = $Assembly.GetType('System.Windows.Forms.FileDialogNative+FOS').GetField('FOS_PICKFOLDERS').GetValue($null)
$FolderOptions = $OpenFileDialogType.GetMethod('get_Options',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null) -bor $PickFoldersOption
$FileDialogInterfaceType.GetMethod('SetOptions',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$FolderOptions)
$VistaDialogEvent = [System.Activator]::CreateInstance($AssemblyFullName,'System.Windows.Forms.FileDialog+VistaDialogEvents',$false,0,$null,$OpenFileDialog,$null,$null).Unwrap()
[uint32]$AdviceCookie = 0
$AdvisoryParameters = @($VistaDialogEvent,$AdviceCookie)
$AdviseResult = $FileDialogInterfaceType.GetMethod('Advise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdvisoryParameters)
$AdviceCookie = $AdvisoryParameters[1]
$Result = $FileDialogInterfaceType.GetMethod('Show',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,[System.IntPtr]::Zero)
$FileDialogInterfaceType.GetMethod('Unadvise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdviceCookie)
if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
    $FileDialogInterfaceType.GetMethod('GetResult',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$null)
    
}
[string]$output=$OpenFileDialog.FileName

write-host $output.trim()
if ($output.trim() -ne ""){ return $output.trim()}


}



#endregion 


#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }
################
#get SCT
################

Try{
    $SCTPath=Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Microsoft\CCM_OfflineUpdateTool -Name SCTPath -ErrorAction SilentlyContinue -Verbose 
}catch{
    $SCTPath=$null
        
}
if($SCTPath -ne $null){
    $ExSCTLocation.Text=$SCTPath.trim()

}elseif(test-path $env:SMS_LOG_PATH\..\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool\ServiceConnectionTool.exe){
    $ExSCTLocation.Text="$($env:SMS_LOG_PATH)\..\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool"

}
if(Test-Path "$PSScriptRoot\ServiceConnectionTool\ServiceConnectionTool.exe"){
    $ULSCTLocation.Text="$PSScriptRoot\ServiceConnectionTool"
}
$ExTab.IsFocused
##################################
#Export Usage Data (Ex)
##################################
<#
if(test-path $env:SMS_LOG_PATH\..\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool\ServiceConnectionTool.exe){
    $ExSCTLocation.Text="$($env:SMS_LOG_PATH)\..\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool"
    $ExSCTLocation.Text=$ExSCTLocation.Text.trim()
    $DLSCTLocation.Text=$ExSCTLocation.Text.trim()
    $ULSCTLocation.Text=$ExSCTLocation.Text.trim()
}elseif(test-path $PSScriptRoot\ServiceConnectionTool\ServiceConnectionTool.exe){
    $ExSCTLocation.Text="$($PSScriptRoot)\ServiceConnectionTool"
    $ExSCTLocation.Text=$ExSCTLocation.Text.trim()
    $DLSCTLocation.Text=$ExSCTLocation.Text.trim()
    $ULSCTLocation.Text=$ExSCTLocation.Text.trim()
}else{
    
    do{
        [System.Windows.MessageBox]::Show('I need the location of the Service Connection tool to continue. It is usually located at "\\<SiteServer>\SMS_<SiteCode>\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool". Please browse to that location.')
        $ExSCTLocation.Text=Read-FolderBrowserDialog -Message "I cannot locate the folder ServiceConnectionTool folder. Please browse to it." 
        $ExSCTLocation.Text=$ExSCTLocation.Text.trim()
        $DLSCTLocation.Text=$ExSCTLocation.Text.trim()
        $ULSCTLocation.Text=$ExSCTLocation.Text.trim()
        test-path "$($ExSCTLocation.Text.trim())\ServiceConnectionTool.exe"
       
    }While(!(test-path "$($ExSCTLocation.Text.trim())\ServiceConnectionTool.exe"))
}
#>


if(Test-Path "$PSScriptRoot\ServiceConnectionTool\ServiceConnectionTool.exe"){
    $DLSCTLocation.Text="$PSScriptRoot\ServiceConnectionTool"
}
$ExTab.Add_GotFocus({
    if(test-path $env:SMS_LOG_PATH\..\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool\ServiceConnectionTool.exe){
        $ExSCTLocation.Text="$($env:SMS_LOG_PATH)\..\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool"
        $ExSCTLocation.Text=$ExSCTLocation.Text.trim()
        $ULSCTLocation.Text=$ExSCTLocation.Text.trim()
    
    }elseif($SCTPath -ne $null){
        $ExSCTLocation.Text=$SCTPath.trim()
        $ULSCTLocation.Text=$ExSCTLocation.Text.trim()

    }elseif(test-path $PSScriptRoot\ServiceConnectionTool\ServiceConnectionTool.exe){
        $ExSCTLocation.Text="$($PSScriptRoot)\ServiceConnectionTool"
        $ExSCTLocation.Text=$ExSCTLocation.Text.trim()
        $ULSCTLocation.Text=$ExSCTLocation.Text.trim()
    }else{
    
        do{
            if(!(test-path "$($ExSCTLocation.Text.trim())\ServiceConnectionTool.exe")){
                [System.Windows.MessageBox]::Show('I need the location of the Service Connection tool to continue. It is usually located at "\\<SiteServer>\SMS_<SiteCode>\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool". Please browse to that location. The Service Connection tool needs to be the version that came with the site being updated.')
                $ExSCTLocation.Text=Read-FolderBrowserDialog -Message "I cannot locate the folder ServiceConnectionTool folder. Please browse to it." 
                $ExSCTLocation.Text=$ExSCTLocation.Text.trim()
            }
            
       
        }While(!(test-path "$($ExSCTLocation.Text.trim())\ServiceConnectionTool.exe"))
    }

})

$ExPackageLocation.Text=$PSScriptRoot

$ExSCTBrowse.Add_Click({
    [string]$Value=$null
    $Value=Read-FolderBrowserDialog -Message "Please select a directory" -initialDirectory $ExSCTLocation.Text.trim()
    if($value.trim() -ne "" -and $value -ne $null){$ExSCTLocation.Text=$value.trim()}
})


$ExPackageBrowse.Add_Click({
    [string]$Value=$null
    $Value=Read-FolderBrowserDialog -Message "Please select a directory" -initialDirectory $ExPackageLocation.Text.trim()
    if($value.trim() -ne "" -and $value -ne $null){
        $ExPackageLocation.Text=$value.trim()
        Write-Host $ExPackageLocation.Text.trim()
    }
})

$ExportUsageData.Add_Click({

        
    if($SCPath -eq $Null){
        New-Item -Path HKLM:\SOFTWARE\Microsoft\CCM_OfflineUpdateTool -Verbose -Force
        New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\CCM_OfflineUpdateTool -Name SCTPath -PropertyType String -Value $ExSCTLocation.Text.trim() -Verbose -Force
        $SCTPath=$ExSCTLocation.Text.trim()
    }elseif($SCTPath.trim() -ne $ExSCTLocation.Text.trim()){
            Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\CCM_OfflineUpdateTool -Name SCTPath -PropertyType String -Value $ExSCTLocation.Text.trim() -Verbose -Force
    }
        
    $DateSerial=Get-Date -Format "yyyy.MM.dd.HH.mm.ss"
    $reg=$null
    $reg=Get-ChildItem -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SMS\Components\SMS_SERVICE_CONNECTOR -ErrorAction SilentlyContinue -Verbose
    if($reg.count -lt 1){[System.Windows.MessageBox]::Show('Export Usage Data needs to be run on the server with the MECM Service Connection Point role installed')}else{
        [string]$Value=$null
        if("$($ExPackageLocation.Text.trim())\ServiceConnectionTool" -ne "$($ExSCTLocation.Text.trim())"){
            if(Test-Path "$($ExPackageLocation.Text.trim())\ServiceConnectionTool"){Rename-Item -Path "$($ExPackageLocation.Text.trim())\ServiceConnectionTool" -NewName "$($ExPackageLocation.Text.trim())\ServiceConnectionTool.$DateSerial"}
            Copy-Item -Path $ExSCTLocation.Text.trim() -Destination "$($ExPackageLocation.Text.trim())\ServiceConnectionTool" -Recurse -Verbose
            Copy-Item -Path $ScriptFullPath -Destination "$($ExPackageLocation.Text.trim())\$ScriptName" -Force -ErrorAction SilentlyContinue -Verbose 
        }
        if(Test-Path "$($ExPackageLocation.Text.trim())\UsageData"){Rename-Item -Path "$($ExPackageLocation.Text.trim())\UsageData" -NewName "$($ExPackageLocation.Text.trim())\UsageData.$DateSerial"}
        New-Item "$($ExPackageLocation.Text.trim())\UsageData" -ItemType Directory -Force -Verbose
        Start-Process "$($ExPackageLocation.Text.trim())\ServiceConnectionTool\ServiceConnectionTool.exe" -ArgumentList "-prepare -usagedatadest ""$($ExPackageLocation.Text.trim())\UsageData\$($env:ComputerName)-UsageData.cab""" -Wait -Verbose
        Start-Process "$($ExPackageLocation.Text.trim())\ServiceConnectionTool\ServiceConnectionTool.exe" -ArgumentList "-export -dest ""$($ExPackageLocation.Text.trim())\UsageData\$($env:ComputerName)-UsageDate.csv""" -Wait -Verbose
    }
})

##################################
#Download Updates (DL)
##################################

$DLTab.Add_GotFocus({
    if(test-path $PSScriptRoot\ServiceConnectionTool\ServiceConnectionTool.exe){
        $DLSCTLocation.Text="$($PSScriptRoot)\ServiceConnectionTool"
        $DLSCTLocation.Text=$DLSCTLocation.Text.trim()
    }else{
    
        do{
            if(!(test-path "$($DLSCTLocation.Text.trim())\ServiceConnectionTool.exe")){
                [System.Windows.MessageBox]::Show('I need the location of the Service Connection tool to continue. It is usually located at "\\<SiteServer>\SMS_<SiteCode>\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool". Please browse to that location. The Service Connection tool needs to be the version that came with the site being updated.')
                $DLSCTLocation.Text=Read-FolderBrowserDialog -Message "I cannot locate the folder ServiceConnectionTool folder. Please browse to it." 
                $DLSCTLocation.Text=$DLSCTLocation.Text.trim()
            }
            
       
        }While(!(test-path "$($DLSCTLocation.Text.trim())\ServiceConnectionTool.exe"))
    }
})
$DLPackageLocation.Text="$PSScriptRoot"

$DLPackageBrowse.Add_Click({
    [string]$Value=$null
    $Value=Read-FolderBrowserDialog -Message "Please select the location of the proper version of the Service Connection Tool." -initialDirectory $DLPackageLocation.Text.trim()
    if($value.trim() -ne "" -and $value -ne $null){
        $DLPackageLocation.Text=$value.trim()
        Write-Host $DLPackageLocation.Text.trim()
    }
})

$DownloadUpdates.Add_Click({
    if($DownloadAll.IsChecked){$DownloadType="-downloadall"}
    elseif($DownloadHotfix.IsChecked){$DownloadType="-downloadhotfix"}
    elseif($DownloadSiteVersion.IsChecked){$DownloadType="-downloadsiteversion"}
    $DateSerial=Get-Date -Format "yyyy.MM.dd.HH.mm.ss"
    if(Test-Path "$($DLPackageLocation.Text.trim())\UpdatePacks"){Rename-Item -Path "$($DLPackageLocation.Text.trim())\UpdatePacks" -NewName "$($DLPackageLocation.Text.trim())\UpdatePacks.$DateSerial"}
    New-Item "$($DLPackageLocation.Text.trim())\UpdatePacks" -ItemType Directory -Force -Verbose
    Start-Process "$($DLSCTLocation.Text.trim())\ServiceConnectionTool.exe" -ArgumentList "-connect $DownloadType -usagedatasrc ""$($DLPackageLocation.Text.trim())\UsageData"" -updatepackdest ""$($DLPackageLocation.Text.trim())\UpdatePacks""" -Wait -Verbose
    
})

##################################
#Upload Updates (UL)
##################################
$ULPackageLocation.Text="$PSScriptRoot\UpdatePacks"

$ULTab.Add_GotFocus({
    if(test-path "$($PSScriptRoot)\ServiceConnectionTool\ServiceConnectionTool.exe"){
        
        $ULSCTLocation.Text="$($PSScriptRoot)\ServiceConnectionTool"
    }elseif(test-path $env:SMS_LOG_PATH\..\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool\ServiceConnectionTool.exe){
        $ULSCTLocation.Text="$($env:SMS_LOG_PATH)\..\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool"

    
    }elseif($SCTPath -ne $null -and (Test-Path "$SCTPath\ServiceConnectionTool.exe" -Verbose)){
        $ULSCTLocation.Text=$SCTPath.trim()
    }else{
        do{
            if(!(test-path "$($ULSCTLocation.Text.trim())\ServiceConnectionTool.exe")){
                [System.Windows.MessageBox]::Show('I need the location of the Service Connection tool to continue. It is usually located at "\\<SiteServer>\SMS_<SiteCode>\cd.latest\SMSSETUP\TOOLS\ServiceConnectionTool". Please browse to that location. The Service Connection tool needs to be the version that came with the site being updated.')
                $ULSCTLocation.Text=Read-FolderBrowserDialog -Message "I cannot locate the folder ServiceConnectionTool folder. Please browse to it." 
                $ULSCTLocation.Text=$ULSCTLocation.Text.trim()
            }

       
        }While(!(test-path "$($ULSCTLocation.Text.trim())\ServiceConnectionTool.exe"))
    }
    
    if(Test-Path $($ULPackageLocation.Text.trim())){
        $ULPackageLocation.Text="$PSScriptRoot\UpdatePacks"
    }else{
        do{
            if(!(test-path "$($ULPackageLocation.Text.trim())")){
                [System.Windows.MessageBox]::Show('I need the location of the UpdatePacks to continue. It is usually located at "scriptfolder\UpdatePacks". Please browse to that location. The Service Connection tool needs to be the version that came with the site being updated.')
                $ULPackageLocation.Text=Read-FolderBrowserDialog -Message "I cannot locate the folder UpdatePacks folder. Please browse to it." -InitialDirectory $PSScriptRoot
                $ULPackageLocation.Text=$ULPackageLocation.Text.trim()
            }
       
        }While(!(test-path "$($ULPackageLocation.Text.trim())"))
    }
})

$ULPackageBrowse.Add_Click({
    [string]$Value=$null
    $Value=Read-FolderBrowserDialog -Message "Please select the location of the proper version of the Service Connection Tool." -initialDirectory $ULPackageLocation.Text.trim()
    if($value.trim() -ne "" -and $value -ne $null){
        $ULPackageLocation.Text=$value.trim()
        Write-Host $ULPackageLocation.Text.trim()
    }
})

$UploadUpdates.Add_Click({
    Start-Process "$($ULSCTLocation.Text.trim())\ServiceConnectionTool.exe" -ArgumentList "-import -updatepacksrc ""$($ULPackageLocation.Text.trim())""" -Wait -Verbose
    
})

######################################
$Global:SyncHash = [HashTable]::Synchronized(@{})
$SyncHash.Window = $Window
$Jobs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$initialSessionState = [initialsessionstate]::CreateDefault()

Function Start-RunspaceTask
{
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,Position=0)][ScriptBlock]$ScriptBlock,
          [Parameter(Mandatory=$True,Position=1)][PSObject[]]$ProxyVars)
            
    $Runspace = [RunspaceFactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = 'STA'
    $Runspace.ThreadOptions  = 'ReuseThread'
    $Runspace.Open()
    ForEach($Var in $ProxyVars){$Runspace.SessionStateProxy.SetVariable($Var.Name, $Var.Variable)}
    $Thread = [PowerShell]::Create('NewRunspace')
    $Thread.AddScript($ScriptBlock) | Out-Null
    $Thread.Runspace = $Runspace
    [Void]$Jobs.Add([PSObject]@{ PowerShell = $Thread ; Runspace = $Thread.BeginInvoke() })
}

$JobCleanupScript = {
    Do
    {    
        ForEach($Job in $Jobs)
        {            
            If($Job.Runspace.IsCompleted)
            {
                [Void]$Job.Powershell.EndInvoke($Job.Runspace)
                $Job.PowerShell.Runspace.Close()
                $Job.PowerShell.Runspace.Dispose()
                $Job.Powershell.Dispose()
                
                $Jobs.Remove($Job)
            }
        }

        Start-Sleep -Seconds 1
    }
    While ($SyncHash.CleanupJobs)
}

Get-ChildItem Function: | Where-Object {$_.name -notlike "*:*"} |  select name -ExpandProperty name |
ForEach-Object {       
    $Definition = Get-Content "function:$_" -ErrorAction Stop
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "$_", $Definition
    $InitialSessionState.Commands.Add($SessionStateFunction)
}


$Window.Add_Closed({
    Write-Verbose 'Halt runspace cleanup job processing'
    $SyncHash.CleanupJobs = $False
})

$SyncHash.CleanupJobs = $True
function Async($scriptBlock){ Start-RunspaceTask $scriptBlock @([PSObject]@{ Name='DataContext' ; Variable=$DataContext},[PSObject]@{Name="State"; Variable=$State},[PSObject]@{Name = "SyncHash";Variable = $SyncHash})}

Start-RunspaceTask $JobCleanupScript @([PSObject]@{ Name='Jobs' ; Variable=$Jobs })



$Window.ShowDialog()

