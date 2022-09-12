#Find updates
$Session = New-Object -ComObject Microsoft.Update.Session           
$Searcher = $Session.CreateUpdateSearcher() 
$Criteria = "IsInstalled=0"

Write-Host('Searching for Updates...') -Fore Green  
$SearchResult = $Searcher.Search($Criteria)
$Updates = $SearchResult.Updates

#Show available updates...
$Updates | fl

#Download updates
$UpdatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
$updates | % { $UpdatesToDownload.Add($_) | out-null }
Write-Host('Downloading Updates...')  -Fore Green  
$UpdateSession = New-Object -Com Microsoft.Update.Session
$Downloader = $UpdateSession.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
$Downloader.Download()

#Install updates
$UpdatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
$updates | % { if($_.IsDownloaded) { $UpdatesToInstall.Add($_) | out-null } }

#Only works installing one at a time
Write-Host("Installing $($UpdatesToInstall.Count) Updates...")  -Fore Green  
$WURebootRequired = $false
ForEach ($SingleUpdateToInstall in $UpdatesToInstall) {
	Write-Host("Installing update: " + $SingleUpdateToInstall.Title)
	$SingleUpdate = New-Object -Com Microsoft.Update.UpdateColl
	$SingleUpdate.Add($SingleUpdateToInstall)
	$Installer = $UpdateSession.CreateUpdateInstaller()
	$Installer.Updates = $SingleUpdate
	$InstallationResult = $Installer.Install()
	Write-Host("Install Result: " + $InstallationResult.ResultCode)
	if($InstallationResult.RebootRequired) {
		$WURebootRequired = $true
	}
}
if($WURebootRequired) {
	Write-Host('Reboot required! please reboot now..') -Fore Red
	$Installer.Commit(0)
	$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
	$tsenv.Value('WUUpdateReboot') = 'True'
} else {
	Write-Host('Done..') -Fore Green
}
