$version=Get-ItemPropertyValue -Path 'HKCU:\SOFTWARE\Microsoft\Edge\BLBeacon' -Name "version"
$drivernow=(get-item C:\Matter_AI\cmdcollecting_tool\tool\msedgedriver.exe).VersionInfo.fileversion
if($drivernow -ne $version){
    $downloadurl = "https://msedgedriver.microsoft.com/" +  $version + "/edgedriver_win64.zip"
            
    $outputPath = "$env:userprofile\downloads\edgedriver.zip"
    Invoke-WebRequest -Uri $downloadurl -OutFile $outputPath
    Expand-Archive $outputPath -destinationpath C:\Matter_AI\cmdcollecting_tool\tool\ -Force

    remove-item $outputPath -force

}
