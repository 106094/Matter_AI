
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
Add-Type -AssemblyName Microsoft.VisualBasic,System.Windows.Forms,System.Drawing
$shell=New-Object -ComObject shell.application
$wshell=New-Object -ComObject wscript.shell

#region check test type
while(!$global:testtype -or ($global:testtype -ne 1 -and $global:testtype -ne 2)){
  $global:testtype=read-host "Which kind of testing? 1. Python 2. Manual (input 1 or 2) (q for quit)"
  if($global:testtype -eq "q"){
    exit
  }
 }
 #endregion

 #region check ssh connection
$settings=get-content C:\Matter_AI\settings\config_linux.txt
$sship=($settings[0].split(":"))[-1]
#$sshusername=($settings[1].split(":"))[-1]
write-host "check ssh ip $sship if connected"
if (!(Test-Connection -ComputerName $sship -Count 1 -ErrorAction SilentlyContinue)) {
  $messinfo="SSH IP is not connected, please check RPI connection or SSH IP is correct"
  [System.Windows.Forms.MessageBox]::Show($messinfo,"Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  start-process C:\Matter_AI\settings\config_linux.txt
  exit
}
write-host "ssh ip $sship is connected"
#endregion


$timestart=get-date

Import-Module C:\Matter_AI\Matter_functions.psm1

$regfile="C:\Matter_AI\puttyreg.reg"
if($sship -ne "192.168.2.201"){
    (get-content "C:\Matter_AI\puttyreg.reg").replace("192.168.2.201",$sship)|Set-Content "C:\Matter_AI\puttyreg1.reg"
    $regfile="C:\Matter_AI\puttyreg1.reg"
    }

#$sshcmd=$sshusername+"@"+"$sship"
#$cdpath="cd $sshpath"

$logpath="C:\Matter_AI\logs"
if(!(test-path $logpath)){
new-item -Path $logpath -ItemType directory|out-null
}
#reg import $regfile
start-process reg -ArgumentList "import $regfile"
<#
$checksession=(Get-ChildItem HKCU:\Software\SimonTatham\PuTTY\Sessions|Select-Object -Property PSChildName).PSChildName
if((!$checksession) -or "matter" -notin $checksession){
 $checklog= (Get-ItemProperty HKCU:\Software\SimonTatham\PuTTY\Sessions\matter -name LogFileName).LogFileName
 if((!$checklog) -or $checklog -ne "C:\Matter_AI\logs\&Y&M&D&T_&H_putty.log"){
  reg import "C:\Matter_AI\puttyreg.reg"
  }
  }
$registryPath = "HKCU:\Software\SimonTatham\PuTTY\Sessions\matter"
$currentsetting=(Get-ItemProperty -Path $registryPath -Name "HostName").HostName
if($currentsetting -ne $sshcmd){
Set-ItemProperty -Path $registryPath -Name "HostName" -Value $sshcmd
 } 
 # Verify the change
 #(Get-ItemProperty -Path $registryPath -Name "HostName").HostName
 #>

write-host "add reg done"

if ($global:testtype -eq 1){
  
 $getcmdpsfile="C:\Matter_AI\cmdcollecting_tool\Matter_getpy.ps1"
  $selectionpsfile="C:\Matter_AI\selections.ps1"
   $cmdcsvfile="C:\Matter_AI\settings\_py\py.csv"

. $getcmdpsfile
$checkfile=Get-ChildItem $cmdcsvfile|Where-Object{$_.LastWriteTime -gt $timestart}
if(!$checkfile){
    [System.Windows.Forms.MessageBox]::Show("Fail to get (update) cmd csv","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}
$selchek=. $selectionpsfile
if(!$selchek){
   [System.Windows.Forms.MessageBox]::Show("Fail to select the test case id","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
   exit
}
}
if ($global:testtype -eq 2){
  $getcmdpsfile="C:\Matter_AI\cmdcollecting_tool\Matter_getchiptool.ps1"
  $result = [System.Windows.Forms.MessageBox]::Show("Need update UI-Manual database?", "Check", [System.Windows.Forms.MessageBoxButtons]::YesNo)
  if ($result -eq "Yes") {
    $InfoParams = @{
      Title = "INFORMATION"
      TitleFontSize = 22
      ContentFontSize = 30
      TitleBackground = 'LightSkyBlue'
      ContentTextForeground = 'Red'
      ButtonType = 'OK'
        }
    New-WPFMessageBox @InfoParams -Content "Need About 5 to 10 minutes to update UI-Manual database"
      
   $getchiptool=. $getcmdpsfile
   $global:csvfilename=$getchiptool[-1]
   if($global:excelfile -eq 0){
    exit
     }
     if(!(test-path $global:csvfilename)){      
      exit
     }
    }
    else{
      $excelfile=. "C:\Matter_AI\cmdcollecting_tool\selections_xlsx.ps1"
      if(!$excelfile){
        [System.Windows.Forms.MessageBox]::Show("Fail to select the excel file","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        exit
      }
    }
    $csvname="C:\Matter_AI\settings\_manual\manualcmd_"+(Get-ChildItem -path $excelfile).basename.replace("TestPlanVerificationSteps_Auto","")+".csv"
    $data=Import-Csv $csvname
    $selchek=selection_manual -data $data -column1 "catg" -column2 "TestCaseID"
    if(!$selchek){
      [System.Windows.Forms.MessageBox]::Show("Fail to select the test case id","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      exit
    }

}
###########################
$starttime=get-date
. C:\Matter_AI\putty_starting.ps1

$continueq="Yes"
while ($continueq -eq "Yes"){
. C:\Matter_AI\pyflow.ps1
$continueq = [System.Windows.Forms.MessageBox]::Show("Need test again?", "Check", [System.Windows.Forms.MessageBoxButtons]::YesNo)
if($continueq -eq "Yes"){
    $selchek
  if(!$selchek){
   [System.Windows.Forms.MessageBox]::Show("Fail to create test case id lists, test will be stopped","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
   }
  }
}
  
#puttyexit

$timepassed=New-TimeSpan -start $starttime -end (get-date)
$timegap="{0} Hours, {1} minutes, {2} seconds" -f $timepassed.Hours, $timepassed.Minutes, $timepassed.Seconds
[System.Windows.Forms.MessageBox]::Show("Matter auto test completed in $timegap","Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)