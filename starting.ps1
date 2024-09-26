﻿
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
$global:puttyset = @()
$logpath="C:\Matter_AI\logs"
if(!(test-path $logpath)){
new-item -Path $logpath -ItemType directory|out-null
}

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
$global:selchek=. $selectionpsfile
if(!$global:selchek){
   [System.Windows.Forms.MessageBox]::Show("Fail to select the test case id","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
   exit
}
#create a log folder
$datetime=get-date -Format yyyyMMdd_HHmmss
$logtc="C:\Matter_AI\logs\_py\$($datetime)"
if(!(test-path $logtc)){
  new-item -ItemType Directory -Path $logtc | Out-Null
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
    New-WPFMessageBox @InfoParams -Content "Need About 10+ minutes to update UI-Manual database"
      
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
      $global:excelfile=. "C:\Matter_AI\cmdcollecting_tool\selections_xlsx.ps1"
      if(!$global:excelfile){
        [System.Windows.Forms.MessageBox]::Show("Fail to select the excel file","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        exit
      }
      $global:csvfilename="C:\Matter_AI\settings\_manual\manualcmd_"+(get-childitem $global:excelfile).basename.replace("TestPlanVerificationSteps_Auto","")+".csv"
    }
    $data=Import-Csv  $global:csvfilename
    $selchek=selection_manual -data $data -column1 "catg" -column2 "TestCaseID"
    if(!$selchek){
      [System.Windows.Forms.MessageBox]::Show("Fail to select the test case id","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      exit
    }
#create a log folder
$datetime=get-date -Format yyyyMMdd_HHmmss
$logtc="C:\Matter_AI\logs\_manual\$($datetime)"
if(!(test-path $logtc)){
  new-item -ItemType Directory -Path $logtc | Out-Null
}
}
###########################
$starttime=get-date
puttystart

$continueq="Yes"
while ($continueq -eq "Yes"){
  if($selchek){
    . C:\Matter_AI\pyflow.ps1
  }
$endtime=get-date
$continueq = [System.Windows.Forms.MessageBox]::Show("Need test again?", "Check", [System.Windows.Forms.MessageBoxButtons]::YesNo)
if($continueq -eq "Yes"){
  if ($global:testtype -eq 1){
    $selchek=. $selectionpsfile
  }
  if ($global:testtype -eq 2){
    $selchek=selection_manual -data $data -column1 "catg" -column2 "TestCaseID"
  }
  if(!$global:selchek){
   [System.Windows.Forms.MessageBox]::Show("Fail to create test case id lists, test will be stopped","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
   }
   
  }
}
  
#puttyexit

$timepassed=New-TimeSpan -start $starttime -end $endtime
$timegap="{0} Hours, {1} minutes, {2} seconds" -f $timepassed.Hours, $timepassed.Minutes, $timepassed.Seconds
[System.Windows.Forms.MessageBox]::Show("Matter auto test completed in $timegap","Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)