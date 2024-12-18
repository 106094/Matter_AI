
param (
    [switch]$testing               
)

if($testing){
  [int32]$global:testing=1
}
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
Add-Type -AssemblyName Microsoft.VisualBasic,System.Windows.Forms,System.Drawing
$shell=New-Object -ComObject shell.application
$wshell=New-Object -ComObject wscript.shell
new-item -path C:\Matter_AI\logs\testing.log -Force|Out-Null
Import-Module C:\Matter_AI\Matter_functions.psm1
#region check test type
while(!$global:testtype -or ($global:testtype -ne 1 -and $global:testtype -ne 2 -and $global:testtype -ne 3)){
  $global:testtype=read-host "Which kind of testing? 1. Python 2. Manual 3. Auto (input 1, 2 or 3) (q for quit)"
  if($global:testtype -eq "q"){
    exit
  }
 }
 #endregion
#region check dut contril mode
while(!$global:dutcontrol -or ($global:dutcontrol -ne 1 -and $global:dutcontrol -ne 2 -and $global:dutcontrol -ne 3)){
  $global:dutcontrol=read-host "The DUT Reset mode is ? 1.Manual 2. Power on/off 3. Simulator switch (input 1/2/3) (q for quit)"
  if($global:dutcontrol -eq "q"){
    exit
  }
  if($global:dutcontrol -ne 1){
    $currnetset=get-content C:\Matter_AI\settings\config_linux.txt
   if (!($currnetset|Where-Object{$_ -match "serialport"})){
     $newsettings=get-content C:\Matter_Git\settings\config_linux.txt
     Compare-Object $newsettings $currnetset|where-object{$_.sideIndicator -eq "<="}|ForEach-Object{
      $newadd+=@($_.inputObject)
      add-content C:\Matter_AI\settings\config_linux.txt -value $_.inputObject
     }
     $newadds=[string]::Join("`n",$newadd)
    $messinfo="please update config_linux.txt for new settings of $newadds"
      [System.Windows.Forms.MessageBox]::Show($messinfo,"Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
     exit
   }
   dutcontrol -mode testcom
   if ($Global:seialport -ne "ok"){
    [System.Windows.Forms.MessageBox]::Show("Fail to connect SerialPort","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      exit  
   }
  }

 }
 #endregion
#region check internet connection
if (!(test-Connection "www.google.com" -count 1 -ErrorAction SilentlyContinue)) {
  $messinfo="Internet disconnected, please check internet connection"
  [System.Windows.Forms.MessageBox]::Show($messinfo,"Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  exit
}
#endregion

#region check ssh/webui connection
  $settings=get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "sship"}
  $sship=($settings.split(":"))[-1]
  #$sshusername=($settings[1].split(":"))[-1]
  write-host "check ssh ip $sship if connected"
  if (!(Test-Connection -ComputerName $sship -Count 1 -ErrorAction SilentlyContinue)) {
    $messinfo="SSH IP disconnected, please check RPI connection or SSH IP is correct"
    [System.Windows.Forms.MessageBox]::Show($messinfo,"Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    start-process C:\Matter_AI\settings\config_linux.txt
    exit
  }
  write-host "ssh ip $sship is connected"

#endregion

$ctcmds=import-csv C:\Matter_AI\settings\chiptoolcmds.csv
$global:matchcmds=$ctcmds.name|Get-Unique

$timestart=get-date
$global:puttyset = @()
$logpath="C:\Matter_AI\logs"
if(!(test-path $logpath)){
new-item -Path $logpath -ItemType directory|out-null
}

if ($global:testtype -eq 1){
 $getcmdpsfile="C:\Matter_AI\cmdcollecting_tool\Matter_getpy.ps1"
   $cmdcsvfile="C:\Matter_AI\settings\_py\py.csv"
. $getcmdpsfile
$checkfile=Get-ChildItem $cmdcsvfile|Where-Object{$_.LastWriteTime -gt $timestart}
if(!$checkfile){
    [System.Windows.Forms.MessageBox]::Show("Fail to get (update) cmd csv","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}
$caseids=(import-csv C:\Matter_AI\settings\_py\py.csv).TestCaseID
#$puttystart=1
}

if ($global:testtype -eq 3){
   $getprojects=(get-childitem C:\Matter_AI\settings\_auto\ -Directory).Name
   if (!$getprojects){
    [System.Windows.Forms.MessageBox]::Show("Please create C:\Matter_AI\settings\_auto\<Project name> include ""json.txt"" and ""xml"" folder with xml files"  ,"Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    exit
   }
 }

if ($global:testtype -eq 2){
  $getcmdpsfile="C:\Matter_AI\cmdcollecting_tool\Matter_getchiptool.ps1"
  $global:updatechiptool = [System.Windows.Forms.MessageBox]::Show("Need update UI-Manual database?", "Check", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
  if ($global:updatechiptool -eq "Yes") {
    $InfoParams = @{
      Title = "INFORMATION"
      TitleFontSize = 22
      ContentFontSize = 30
      TitleBackground = 'LightSkyBlue'
      ContentTextForeground = 'Red'
      ButtonType = 'OK'
        }
    New-WPFMessageBox @InfoParams -Content "Need About 10+ minutes to update UI-Manual database"

      }
      $getchiptool=. $getcmdpsfile
      
    if(!$getchiptool){
      [System.Windows.Forms.MessageBox]::Show("Fail to create import-excel module","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      exit  
    }

    if($getchiptool[-1].length -eq 1){
      $global:csvfilename=$getchiptool 
    }
    if(!$global:excelfile){
      [System.Windows.Forms.MessageBox]::Show("Fail to select the excel file","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      exit
      }
    if(!(test-path $global:csvfilename)){
    [System.Windows.Forms.MessageBox]::Show("Fail to get csv file","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      exit
    }    

    #region download manual speacial settings
    $goo_link="https://docs.google.com/spreadsheets/d/19ZPA2Z6SYYtvIj9qXM0FuDASF0FaZP2xjx7jcebrJEQ/"
    $gid="1307777084"
    $sv_range="A1:N1000"
    $savepath="C:\Matter_AI\settings\"
    $errormessage="matter manual set download failed"
    $checkdownload=webdownload -goo_link $goo_link -gid $gid -sv_range $sv_range -savepath $savepath -errormessage $errormessage
     if($checkdownload -match "fail"){      
      $InfoParams = @{
        Title = "INFORMATION"
        TitleFontSize = 22
        ContentFontSize = 30
        TitleBackground = 'LightSkyBlue'
        ContentTextForeground = 'Red'
        ButtonType = 'OK'
        ButtonTextForeground = "Blue"
          }
       New-WPFMessageBox @InfoParams -Content "Please login in authorized google account at edge first"
        exit
    }
    #endregion

    if(!(test-path "C:\Matter_AI\settings\chip-tool_clustercmd - id_list.csv")){
      if(test-path "C:\Matter_AI\settings\chip-tool_clustercmd*.csv"){
      remove-item "C:\Matter_AI\settings\chip-tool_clustercmd*.csv" -ErrorAction SilentlyContinue
      }
    #region download manual endpoint referance
    $goo_link="https://docs.google.com/spreadsheets/d/1-vSsxIMLxcSibvRLyez-SJD0ZfF-Su7aVUCV2bUJuWk/"
    $gid="1082391814"
    $sv_range="A1:E7000"
    $savepath="C:\Matter_AI\settings\"
    $errormessage="matter endpoint referance download failed"
    webdownload -goo_link $goo_link -gid $gid -sv_range $sv_range -savepath $savepath -errormessage $errormessage
    #endregion
    }

}

$starttime=get-date
###########################

$continueq="Yes"
while ($continueq -eq "Yes"){
  if ($global:testtype -eq 1){
    selguis -Inputdata $caseids -instruction "Please select caseids" -errmessage "No caseid selected"
    if(!$global:selss){
      [System.Windows.Forms.MessageBox]::Show("Fail to select the test case id, test will be stopped","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      $continueq=0
    }else{
      #create a log folder
      $datetime=get-date -Format yyyyMMdd_HHmmss
      $logtc="C:\Matter_AI\logs\_py\$($datetime)"
      if(!(test-path $logtc)){
        new-item -ItemType Directory -Path $logtc | Out-Null
      }
    }
    <#
    if($puttystart){
    $puttystart=0
    puttystart
    }
    #>
  }
  if ($global:testtype -eq 3){
    $getcmdpsfile="C:\Matter_AI\cmdcollecting_tool\Matter_getauto.ps1"
    . $getcmdpsfile
    if($global:excelfile.length -eq 0  -or $global:excelfile -eq 0){
      [System.Windows.Forms.MessageBox]::Show("Fail to select excel file","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      $continueq=0  
    }
    
   $getprojects=(get-childitem C:\Matter_AI\settings\_auto\ -Directory).Name
   $global:getproject=selgui -Inputdata $getprojects -instruction "Please select project" -errmessage "No project selected"
   if(!($global:getproject[-1])){
       [System.Windows.Forms.MessageBox]::Show("Fail to get project setting (folder)","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
       $continueq=0 
   }
   else{
    webuiSelections -projectname $global:getproject
    if(!$global:webuiselects){
      [System.Windows.Forms.MessageBox]::Show("Fail to get Project Name (webUI)","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      $continueq=0
   }
  }

  }
  if($global:testtype -eq 2){
    $data=Import-Csv  $global:csvfilename
    $selchek=selection_manual -data $data -column1 "catg" -column2 "TestCaseID"
    if($selchek[-1] -eq 0 -or $global:sels -match "xlsx" ){
      [System.Windows.Forms.MessageBox]::Show("Fail to select the Project/test case id, test will be stopped","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      $continueq=0 
    }
    else{
      #create a log folder
      $datetime=get-date -Format yyyyMMdd_HHmmss
      $logtc="C:\Matter_AI\logs\_manual\$($datetime)"
      if(!(test-path $logtc)){
        new-item -ItemType Directory -Path $logtc | Out-Null
      }
    }
  }
  if($continueq){
    . C:\Matter_AI\pyflow.ps1 
   #create result html
    if ($global:testtype -eq 2){
      . C:\Matter_AI\resultshtml.ps1
    }
#>
}
$endtime=get-date
$continueq = [System.Windows.Forms.MessageBox]::Show("Need Retest?", "Check", [System.Windows.Forms.MessageBoxButtons]::YesNo)
}
  
#puttyexit
if ($global:testtype -eq 2 -and !$testing){
$resultlog=(get-childitem "C:\Matter_AI\logs\_manual\" -directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1).fullname
$reportPath = join-path $resultlog "report.html"
Start-Process $reportPath -ErrorAction SilentlyContinue
}

$timepassed=New-TimeSpan -start $starttime -end $endtime
$timegap="{0} Hours, {1} minutes, {2} seconds" -f $timepassed.Hours, $timepassed.Minutes, $timepassed.Seconds
[System.Windows.Forms.MessageBox]::Show("Matter auto test completed in $timegap","Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)

if($global:dutcontrol -ne 1){
  dutcontrol -mode "close"
}
