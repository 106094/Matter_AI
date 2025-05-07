Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
$shell=New-Object -ComObject shell.application

if($PSScriptRoot.length -eq 0){
  $scriptRoot="C:\Matter_AI\cmdcollecting_tool\"
  }
  else{
  $scriptRoot=$PSScriptRoot
  }

$spath="C:\Matter_AI\settings\_py"
#reg insatll importexcel
$chkmod=Get-Module -name importexcel
if(!($chkmod)){
  ##write-host "need install importexcel"
  $PSfolder=(($env:PSModulePath).split(";")|Where-Object{$_ -match "user" -and $_ -match "WindowsPowerShell"})+"\"+"importexcel"
  $checkPSfolder=Get-ChildItem $PSfolder  -Recurse -file -Filter ImportExcel.psd1 -ErrorAction SilentlyContinue
 
 if(!($checkPSfolder)){
  New-Item -ItemType directory $PSfolder -ea SilentlyContinue|out-null
  $A1=(Get-ChildItem "$scriptRoot\tool\importexcel*.zip").fullname
  $shell.NameSpace($PSfolder).copyhere($shell.NameSpace($A1).Items(),4)
  }
 
  $checkPSfolder=Get-ChildItem $PSfolder -Recurse -file -Filter ImportExcel.psd1
 
   if(!$checkPSfolder){
   Write-Output "importexcel Package Tool unzip FAILED"
     }
 
   if(test-path "$($PSfolder)\importexcel.psd1"){
    Get-ChildItem -path $PSfolder -Recurse|Unblock-File
      Import-Module importexcel 
     
     try{ 
      Get-Command Import-Excel |out-null
      } catch{
     Write-Output "importexcel Package Tool install FAILED"
        }
 
 
   }
}
 
#endregion

#reg read excel to csv
$a=(Import-Excel $global:excelfile -WorksheetName "Python Script Command" -StartRow 2 -EndRow 1 -StartColumn 7)
$a[-1]|export-csv $spath\settings.csv -NoTypeInformation -force
if($testtype -eq 2){
write-host "get manual pairing settings"
exit
}
$clnsets=((Import-Excel $global:excelfile -WorksheetName "Python Script Command" -StartRow 2 -EndRow 2 -startcolumn 7)[0]).psobject.Properties | Select-Object -ExpandProperty Name
$clnsets1= @("Test Case ID" , "*sample command*") + $clnsets
$clnsets2= @("TestCaseID" , "command")+ $clnsets
$newclns=$clnsets2 -join ","

Import-Excel $global:excelfile -WorksheetName "Python Script Command" -StartRow 2 |Select-Object -Property  $clnsets1 `
|Export-Csv $spath\py0.csv -NoTypeInformation

$filtercsv=import-csv  $spath\py0.csv |Where-Object{$_."Test Case ID".length -gt 0}
$tcfilters=(import-csv "C:\Matter_AI\settings\manualcmd_Matter - TC_filter.csv")
$matchtcs=($tcfilters|where-object{$_."matched_py" -ne ""})."TC"
#$excludetcs=($tcfilters|where-object{$_."exclude_py" -ne ""})."TC"
if($matchtcs){
  $filtercsv=$filtercsv|Where-Object{$_."Test Case ID" -in $matchtcs}
}
<#
if($excludetcs){  
  $filtercsv=$filtercsv|Where-Object{$_."Test Case ID" -notin $excludetcs}
}
#>
$filtercsv|export-csv $spath\py0.csv -NoTypeInformation
new-item -path $spath\py.csv -force |out-null
add-content -path $spath\py.csv -value $newclns
add-content -path $spath\py.csv -value (get-content $spath\py0.csv|select-object -skip 1)
remove-item $spath\py0.csv -force

$newcontent=import-csv $spath\py.csv 
foreach($new in $newcontent){
  $new."command"=$new."command".trim()
  $splitlines= ($new."command".trim()).split("`n")
  if($splitlines.count -gt 1){
    $check=$splitlines -match "python3"
    if($check -and $check.count -eq 1 -and $check -match "^python3"){
      $new."command"=$($check)
      #$new.TestCaseID
      #$splitlines.count
      #$check
    }
   else{
    $new."command"=""
   }

  }
  }
  $newcontent|Where-Object{$_.command.length -gt 0}|export-csv $spath\py.csv -NoTypeInformation

#endregion
