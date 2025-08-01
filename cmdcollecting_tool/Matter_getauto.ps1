Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
$shell=New-Object -ComObject shell.application

if($PSScriptRoot.length -eq 0){
  $scriptRoot="C:\Matter_AI\cmdcollecting_tool\"
  }
  else{
  $scriptRoot=$PSScriptRoot
  }

importmodule "importexcel" "import-excel" 
$excelfilename=(Get-ChildItem $global:excelfile).name
#reg read excel to csv
$columncor=((import-csv "$rootpathset\filesettings.csv"|Where-Object{$_.filename -eq $excelfilename}|Select-Object -Property webui_column_No).webui_column_No).trim()
$sumsheetname=((import-csv "$rootpathset\filesettings.csv"|Where-Object{$_.filename -eq $excelfilename}|Select-Object -Property webui_page).webui_page).trim()
#$worksheetNames = (Get-ExcelSheetInfo -Path $excelfull).Name
#$excelPackage = [OfficeOpenXml.ExcelPackage]::new((Get-Item $global:excelfile))
$worksheetsum=Import-Excel $global:excelfile -WorksheetName $sumsheetname
$columnName = ($worksheetsum[0].PSObject.Properties.Name)[[int32]$columncor-1]
$filteredtcs = ($worksheetsum |Where-Object{$_."Test Case ID".length -gt 0}|  Where-Object {$_.$columnName -eq "UI-Automated" -or $_.$columnName -eq "Verification Step Document"})."Test Case ID"|Sort-Object
#tc-filter
$tcfilters=(import-csv "$rootpathset\TC_filter.csv")
$matchtcs=($tcfilters|where-object{$_."matched_webui" -ne ""})."TC"
#$excludetcs=($tcfilters|where-object{$_."exclude_webui" -ne ""})."TC"
if($matchtcs){
  $filteredtcs= $filteredtcs|Where-Object{$_ -in $matchtcs} 
}
<#
if($excludetcs){
  $filteredtcs= $filteredtcs|Where-Object{ $_ -notin $excludetcs} 
}
#>
$global:webuicases=selguis -Inputdata $filteredtcs -instruction "Please select Auto caseids" -errmessage "No caseid selected"

<#
#endregion
$a=(Import-Excel $global:excelfile -WorksheetName $worksheetNames -StartRow 2 -EndRow 1 -StartColumn 7)
$a[-1]|export-csv $spath\settings.csv -NoTypeInformation -force
$clnsets=((Import-Excel $global:excelfile -WorksheetName "Python Script Command" -StartRow 2 -EndRow 2 -startcolumn 7)[0]).psobject.Properties | Select-Object -ExpandProperty Name
$clnsets1= @("Test Case ID" , "*sample command*") + $clnsets
$clnsets2= @("TestCaseID" , "command")+ $clnsets
$newclns=$clnsets2 -join ","
Import-Excel $global:excelfile -WorksheetName "Python Script Command" -StartRow 2 |Select-Object -Property  $clnsets1 `
|Export-Csv $spath\py0.csv -NoTypeInformation

$filtercsv=import-csv  $spath\py0.csv |Where-Object{$_."Test Case ID".length -gt 0}
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
  #>




#endregion