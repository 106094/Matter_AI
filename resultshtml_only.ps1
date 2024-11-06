Import-Module C:\Matter_AI\Matter_functions.psm1
$getcmdpsfile="C:\Matter_AI\cmdcollecting_tool\Matter_getchiptool.ps1"
$global:csvfilename=. $getcmdpsfile
$data=Import-Csv  $global:csvfilename
$global:sels =selection_manual -data $data -column1 "catg" -column2 "TestCaseID"
. C:\Matter_AI\resultshtml.ps1