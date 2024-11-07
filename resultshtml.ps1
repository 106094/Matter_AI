# Path to your CSV file
param (
    [string]$csvfile,
    [string]$selections
)

if($csvfile){
  $global:csvfilename=$csvfil
}
if($selections){
  $global:sels =$selections
}
$global:csvfilename
$global:sels 
$htmlContent=$null
# Import the CSV data
$showpct=[double]((get-content C:\Matter_AI\settings\config_linux.txt|where-object {$_ -match "showpercentage"}).split(":"))[1]/100
$csvData = Import-Csv -Path $global:csvfilename|Where-Object{$_.TestCaseID -in $global:sels}
$eckeys=(import-csv C:\Matter_AI\settings\report_exclude.csv).e_key|Get-Unique
$resultlog=(get-childitem "C:\Matter_AI\logs\_manual\" -directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1).fullname
$resultpaths=Get-ChildItem $resultlog -Directory |Where-Object{$_.name -ne "html"}
$reportPathlog = "$resultlog\html"
if(!(Test-Path $reportPathlog)){
  New-Item -Path $reportPathlog -ItemType Directory|out-null
}
$reportPath = "$resultlog\report.html"
#check if there is availble test case folder
$checktcfile=Get-ChildItem $resultlog -File -Recurse|where-object{$_.name -like "*.log" -and !($_.name -like "*pairing*.log") -and !($_.fullname -like "*html*")}
if($checktcfile.count -gt 0){
# Start building the HTML content
# Start building the HTML content with enhanced CSS for text wrapping
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>CSV Report</title>
    <style>
        body { font-family: Arial, sans-serif; }
        table { width: 99%; border-collapse: collapse; display: inline-block; margin: 20px; vertical-align: top; }
        th, td { 
            padding: 8px 12px; 
            border: 1px solid #ddd; 
            text-align: left; 
            word-wrap: break-word;     /* Allow wrapping within words if necessary */
            white-space: normal;       /* Allow normal wrapping */
            word-break: break-all;     /* Break long words if necessary */
            top-align { vertical-align: top; }
        }
        .top-align {
          vertical-align: top;
        }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        <meta name="viewport" content="width=device-width, initial-scale=1">
        .collapsible {
          background-color: #777;
          color: white;
          cursor: pointer;
          padding: 18px;
          width: 100%;
          border: none;
          text-align: left;
          outline: none;
          font-size: 15px;
        }

        .active, .collapsible:hover {
          background-color: #d9e2f2;
        }

        .content {
          padding: 0 18px;
          display: none;
          overflow: hidden;
        }
    </style>
</head>
<script>
  function toggleCollapsible(element) {
    // Toggle active class for the clicked element
    element.classList.toggle("active");

    // Get the next sibling element (the collapsible content)
    var content = element.nextElementSibling;

    // Toggle the display of the content
    if (content.style.display === "block") {
      content.style.display = "none";
    } else {
      content.style.display = "block";
    }
  }
</script>
<body>
    <h2>$resultlog</h2>

"@

# Add column headers (using the first object in the CSV data)

foreach($resultpath in $resultpaths){
  $fullpath=$resultpath.fullname
  $tcname=$resultpath.name
$htmlContent += @"
<a class="collapsible" href="javascript:void(0)" onclick="toggleCollapsible(this)">$tcname</a>
  <div class="content">
    <p>
      <table width="100%">
       <thead>
        <tr>
"@
      $headers=@("caseid","step","#","cmd","logs","varify","checks (green words)","result","ref")
      foreach ($header in $headers) {
      $htmlContent += "<th>$header</th>"
          }
    
$htmlContent += "</tr></thead><tbody>"

  $csvfilter=$csvData|Where-Object{$_.TestCaseID -like "*$tcname*"}
  foreach($csv in $csvfilter){    
    $tcstep=$csv.step
    $logcontent=(get-childitem $fullpath -Recurse -file |Where-Object{!($_.Name -like "*0pairing*") -and  $_.name -like "*_$($tcstep)-*.log"}|Sort-Object LastWriteTime).FullName
    $k=0
    foreach($log in $logcontent){
      $tdlog=@()
      $matchedlines=@()
      $k++
      $logdata=get-content $log |Where-Object{$_.length -gt 0}| Select-Object -skip 2
      $checkitems = $csv.example.split("`n")|Where-Object{$_.trim().length -gt 0}|Sort-Object|Get-Unique
      if($checkitems){     
        $passmatch=@()
        foreach ($checkitem in $checkitems){
          $passmatch += New-Object -TypeName PSObject -Property @{
            checkline=$checkitem
            matched=0
          }

        }
      foreach($logline in $logdata){
        if($logline -match "CHIP:" ){
          $logline=($logline -split "CHIP:")[1]
        }
            
        $j=0
        $maxdcm=0
        foreach ($checkit in $checkitems){
          $match2=$maxdcm2=0
          foreach($ekey in $eckeys){
            $checkit=$checkit.replace($ekey," ")
          }
          $splitchecjs=($checkit.split(" ")|Where-Object{[int]::TryParse($_, [ref]$null) -or $_.trim().length -gt 2}|Sort-Object|Get-Unique)
          $totalc=$splitchecjs.count
          $splitchecjs|ForEach-Object{
            $checkkit1=$_.trim()
              $newcheckit=$checkkit1.replace("[","\[").replace("]","\]").replace(")","\)").replace("(","\(").replace(":","\:").replace("{","\{").replace("}","\}")
              #if($logline -like "*$newcheckit*"){
                if($logline -match "(^|\s|\b)$newcheckit($|\s|\b)" ){
                $match2++
                $maxdcm2=[math]::round($match2/$totalc,3)   
                if($maxdcm2 -gt $maxdcm){
                   $maxdcm=$maxdcm2                
                   #Write-Output "$logline match $checkkit1 in $checkit with $maxdcm"  # Output: 12.3%                  
                   $matchgline=$checkit.ToString()
                   }
                }
           
          }
          if ($maxdcm2 -eq 1){
            $passmatch[$j].matched=1
          }
          $j++
        }
        if($maxdcm -ge $showpct){ # if setting % then log
          #$matchedlines+=@("$($match3) matched [$($matchgline.trim())], $logline")
          $maxdcmp = "{0:P1}" -f  $maxdcm
          $matchedlines+=@($logline+" ($maxdcmp)")
          }
      }
      }
        
        $htmllog=$reportPathlog+"\"+(get-childitem $log).name
        
        $linkfile=(get-childitem $log).name
        $linkfolder=((get-childitem $log).Directory).Name
        $tdlog = "<a href='$linkfolder/$linkfile' target='_blank'>checklog</a><br>"
        $matchedlines| foreach-object{
          $newline=$_
          if($_ -like "*CHIP:*"){
            $newline=(($_ -split "CHIP\:")[1])
          }
          $tdlog+=@($newline+ "<br>")
        }       
        $tdlog|Select-Object -skip 1| set-content $htmllog -force
        $tdlog=$tdlog|out-string
        <#
          $refdata = $csv.flow|foreach-object{
          $newline=(($_ -split "CHIP\:")[1]) + "<br>"
          $newline
             }
             #>
        $varify = $csv.verify.split("`n")|foreach-object{
        $newline=$_  + "<br>"
        $newline
        }
        $example = $csv.example.split("`n")|foreach-object{
          if($_.trim().length -gt 0){
            $newline=$_ + "<br>"
            $newline
          }
        }

        $cmd=($($csv.cmd) -split "`n")[$k-1]
    
        if($k -gt 1){
          #$varify=$example="Å™"
          $varify=$example="(same as last one)"
        }

        $passresult="Failed"
        if($checkitems.count -eq 0){
          $passresult="N/A"
        }
        if(($passmatch|where-object{$_.matched -eq "1"}).matched.count -eq $checkitems.count -and $checkitems.count -gt 0){
          $passresult="Passed"
        }

      $htmlContent += "<tr>"
      $htmlContent += "<td class='top-align' style='width: 5%;'>$($tcname)</td>"
      $htmlContent += "<td class='top-align' style='width: 4%;'>$($csv.step)</td>"
      $htmlContent += "<td class='top-align' style='width: 1%;'>$($k)</td>"      
      $htmlContent += "<td class='top-align' style='width: 10%;'>$($cmd)</td>"
      $htmlContent += "<td class='top-align' style='width: 36%;'>$($tdlog)</td>"
      $htmlContent += "<td class='top-align' style='width: 18%;'>$($varify)</td>"
      $htmlContent += "<td class='top-align' style='width: 18%;'>$($example)</td>"
      $htmlContent += "<td class='top-align' style='width: 4%;'>$($passresult)</td>"      
      $htmlContent += "<td class='top-align' style='width: 4%;'></td>"
      $htmlContent += "</tr>"
    }    
  }
$htmlContent += @"
 </tbody>
 </table>
 </p>
 </div>
 
"@
  }

# Close the table and HTML structure
$htmlContent += @"
</body>
</html>
"@

# Path where you want to save the HTML report

# Output the HTML content to the file
$htmlContent | Out-File -FilePath $reportPath -Encoding UTF8

}
# Open the HTML report (optional)
# Start-Process $reportPath
