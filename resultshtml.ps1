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
$mancaseids
$htmlContentmain=$tableContent=$htmlsub=$casecontent=$null
# Import the CSV data
$showpct=[double]((get-content C:\Matter_AI\settings\config_linux.txt|where-object {$_ -match "showpercentage"}).split(":"))[1]/100
$csvData = Import-Csv -Path $global:csvfilename|Where-Object{$_.TestCaseID -in $mancaseids}
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
$htmlContentmain = @"
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
        .highlight {
          background-color: #e6ff99;
          font-weight: bold;
        }
        .subtable {
          margin-top: 10px;
          border: 1px solid gray;
          border-collapse: collapse;
        }
        .subtable td {
          border: 1px solid gray;
          padding: 5px;
        }
        /* Container for the subtable is hidden by default */
        #subTableContainer {
          display: none;
          margin-top: 20px;
        }
        /* Styling for Pass and Fail */
        .pass {
          color: green;
          font-weight: bold;
        }
        .fail {
          background-color: red;
          color: white;
          font-weight: bold;
          padding: 5px;
          display: inline-block;
        }
        .na {
          background-color: grey;
          color: white;
          font-weight: bold;
          padding: 5px;
          display: inline-block;
        }
    </style>
</head>
<script>
function toggleSubTable(id) {
  const container = document.getElementById('subTableContainer');
  // If the same subtable is already open, toggle (hide) it.
  if (container.getAttribute('data-active') === id) {
    container.style.display = 'none';
    container.setAttribute('data-active', '');
    container.innerHTML = '';
  } else {
    // Otherwise, show the container and update its content.
    container.style.display = 'block';
    container.setAttribute('data-active', id);
    
    // Set subtable content based on the clicked id.
    let content = '';
    switch (id) {
      #casecontent
        break;
      default:
        content = '';
    }
    container.innerHTML = content;
  }
}
</script>
<body>
    <div id="top"></div>
    <h1>$resultlog</h1>
    <table>
    <thead>
      <tr>
        <th style="text-align: center;">ID</th>
        <th style="text-align: center;">Pass</th>
        <th style="text-align: center;">Fail</th>
        <th style="text-align: center;">NA</th>
        <th style="text-align: center;">Not Support</th>
        <th style="text-align: center;">Total</th>
        <th style="text-align: center;">Result</th>
      </tr>
    </thead>
    <tbody>

"@

# Add column headers (using the first object in the CSV data)

foreach($resultpath in $resultpaths){
  $fullpath=$resultpath.fullname
  $tcname=($resultpath.name).replace("_FAIL","").replace("_PASS","")
  $htmlsub+=@"
  case '$tcname':

"@
  $htmlsub+=@'
    content = `
    <table class="subtable">
    </tr><tbody>
'@
      $headers=@("caseid","step","#","cmd","runcmd","logs","varify","checks (green words)","result","pics")
      foreach ($header in $headers) {
        $htmlsub += "<th>$header</th>"
          }
    

  $csvfilter=$csvData|Where-Object{$_.TestCaseID -like "*$tcname*"}
  $totalstep=$csvfilter.step.count
  $stepcount=0
  $passcount=0
  $nacount=0
  $nscount=0
  $failcount=0
  foreach($csv in $csvfilter){
    
    $stepcount++
    $tcstep=$csv.step
    $tcsubstep=$csv.substep
    $passresult="Failed"
    $picscheck=@()
    if(($csv.pics_check) -match "0"){
      $passresult="Not Support"
     }     
    $pics=($csv.pics).split("`n")
    $picssupport=($csv.pics_check).split("`n")
    $picscount=$pics.split("`n").count
    if($csv.pics.length -gt 0){
    for($i=0;$i -lt $picscount;$i++){
      $picscheck+=($pics[$i]|out-string).trim()+" (" +($picssupport[$i]|out-string).trim()+")"
    }
    }
   $picschecks=($picscheck -join "<BR>")

    $logcontent=(get-childitem $fullpath -Recurse -file |Where-Object{!($_.Name -like "*0pairing*") -and  $_.name -like "*_$($tcstep)-$($tcsubstep)*.log"}|Sort-Object LastWriteTime).FullName
    $realcmd="-"
    $tdlog="-"
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
     $cmd=$csv.cmd
     $linec=0
     $logcount=$logcontent.count
     $cmdhighlight=$false
    foreach($log in $logcontent){
      $linec++
      $tdlog=@()
      $matchedlines=@()
      $realcmd=((get-content $log|Where-Object{$_.length -gt 0})[1]|Out-String).trim()
      if($realcmd -match "\#"){
        $realcmd=($realcmd.split("#")[1]).trim()
      }
       $checkdouble=($realcmd.split(" ")|Sort-Object|Get-Unique).count -eq ($realcmd.split(" ")).count/2
        if($checkdouble){
          $realcmd=(($realcmd.split(" "))|Select-Object -first (($realcmd.split(" ")).count/2)) -join " "
        }
        if($realcmd.trim() -ne $cmd.trim()){
          $cmdhighlight=$true
        }
      
      $logdata=get-content $log |Where-Object{$_.length -gt 0}| Select-Object -skip 2
       $newcsv=  ($csv.example).replace($ekey," ")
         foreach($ekey in $eckeys){
             $newcsv=  $newcsv.replace($ekey," ")
          }
           $checkitems = $newcsv.split("`n")|Where-Object{$_.trim().length -gt 2}|Sort-Object|Get-Unique
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
              $newcheckit=$checkkit1.replace("[","\[").replace("]","\]").replace(")","\)").replace("(","\(").replace(":","\:").replace("{","\{").replace("}","\}").replace("""","").replace(",","\,")
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
        #$linkfolder=((get-childitem $log).Directory).Name
        $tdlog = "<a href='thelinkfolder/$linkfile' target='_blank'>checklog</a><br>"
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


        
        if($checkitems.count -eq 0){
          $passresult="N/A"
        }
        if(($passmatch|where-object{$_.matched -eq "1"}).matched.count -eq $checkitems.count -and $checkitems.count -gt 0){
          $passresult="Passed"
        }
    }
    
    if($passresult -eq "Passed"){
      $passcount+=1
      $resulthtmk="<span class='pass'>Pass</span>"
    }
    if($passresult -eq "N/A"){
      $nacount+=1
      $resulthtmk="<span class='na'>N/A</span>"
    }
    if($passresult -eq "Not Support"){
      $nscount+=1
      $resulthtmk="<span class='na'>Not Support</span>"
    }
    if($passresult -eq "Failed"){
      $failcount+=1
      $resulthtmk="<span class='fail'>Fail</span>"
    }
    $htmlsub += "<tr>"
    $htmlsub += "<td class='top-align' style='width: 5%;'>$($tcname)</td>"
    $htmlsub += "<td class='top-align' style='width: 4%;'>$($tcstep)</td>"
    $htmlsub += "<td class='top-align' style='width: 1%;'>$($tcsubstep)</td>"      
    $htmlsub += "<td class='top-align' style='width: 8%;'>$($cmd)</td>"       
    if($cmdhighlight){
    $htmlsub += "<td class='top-align highlight' style='width: 8%;'>$($realcmd)</td>"
    }
    else{   
      $htmlsub += "<td class='top-align' style='width: 8%;'>$($realcmd)</td>"
    }
    $htmlsub += "<td class='top-align' style='width: 24%;'>$($tdlog)</td>"
    $htmlsub += "<td class='top-align' style='width: 18%;'>$($varify)</td>"
    $htmlsub+= "<td class='top-align' style='width: 18%;'>$($example)</td>"
    $htmlsub += "<td class='top-align' style='width: 4%;'>$($resulthtmk)</td>"            
    $htmlsub += "<td class='top-align' style='width: 6%;'>$($picschecks)</td>"
    $htmlsub += "</tr>"

if($stepcount -eq $totalstep){
  $htmlsub += @'
  </tbody>
  </table>
  <p style="text-align: left; margin-left: 20px;"><a href="#top">[Go to Top]</a></p>
  `;
  break;
'@

if($totalstep-$nacount-$nscount -eq 0){
  $newfoldername="$($tcname)_NA"  
  $tcresult="<span class='pass'>N/A</span>"
  $passsum="-"
  $failsum="-"
}
else{
  if($passcount -eq $totalstep-$nacount-$nscount){
    $newfoldername="$($tcname)_PASS"
    $tcresult="<span class='pass'>Pass</span>"
  }
  else{
    $newfoldername="$($tcname)_FAIL"
    $tcresult="<span class='fail'>Fail</span>"
  }
  $passrate=[math]::round(($passcount/($totalstep-$nacount-$nscount))*100,"0")
  $passsum="$("$passcount")"+" ($passrate%)"
  $failrate=[math]::round(($failcount/($totalstep-$nacount-$nscount))*100,"0")
  $failsum="$("$failcount")"+" ($failrate%)"
}
  rename-item  $fullpath -NewName $newfoldername -ErrorAction SilentlyContinue 
  $htmlsub=$htmlsub -replace "thelinkfolder" , $newfoldername
  
  $tableContent += @"
<tr>
<td><a href="#" onclick="toggleSubTable('$tcname'); return false;">$tcname</a></td>
<td style="text-align: center;">$passsum</td>
<td style="text-align: center;">$failsum</td>
<td style="text-align: center;">$nacount</td>
<td style="text-align: center;">$nscount</td>
<td style="text-align: center;">$totalstep</td>
<td style="text-align: center;">$tcresult</td>
</tr>
"@
}

  }
  $casecontent+=$htmlsub

  }

# Close the table and HTML structure
$tableContent += @"
</tbody>
</table>
<div id="subTableContainer" data-active="">
</div>
</body>
</html>
"@

# Output the HTML content to the file

$htmlContent=$htmlContentmain -replace "#casecontent",$casecontent
$htmlContent+=$tableContent
$htmlContent| Out-File -FilePath $reportPath -Encoding UTF8

}
# Open the HTML report (optional)
# Start-Process $reportPath
