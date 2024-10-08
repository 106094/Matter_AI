# Path to your CSV file
$global:csvfilename
$global:sels 
$htmlContent=$null
# Import the CSV data
$csvData = Import-Csv -Path $global:csvfilename|Where-Object{$_.TestCaseID -in $global:sels}
$resultlog=(get-childitem "C:\Matter_AI\logs\_manual\" -directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1).fullname
$reportPath = "$resultlog\report.html"
$resultpaths=Get-ChildItem $resultlog -Directory
#check if there is availble test case folder
$checktcfolder=Get-ChildItem $resultlog -directory|Sort-Object LastWriteTime|Select-Object -First 1
$checktcfile=Get-ChildItem $checktcfolder.FullName -File
if($checktcfile){
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
      <table>
       <thead>
        <tr>
"@
      $headers=@("caseid","step","cmdstep","cmd","result","checks","referance")
      foreach ($header in $headers) {
      $htmlContent += "<th>$header</th>"
          }
    
$htmlContent += "</tr></thead><tbody>"

  $csvfilter=$csvData|Where-Object{$_.TestCaseID -like "*$tcname*"}
  foreach($csv in $csvfilter){
    $tcstep=$csv.step
    $logcontent=(get-childitem $fullpath -Recurse -file |Where-Object{!($_.Name -like "*0pairing*") -and  $_.name -like "*_$($tcstep)*.log"}|Sort-Object LastWriteTime).FullName
    $k=0
    foreach($log in $logcontent){
      $k++
      $logdata=get-content $log | Select-Object -skip 2|foreach-object{
       $newline=(($_ -split "CHIP\:")[1]) + "<br>"
       $newline
          }
        $refdata = $csv.flow|foreach-object{
          $newline=(($_ -split "CHIP\:")[1]) + "<br>"
          $newline
             }

      $htmlContent += "<tr>"
      $htmlContent += "<td>$($tcname)</td>"
      $htmlContent += "<td>$($csv.step)</td>"
      $htmlContent += "<td>$($k)</td>"      
      $htmlContent += "<td>$($csv.cmd)</td>" 
      $htmlContent += "<td>$($logdata)</td>"
      $htmlContent += "<td></td>"     
      $htmlContent += "<td>$($refdata)</td>"
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
