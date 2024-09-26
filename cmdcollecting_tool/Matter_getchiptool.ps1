Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
$shell=New-Object -ComObject shell.application

$global:excelfile=. "C:\Matter_AI\cmdcollecting_tool\selections_xlsx.ps1"
if($global:excelfile -eq 0){
exit
 }

$starttime=get-date

if($PSScriptRoot.length -eq 0){
  $scriptRoot="C:\Matter_AI\cmdcollecting_tool\"
  }
  else{
  $scriptRoot=$PSScriptRoot
  }

#$spath="C:\Matter_AI"
#reg insatll importexcel
$chkmod=Get-Module -name importexcel
if(!($chkmod)){
  write-host "need install importexcel"
  $PSfolder=(($env:PSModulePath).split(";")|Where-Object{$_ -match "user" -and $_ -match "WindowsPowerShell"})+"\"+"importexcel"
  $checkPSfolder=Get-ChildItem $PSfolder  -Recurse -file -Filter ImportExcel.psd1 -ErrorAction SilentlyContinue
 
 if(!($checkPSfolder)){
  New-Item -ItemType directory $PSfolder -ea SilentlyContinue|out-null
  $A1=(Get-ChildItem "$scriptRoot\cmdcollecting_tool\tool\importexcel*.zip").fullname
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
      Get-Command Import-Excel
      } catch{
     Write-Output "importexcel Package Tool install FAILED"
        }
 
 
   }
}
 #endregion

#region read PICS excusion
$pictxt="C:\Matter_AI\settings\ci-pics-values*.txt"
$picfile=(Get-ChildItem $pictxt|Sort-Object LastWriteTime|Select-Object -Last 1).FullName
if(!$picfile){
  [System.Windows.Forms.MessageBox]::Show("No PIC file, please check","Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  return "error: no PICS file"
  exit
}
$piccontent=get-content $picfile
$picexclusions=@()
foreach($line in $piccontent){
  if($line -match "\=0"){
    $excutionline=$line.replace("=0","")
    $picexclusions+=@($excutionline)
    $index=$excutionline.indexof(".")
    if($index -gt 0){
    $picexclusions+=@($excutionline.remove($index,1).insert($index,"_"))

      }
  }
}

#endregion

#region get chiptool related command
$ctcmds=import-csv C:\Matter_AI\settings\chiptoolcmds.csv
$matchcmds=$ctcmds.name|Get-Unique
#endregion
$excelfile=get-childitem -path $global:excelfile
$excelfull=$excelfile.FullName
#$excelfiles=get-childitem "C:\Matter_AI\settings\_docs\*TestPlanVerificationSteps_Auto.xlsx"
$csvname="C:\Matter_AI\settings\_manual\manualcmd_"+$excelfile.basename.replace("TestPlanVerificationSteps_Auto","")+".csv"
$TH2list="C:\Matter_AI\settings\_manual\TH2_TClist_"+$excelfile.basename.replace("TestPlanVerificationSteps_Auto","")+".txt"
if(test-path $TH2list){
  Remove-Item $TH2list -force
}
#reg read excel to csv
$columncor=(import-csv "C:\Matter_AI\settings\filesettings.csv"|Where-Object{$_.filename -eq ($excelfile).name}|Select-Object -Property column_title).column_title
$worksheetNames = (Get-ExcelSheetInfo -Path $excelfull).Name

#save parameter settings
$a=(Import-Excel $excelfull -WorksheetName "Python Script Command" -StartRow 2 -EndRow 1 -StartColumn 7)
$a[-1]|export-csv C:\Matter_AI\settings\_manual\settings.csv -NoTypeInformation -force

#filter manual and as client and UI-Manual
$sumsheetname=$worksheetNames|Where-Object{$_ -match "cert_repo"}
$worksheetsum=Import-Excel $excelfull -WorksheetName $sumsheetname
$filteredtcs = ($worksheetsum |Where-Object{$_."Test Case ID".length -gt 0}|  Where-Object {$_."$columncor" -eq "UI-Manual" `
 -and $_."Test Case Name" -notlike "*as client*"})."Test Case ID"

 $filteredsheets=$filteredtcs|foreach-object{($_.split("-"))[1]}|Get-Unique

$Indexfirst=($worksheetNames.trim()).IndexOf("ACE")
$Indexlast=($worksheetNames.trim()).IndexOf("WNCV")
$outputcsv = @()
for($i=$Indexfirst;$i -le $Indexlast;$i++){
 $thisneed=0
 $sheetname=$($worksheetNames[$i])
 ($sheetname.split("(").split(")").split(","))|Where-Object{
  if($_ -in $filteredsheets){
   $thisneed=1
  } 
 }
 #$sheetname
 if($thisneed){
 #$sheetname
 $sheetdate= Import-Excel $excelfull -WorksheetName $sheetname -NoHeader
 $worksheet = (Open-ExcelPackage -path $excelfull).Workbook.WorkSheets[$sheetname]
 $colproperty = ($sheetdate[0] | Get-Member -MemberType NoteProperty).name
 $tcline=$null
 $numbercol=$null
 $precol=$null
 $toolcmd=$null
 $row=0
 $TH2=0
  foreach($content in $sheetdate){
    $row++
    if($content -match "TH2"){
      $TH2=1
    }
     if($content -match "\[TC\-"){
    $pattern = "\[(.*?)\]"
    $match = $content | Select-String -Pattern $pattern
    $extractedText = ($match.Matches[0].Groups[1].Value).replace(" ","")
    if($extractedText -match "TC\-"){
      $tcline =$null
      $tcstep=$null
      $numbercol=$null
      $cmdcol=$null
      $results=$null
      $precon=$null
      $mergerow=$null
      $preconall=@()
     }
    
    if($extractedText -in $filteredtcs){
     ForEach($col in $colproperty){
      if(($content.$col).length -gt 0 -and ($content.$col) -match "TC\-" ){
        $tcline=($content.$col).trim()
        #$tcline        
        break
        }
       }
      }
      }
      if($tcline -and $TH2){
        if(!(test-path $TH2list)){
          new-item -ItemType File -Path $TH2list | Out-Null
        }
        $th2lists=get-content $TH2list
        if (!($tcline -in $th2lists)){
        add-content $TH2list -value $tcline
        }
        $TH2=0
      }
      if(($content -match "precondition" -or $content -match "Pre-condition") -and $tcline -and !($numbercol) -and !($precol)){
        ForEach($col in $colproperty){
          if(($content.$col) -match "precondition" -or ($content.$col) -match "Pre-condition"){
            $precol=$col
            break     
          }
        }
      }
   
      if( !($numbercol) -and ($precol)){
        $precon=$content.$precol
        if($precon -and ([int32]$(($precon|out-string).trim().length)) -gt $("precondition".length)){
        $preconall+=@($precon)
      }
  }

    if($content -match "\#" -and $content -match "Step" -and $tcline -and !($numbercol)){
      ForEach($col in $colproperty){
        if(($content.$col) -eq "#"){
          $numbercol=$col
          $matchchk=$numbercol -match "\d"
          $numbercol2=$matches[0]    
        }
        if(($content.$col) -match "verification Steps"){
          $cmdcol=$col   
        }
        if(($content.$col) -match "pics"){
          $picscol=$col
          $matchchk=$picscol -match "\d"
          $picscol2=$matches[0]
                  
        }
      }
      if($numbercol -and !$cmdcol){
        $checknumber=$numbercol -match "\d"
        $cmdcol="P"+$([int32]($matches[0])+5)
      }

      if(($preconall|out-string).trim().length -gt 0){
        $outputcsv+=[PSCustomObject]@{
          catg= $sheetname
          TestCaseID=$tcline
          step="precondition"
          flow=($preconall|out-string).trim()
          cmd=$toolcmd
          results=$results
          session=$null
          }
        $precol=$null
        }

    }
    
    if($content -match "\#" -and $content -match "Cert\sDescription"){
      $picscol=$picscol2=$null
      ForEach($col in $colproperty){
        if(($content.$col) -eq "#"){
          $numbercol=$col
          $matchchk=$numbercol -match "\d"
          $numbercol2=$matches[0]   
          $cmdcol="P7"
          break      
        }
      }
      
      if(($preconall|out-string).trim().length -gt 0){
        $outputcsv+=[PSCustomObject]@{
          catg= $sheetname
          TestCaseID=$tcline
          step="precondition"
          flow=($preconall|out-string).trim()
          cmd=$toolcmd
          results=$results          
          session=$null
          }
        $precol=$null
        }
    }
    

    if($content.$numbercol -ne "#" -and ($content.$cmdcol.Length -gt 0)){
      $pics=$content.$picscol
      $tcstep=$content.$numbercol
      $picsmerge=($worksheet.Cells[$row, $picscol2]).merge      
      #$stepmerge=($worksheet.Cells[$row, $numbercol2]).merge
      $results=$null
      if($pics.length -ne 0){
      #check if PICS not support and if merged cell
      $pics.split("(").split("&").split(" ").trim()|ForEach-Object{
          if( $_ -in $picexclusions){
            $results+=@($_)
          }
          }
          if($results){
            $results=$results -join "`n"
            $lastresult=$results
          }
          if(!$mergerow){ #find the 1st merged cell
            $mergerow=$row
          }
          if(!($picsmerge)){
            $mergerow=$null
          }
      }else{
        if(!($picsmerge)){
          $mergerow=$null
          $lastresult=$results
        }      
        else{
          if(!$mergerow){ #find the 1st merged cell which is empty
            $mergerow=$row
            $lastresult=$results
          }
          else{
            $results=$lastresult
          }         
          
         }
      }

      #incase steps is empty (merged cell)
      if($tcstep.Length -gt 0){
        $tcsteplast=$tcstep
      }else{
        $tcstep=$tcsteplast
      }

      <# Action to perform if the condition is true #>
       $tccmd=$content.$cmdcol
         # line by line check
        if($tcline -and $tccmd){
            $outputcsv+=[PSCustomObject]@{
            catg= $sheetname
            TestCaseID=$tcline
            step=$tcstep
            flow=($tccmd|out-string).trim()
            cmd=$toolcmd
            results=$results
            session=$null
            }
         }
       
        
      }
          
        
    }
   }
  }

$outputcsv|export-csv $csvname -NoTypeInformation

#region extract cmd

$csvcontent=import-csv $csvname

foreach($line in $csvcontent){
  if($line.results.length -eq 0){
   $cmd=$null
   $sessioncheck=$($line.flow) -match "\sTH_CR\d+"
   if($sessioncheck){
   $line.session=$matches[0].trim()
   }
 
   $splitcontent=$line.flow.split("`n")|Where-Object{$_.length -gt 0}
    ForEach($splitct in $splitcontent){
      if(!($splitct -match "\sTH_CR\d+") -and !($splitct -match "\sobtained\sfrom\s") -and !($splitct -match "\son\sTH1$")){
      $toolcmd=0
      if ( $splitct -match "\S+\.$"){
          $splitct=$splitct.Substring(0,$splitct.length-1)
      }
     foreach($matchcmd in $matchcmds){
      if(!$toolcmd){
      if($splitct.trim().length -gt 0 -and $splitct -match "$matchcmd\s"){
       $newmathcmds=$ctcmds|Where-Object{$_."name" -eq $matchcmd}|ForEach-Object{$_."name",$_."command" -ne "" -join "\s"}
       foreach($newmathcmd in $newmathcmds){
          if ($splitct -match "$newmathcmd\s"){
              $newcmd=@($splitct.trim())
            if (!($splitct -match "^$newmathcmd\s")){
              $pattern = "$newmathcmd(.*)"
              $match = [regex]::Match($splitct, $pattern)
              $newcmd=@($match.Value.trim())
               }
               $apostrophecheck=( $newcmd  | Select-String -Pattern "'" -AllMatches).Matches.Count
               if($apostrophecheck -eq 1){
                  $newcmd=$newcmd.replace("'","")
               }
               if($newcmd -match "\("){
                  $newcmd=($newcmd.split("("))[0].trim()
               }
             $cmd+=@($newcmd)
           $toolcmd=1 
          break
          }
        }
       } 
      }
      else{
          break
      }
     }
    }
   }
   if($cmd){
      $line.cmd=($cmd|Out-String).trim()
   }
  }
}
$csvcontent|export-csv $csvname -NoTypeInformation
#$csvcontent|Where-Object{$_.cmd.length -gt 0 -and $_.results.length -eq 0}|export-csv $csvname -NoTypeInformation
#endregion

$timegap=(new-timespan -start $starttime -end (get-date)).Minutes
$timegap2=(new-timespan -start $starttime -end (get-date)).Seconds

$checktime=[System.Windows.Forms.MessageBox]::Show("Collecting done. It took $timegap min $timegap2 sec","Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)

$csvname
#pause
#endregion