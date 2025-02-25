Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
$shell=New-Object -ComObject shell.application

$starttime=get-date

if($PSScriptRoot.length -eq 0){
  $scriptRoot="C:\Matter_AI\cmdcollecting_tool\"
  }
  else{
  $scriptRoot=$PSScriptRoot
  }

#$spath="C:\Matter_AI"

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

$excelfile=get-childitem -path $global:excelfile
$excelfull=$excelfile.FullName

#$excelfiles=get-childitem "C:\Matter_AI\settings\_docs\*TestPlanVerificationSteps_Auto.xlsx"
$csvname0="C:\Matter_AI\settings\_manual\manualcmd_"+$excelfile.basename.replace("TestPlanVerificationSteps_Auto","")+"0.csv"
$csvname1="C:\Matter_AI\settings\_manual\manualcmd_"+$excelfile.basename.replace("TestPlanVerificationSteps_Auto","")+"1.csv"
$csvname="C:\Matter_AI\settings\_manual\manualcmd_"+$excelfile.basename.replace("TestPlanVerificationSteps_Auto","")+".csv"
$TH2list="C:\Matter_AI\settings\_manual\TH2_TClist_"+$excelfile.basename.replace("TestPlanVerificationSteps_Auto","")+".txt"
if(test-path $TH2list){
  Remove-Item $TH2list -force
}
#save parameter settings
$a=(Import-Excel $excelfull -WorksheetName "Python Script Command" -StartRow 2 -EndRow 1 -StartColumn 7)
$a[-1]|export-csv C:\Matter_AI\settings\_manual\settings.csv -NoTypeInformation -force

#tc-filter
$tcfilters=(import-csv "C:\Matter_AI\settings\manualcmd_Matter - TC_filter.csv")
$matchtcs=($tcfilters|where-object{$_."matched_manual" -ne ""})."TC"
#$extratcs=($tcfilters|where-object{$_."extra_manual" -ne ""})."TC"
#$excludetcs=($tcfilters|where-object{$_."exclude_manual" -ne ""})."TC"

#filter manual and as client and UI-Manual
if ($global:updatechiptool -eq "Yes"){
#region insatll importexcel
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
        Get-Command Import-Excel  |out-null
        } catch{
       Write-Output "importexcel Package Tool install FAILED"
         }
   }
}
 #endregion

  #region get chiptool related command
  $ctcmds=import-csv C:\Matter_AI\settings\chiptoolcmds.csv
  $matchcmds=$ctcmds.name|Get-Unique
  #endregion
#reg read excel to csv
$columncor=((import-csv "C:\Matter_AI\settings\manualcmd_Matter - filesettings.csv"|Where-Object{$_.filename -eq ($excelfile).name}|Select-Object -Property manual_column_title).manual_column_title).trim()
$sumsheetname=((import-csv "C:\Matter_AI\settings\manualcmd_Matter - filesettings.csv"|Where-Object{$_.filename -eq ($excelfile).name}|Select-Object -Property manual_page).manual_page).trim()
$worksheetNames = (Get-ExcelSheetInfo -Path $excelfull).Name
#$sumsheetname=$worksheetNames|Where-Object{$_ -match "cert_repo"}

$excelPackage = [OfficeOpenXml.ExcelPackage]::new((Get-Item $excelfull))
$worksheetsum=Import-Excel $excelfull -WorksheetName $sumsheetname
$filteredtcs = ($worksheetsum |Where-Object{$_."Test Case ID".length -gt 0}|  Where-Object {($_."$columncor" -eq "UI-Manual" -or $_."$columncor" -eq "UI-Semi-automated") `
 -and $_."Test Case Name" -notlike "*as client*"})."Test Case ID"
 if($extra){
  $filteredtcs+=$extra
 }
$filteredsheets=$filteredtcs|foreach-object{($_.split("-"))[1]}|Sort-Object|Get-Unique
$filteredsheets+="Diag Log"
$Indexfirst=($worksheetNames.trim()).IndexOf("ACE")
$Indexlast=($worksheetNames.trim()).IndexOf("WNCV")
$outputcsv = @()
for($i=$Indexfirst;$i -le $Indexlast;$i++){
 $thisneed=0
 $sheetname=$($worksheetNames[$i])
 ($sheetname.split("(").split(")").split(","))|Where-Object{
  if($_.trim() -in $filteredsheets){
   $thisneed=1
  } 
 }
 #$sheetname
 if($thisneed){
 #$sheetname
 
 $sheetpackage = $excelPackage.Workbook.Worksheets[$sheetname]
 $sheetdate= Import-Excel $excelfull -WorksheetName $sheetname -NoHeader
 #$worksheet = (Open-ExcelPackage -path $excelfull).Workbook.WorkSheets[$sheetname]
 $colproperty = ($sheetdate[0] | Get-Member -MemberType NoteProperty).name
 $tcline=$null
 $numbercol=$null
 $precol=$null
 $toolcmd=$null
 $row=1
 $TH2=0
  foreach($content in $sheetdate){
   
    if($content -match "TH2"){
      $TH2=1
    }
    if($content -match "\[TC\-"){
    $pattern = "\[(.*?)\]"
    $match = $content | Select-String -Pattern $pattern
    $extractedTextg=($match.Matches[0].Groups[1].Value)
    $extractedText = $extractedTextg.replace(" ","")
     if($extractedText -match "TC\-"){
      $tcline =$null
      $tcstep=$null
      $numbercol=$null
      $cmdcol=$null
      $picsnames=$null
      $pics_checks=$null
      $precon=$null
      $mergerow=$null
      $preconall=@()
      $TH2=0
      }
    
    if($extractedText -in $filteredtcs){
     ForEach($col in $colproperty){
      if(($content.$col).length -gt 0 -and ($content.$col) -match "TC\-" ){
        $tcline=($content.$col).trim()
        $tcline=$tcline.replace($extractedTextg,$extractedText)
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
        $TH2tcline+=@($tcline)
        add-content $TH2list -value $tcline
        }
        $TH2=0
      }
      if(($content -match "precondition" -or $content -match "Pre-condition" -or $content -match "Pre\scondition") -and $tcline -and !($numbercol) -and !($precol)){
        ForEach($col in $colproperty){
          if(($content.$col) -match "precondition" -or ($content.$col) -match "Pre-condition" -or ($content.$col) -match "Pre\scondition"){
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
          substep=1
          cmd=$toolcmd
          verify=$null          
          flow=($preconall|out-string).trim()
          example=$null          
          pics=$picsnames
          pics_check=$pics_checks
          session=$null
          }
        $precol=$null
        }

    }
    
    #for TC-DA-1.8 special format
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
          substep=1
          cmd=$toolcmd
          verify=$null
          example=$null
          flow=($preconall|out-string).trim()          
          pics=$picsnames
          pics_check=$pics_checks
          session=$null
          }
        $precol=$null
        }
    }
    

    if($content.$numbercol -ne "#" -and ($content.$cmdcol.Length -gt 0)){
      $pics=$content.$picscol
      $tcstep=$content.$numbercol
      $picsmerge=($sheetpackage.Cells[$row, $picscol2]).merge  
      #$stepmerge=($worksheet.Cells[$row, $numbercol2]).merge
      $picsnames=$pics_checks=$null
      if($pics.length -ne 0){
      #check if PICS not support and if merged cell
      $pics.split("(").split("&").split(" ").trim()|ForEach-Object{
        if($_ -match "\."){
          $picsnames+=@($_)
          $pics_check="1"           
          if( $_ -in $picexclusions){
            $pics_check="0"  
          }
          $pics_checks+=@($pics_check)
           }
           }
          if($picsnames){
            $picsnames=($picsnames|out-string).trim()
            $lastpicsnames=$picsnames
            $pics_checks= ($pics_checks|out-string).trim()
            $lastpics_checks=$pics_checks
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
          $lastpicsnames=$picsnames
          $lastpics_checks=$pics_checks
        }      
        else{
          if(!$mergerow){ #find the 1st merged cell which is empty
            $mergerow=$row
            $lastpicsnames=$picsnames
            $lastpics_checks=$pics_checks
          }
          else{
            $picsnames=$lastpicsnames
            $pics_checks=$lastpics_checks
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
       $tccmd=($content.$cmdcol|out-string).trim()
       $checklines=($sheetpackage.Cells[$row,[int32]($cmdcol.replace("P",""))].RichText|Where-Object{$_.Color.G -gt 100 -and $_.Color.G -gt $_.Color.R}).Text
       $example=$null
       if($checklines){
       $example=($checklines.split("`n")|ForEach-Object{
        $newline=$_
        if($_ -like "*CHIP:*"){
            $newline=($_ -split "CHIP\:")[1]
          }
          $newline
        }  | out-string).trim()
      }
        # line by line check
        if($tcline -and $tccmd){
            $outputcsv+=[PSCustomObject]@{
            catg= $sheetname
            TestCaseID=$tcline
            step=$tcstep
            substep=1
            cmd=$toolcmd
            verify=$null
            example=$example
            flow=$tccmd
            pics=$picsnames
            pics_check=$pics_checks
            session=$null
            }
         }
       
        
      }
          
      $row++  
    }
   }
  }

$outputcsv|export-csv $csvname0 -NoTypeInformation

#region extract cmd

$csvcontent=import-csv $csvname0
$newcsvcontent=@()
foreach($line in $csvcontent){
  $stepnew=$null
  $newcsvcontent+=@($line)
  $tcnow=$line.TestCaseID
  $stepnow=$line.step
  $substepnew=($newcsvcontent|Where-Object{$_.TestCaseID -eq $tcnow -and ($_.step.split("."))[0] -eq $stepnow}).substep.count
  if($substepnew -gt 1){
    $stepnew="$($stepnow).$($substepnew)"
  }
 # if(!($line.pics_check  -like "*0*")){
   $cmd=$null
   $sessioncheck=$($line.flow) -match "\sTH_CR\d+"
   if($sessioncheck){
    ($newcsvcontent[-1]).session=$matches[0].trim()
   }
   
   $stepexample=($line.example).split("`n")
    $splitcontent=($line.flow).split("`n")|Where-Object{$_.length -gt 0}
      $linec=0
    $splitline=@()
    ForEach($splitct in $splitcontent){
      if(!($splitct -match "\sTH_CR\d+") -and !($splitct -match "\sobtained\sfrom\s") -and !($splitct -match "\son\sTH1")){
      $toolcmd=0
      if ( $splitct -match "\S+\.$"){
          $splitct=$splitct.Substring(0,$splitct.length-1)
      }
     foreach($matchcmd in $matchcmds){
      if(!$toolcmd){
      if($splitct.trim().length -gt 0 -and $splitct -match "$matchcmd\s"){
       $newmathcmds=$ctcmds|Where-Object{$_."name" -eq $matchcmd}|ForEach-Object{$_."name",$_."command" -ne "" -join "\s+"}
       foreach($newmathcmd in $newmathcmds){
        $matchword="$newmathcmd\s"
        if ($matchcmd -match "avahi\-browse"){
          $matchword=$newmathcmd.replace("-","\-").replace(".","\.")
        }
          if ($splitct -match "$matchword"){
              $newcmd=@($splitct.trim())
              if (!($splitct -match "^$matchword")){
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
             $splitline+=@($linec)
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
    $linec++
   }

   if($cmd){
    $cmdcount=0
    foreach($cmdx in $cmd){
      $stepes=@()    
      $verifying=$null
         if($cmd.count -eq 1 -or ($cmdcount -eq $cmd.count -1) ){
        $lineend=$splitcontent.count -1
      }
      else{
        $lineend=$splitline[$cmdcount+1] -1
      }
      $cmdflow= $splitcontent[($splitline[$cmdcount])..($lineend)]
      foreach($stepe in $stepexample){
          #The [ character is a special character that begins a character set
          $steper=$stepe -replace '\[', '[[]'
        if(($cmdflow|out-string) -like "*$steper*"){         
          $stepes+=@($stepe)
        }
      }     
      $stepex=$stepes -join "`n"
      #region varify wordings
      $p=0
      $verifying=$cmdflow|ForEach-Object{
       if($_ -like "*verify*"){
            $p=9999
        }
        $p++
        if($p -gt 9999 -and $_.length -gt 0 -and $_ -notlike "*CHIP:*"){
          $recordflag=$true
            $cmd3=$cmdx.replace("[","*").replace("]","*")
            if($_ -like "*$cmd3*"){
             $recordflag=$false
            }
          if($recordflag){
            $_
          }
       }
     }

    #endregion

    if($cmdcount -gt 0){
      $addline=$line|Select-Object *
      $newcsvcontent+=$addline
     }
     if($stepnew){
      $newcsvcontent[-1].step=$stepnew
     }
     $newcsvcontent[-1].substep=$cmdcount+1
     $newcsvcontent[-1].cmd=($cmdx|Out-String).trim()
     $newcsvcontent[-1].example=($stepex|Out-String).trim()
     $newcsvcontent[-1].flow=($cmdflow|Out-String).trim()
     if($verifying){
       $newcsvcontent[-1].verify=($verifying|Out-String).trim()
      }

    $cmdcount++
    }
  }
 
}
$newcsvcontent| export-csv $csvname1 -NoTypeInformation

$timegap=(new-timespan -start $starttime -end (get-date)).Minutes
$timegap2=(new-timespan -start $starttime -end (get-date)).Seconds

$checktime=[System.Windows.Forms.MessageBox]::Show("Collecting done. It took $timegap min $timegap2 sec","Info",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)

}
#endregion

#region filter
$csvcontent=Import-Csv $csvname1
$csvcontentnew=@()
#$filters=($matchtcs+$extratcs)|Sort-Object|Get-Unique
$filters=$matchtcs|Sort-Object|Get-Unique
foreach($filter in $filters){
  $csvcontentnew+=$csvcontent|Where-Object{$_."TestCaseID" -match "\[$filter\]"}
}

<#
if($excludetcs){
  foreach($excludetc in $excludetcs){
    $csvcontentnew=$csvcontentnew|Where-Object{$_."TestCaseID" -notmatch "\[$excludetc\]"}
  }
}
#>

$csvcontentnew|Where-Object{$_.cmd.length -gt 0 -and $_.results.length -eq 0 -and $_.TestCaseID -notin $TH2tcline} |export-csv $csvname -NoTypeInformation
#endregion


$csvname
#pause
#endregion
