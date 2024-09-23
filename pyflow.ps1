
if ($global:testtype -eq 1){

$caseids=$selchek
$csvdata=import-csv C:\Matter_AI\settings\_py\py.csv | Where-Object {$_.TestCaseID -in $caseids}
$settigns=import-csv C:\Matter_AI\settings\_py\settings.csv 
$headers=$settigns[0].PSObject.Properties.Name
$sound = New-Object -TypeName System.Media.SoundPlayer
$sound.SoundLocation = "C:\Windows\Media\notify.wav"
$retesttimecheck=(get-content "C:\Matter_AI\settings\config_linux.txt"|Select-String "retest.") -match "\d+"
if($retesttimecheck){
  $retesttime= [int32]($matches[0])
}
else{
  $retesttime= [int32]1
}

foreach($csv in $csvdata){
    #$sound.Play()
    #([System.Media.SystemSounds]::Asterisk).Play()
    $InfoParams = @{
      Title = "INFORMATION"
      TitleFontSize = 22
      ContentFontSize = 30
      TitleBackground = 'LightSkyBlue'
      ContentTextForeground = 'Red'
      ButtonType = 'OK'
        }
  New-WPFMessageBox @InfoParams -Content "Please Reset Your DUT, then click ok"
    $caseid=($csv.TestCaseID).trim()
    $pyline=($csv.command).trim()
        #start-sleep -s 300
     # revise command
     foreach ($header in $headers){
        if($csv."$header".length -gt 0){
        $checkv=($csv."$header").trim()
         if($checkv.Length -gt 0){
            $newsettings=$checkv
           if($checkv -match "v"){
            $newsettings=$($settigns).$($header)
            $replacement = "$header $newsettings"
           }           
           elseif($checkv -match "n"){
            $replacement = ""
           }
           else{            
            $newsettings=$checkv
            $replacement = "$header $newsettings"
           }
          <#
          if($header -match "PIXIT"){
            $replacement = "$($header):$($newsettings)"
          }
          #>
         if($pyline -match "$header\s" -or $pyline -match "$header\:"){
           $pattern = "$header\s+\S+"
           if($header -match "PIXIT"){
            $header2=$header.replace(".","\.")+":"
            $pattern = "$header2\d+"
            $replacement = "$($header):$($newsettings)"          
            }

           $pyline = $pyline -replace $pattern, $replacement  
          }
          else{
            if($header -match "-N" -and $csv."$header".length -gt 0){
              $newsettings=$csv."$header"
              $pattern = ".py\s"
              $replacement = ".py " +$header +" " + $newsettings +" "
              $pyline = $pyline -replace $pattern, $replacement         
              }
            else{
            $pyline+= " $replacement"
            }
          }
          #$header
          #$replacement
          $pyline= $pyline.replace("  "," ")
          
        }
      }
     }
     #$pyline
    #
    $k=$pycmd=0
    while (!$pycmd -and $k -lt $retesttime){
      $k++
      $pycmd=putty_paste -cmdline "rm -f admin_storage.json && $pyline" -line1 -1 -checkline1 "pass"
      write-host "round $k"
    }
    $datetime=get-date -Format yyyyMMdd_HHmmss
    if ($pycmd){
    copy-item C:\Matter_AI\logs\lastlog.log -Destination C:\Matter_AI\logs\"PASS_"$($caseid)_$($datetime).log
    }else{      
    copy-item C:\Matter_AI\logs\lastlog.log -Destination C:\Matter_AI\logs\"FAIL_"$($caseid)_$($datetime).log
    }
    #>
        
}
}

if ($global:testtype -eq 2){
  $caseids=$selchek
  $csvdata=import-csv $csvname | Where-Object {$_.TestCaseID -in $caseids}
  $sound = New-Object -TypeName System.Media.SoundPlayer
  $sound.SoundLocation = "C:\Windows\Media\notify.wav"
  $retesttimecheck=(get-content "C:\Matter_AI\settings\config_linux.txt"|Select-String "retest.") -match "\d+"
  if($retesttimecheck){
    $retesttime= [int32]($matches[0])
  }
  else{
    $retesttime= [int32]1
  }
  
  foreach($csv in $csvdata){
    $caseid0=($csv.TestCaseID)
    $pattern = 'TC-\w+-\d+\.\d+'
    $caseid = $null
    $check=$caseid0 -match $pattern
    if(!$check){
     $pattern = 'TC-\w+-\d+'
     $check=$caseid0 -match $pattern
    }
    $caseid = $matches[0]
    $stepid=($csv.step)
    $pylines=($csv.command).split("`n")
    if($lastcaseid -ne $caseid){
      $lastcaseid=$caseid
      #$sound.Play()
      #([System.Media.SystemSounds]::Asterisk).Play()
      $InfoParams = @{
        Title = "INFORMATION"
        TitleFontSize = 22
        ContentFontSize = 30
        TitleBackground = 'LightSkyBlue'
        ContentTextForeground = 'Red'
        ButtonType = 'OK'
          }
    New-WPFMessageBox @InfoParams -Content "Please Reset Your DUT, then click ok"
    
    }
    foreach($pyline in $pylines){
      $k=$pycmd=0
      while (!$pycmd -and $k -lt $retesttime){
        $k++
        $pycmd=putty_paste -cmdline "rm -f admin_storage.json && $pyline" -line1 -1 -checkline1 "pass"
        write-host "round $k"
      }

      $datetime=get-date -Format yyyyMMdd_HHmmss
      if ($pycmd){
      copy-item C:\Matter_AI\logs\lastlog.log -Destination C:\Matter_AI\logs\"PASS_"$($caseid)_$($datetime).log
      }else{      
      copy-item C:\Matter_AI\logs\lastlog.log -Destination C:\Matter_AI\logs\"FAIL_"$($caseid)_$($datetime).log
      }

    }

          
  }
  }

  <#
   $date=get-date -Format yyyyMMdd_HHmmss

   $backuppath="C:\Matter_AI\_backup\_py"
   if (!(test-path $backuppath)){
     new-item -ItemType Directory -Path $backuppath | Out-Null
     }
   move-item C:\Matter_AI\settings\_py\caseids.txt -Destination $backuppath\pycaseids_$($date).txt

  #>
#renamelog -testid $($testid)
