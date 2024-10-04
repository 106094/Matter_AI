$retesttimecheck=(get-content "C:\Matter_AI\settings\config_linux.txt"|Select-String "retest.") -match "\d+"
if($retesttimecheck){
  $retesttime= [int32]($matches[0])
}
else{
  $retesttime= [int32]1
}

if ($global:testtype -eq 1){

$caseids=$global:selchek
$csvdata=import-csv C:\Matter_AI\settings\_py\py.csv | Where-Object {$_.TestCaseID -in $caseids -and $_.command.length -gt 0}
$settigns=import-csv C:\Matter_AI\settings\_py\settings.csv 
$headers=$settigns[0].PSObject.Properties.Name
$sound = New-Object -TypeName System.Media.SoundPlayer
$sound.SoundLocation = "C:\Windows\Media\notify.wav"
$logtc=(get-childitem -path "C:\Matter_AI\logs\_py" -directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1).fullname

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
      ButtonTextForeground = "Blue"
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
    copy-item C:\Matter_AI\logs\lastlog.log -Destination "$($logtc)\PASS_$($caseid)_$($datetime).log"
    }else{      
    copy-item C:\Matter_AI\logs\lastlog.log -Destination "$($logtc)\FAIL_$($caseid)_$($datetime).log"
    }
    #>
        
}
}

if ($global:testtype -eq 2){
  $excelfilename=(get-childitem $global:excelfile).Name
  $pairsettings=import-csv C:\Matter_AI\settings\_manual\settings.csv
  $headers=$pairsettings[0].PSObject.Properties.Name
  $nodeid=((get-content C:\Matter_AI\settings\config_linux.txt | Select-String "nodeid"|out-string).split(":"))[-1].trim()
  $logtc=(get-childitem -path "C:\Matter_AI\logs\_manual" -directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1).fullname

  $specialsets=import-csv -path C:\Matter_AI\settings\*manual_special.csv
  $varhash=@()
  $caseids=$global:selchek
  $csvdata=import-csv $global:csvfilename | Where-Object {$_.TestCaseID -in $caseids -and $_.cmd.length -gt 0}
  #$sound = New-Object -TypeName System.Media.SoundPlayer
  #$sound.SoundLocation = "C:\Windows\Media\notify.wav"
   $paring_thread="TBD"
   $paring_manual="./chip-tool pairing code-wifi node-id --wifi-ssid --wifi-passphrase --manual-code --paa-trust-store-path /home/ubuntu/PAA/ --trace_decode 1"
   $paring_ble="./chip-tool pairing ble-wifi node-id --wifi-ssid --wifi-passphrase --discriminator --passcode --paa-trust-store-path /home/ubuntu/PAA/ --trace_decode 1"
   
   if ($pairsettings."--thread".length -gt 0){
    $paringcmd=$paring_thread
   }
   if(!$paringcmd -and $pairsettings."--manual-code".length -gt 0){
    $paringcmd=$paring_manual
   }
   if(!$paringcmd){
    $paringcmd=$paring_ble
   }

  foreach($csv in $csvdata){
    $caseid0=$csv.TestCaseID
    $stepid=$csv.step
    $puttyname=$csv.session
    $pattern = 'TC-\w+-\d+\.\d+'
    $caseid = $null
    $check=$caseid0 -match $pattern
    if(!$check){
     $pattern = 'TC-\w+-\d+'
     $check=$caseid0 -match $pattern
    }
    $caseid = ($matches[0].replace(" ","")).trim()
    $pylines=($csv.cmd).split("`n")
    
    if($lastcaseid -ne $caseid){
        $lastcaseid=$caseid
        $tclogfd="$logtc\$($caseid)"
        if(!(test-path $tclogfd)){
          new-item -ItemType Directory -Path $tclogfd | Out-Null
        }

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
    
    #start pairing with restest
    foreach($header in $headers){
      if ($paringcmd -match $hearder){
        $paringcmd=$paringcmd -replace $header, $pairsettings."$header"
      }
    }   
    $paringcmd=$paringcmd.replace("node-id", $nodeid)
    $datetime2=get-date -Format yyyyMMdd_HHmmss
    $logpair="$tclogfd\$($datetime2)_$($caseid)_0pairing.log"    
    new-item -ItemType File -Path $logpair | Out-Null
        $k=$pairresult=0
        putty_paste -cmdline "rm -rf /tmp/chip_*"
        while (!$pairresult -and $k -lt $retesttime){
          $k++
          $pairresult=putty_paste -cmdline "$paringcmd" -line1 -1 -checkline1 "pass"
          add-content -path $logpair -Value (get-content -path C:\Matter_AI\logs\lastlog.log )
          write-host "round $k"
        }

    }
    #start step cmd if connected pass
    if ($pairresult){ #test
      $datetime2=get-date -Format yyyyMMdd_HHmmss
      $logtcstep="$tclogfd\$($datetime2)_$($caseid)_$($stepid).log"
      new-item -ItemType File -Path $logtcstep | Out-Null
     if($puttyname.length -gt 0){
        $sessionid= ($global:puttyset|Where-Object{$_.name -eq $puttyname}|Select-Object -Last 1).puttypid
        if(!$sessionid -or !($sessionid -and (get-process -id $sessionid -ErrorAction SilentlyContinue))){   
           puttystart -puttyname $puttyname
        }
      }
      $k=0
      foreach($pyline in $pylines){
        $getlastkey=0
        $k++
          $specialset=$specialsets|Where-Object{$_.source -eq $excelfilename -and $_.TC -eq $caseid -and $_.step -eq $stepid -and $_.cmdline -eq $k}
          if ($specialset){
           foreach($special in $specialset){ 
             $method=$specialset."method"
               if($method -match "replace"){
                 $keyword=$specialset."cmd_keyword"
                 $replaceby=$specialset."repalce"
                 if($replaceby -match "var\:"){
                   $paraname=$replaceby.replace("var:","")
                   $replaceby=($varhash|Where-Object{$_.para_name -eq $paraname})."setvalue"
                 }
                 $pyline = $pyline.replace($keyword, $replaceby)
               }
               if($method -match "getlastlog"){
                $getlastkey=$specialset."lastlog_keyword"
                $paraname=$specialset."para_name"
                }
              }
          }
          $pycmd=putty_paste -cmdline "$pyline" -puttyname $puttyname
          $lastlogcontent=get-content -path C:\Matter_AI\logs\lastlog.log
          add-content -path $logtcstep -Value $lastlogcontent
          if ($getlastkey){
           $matchvalue= ([regex]::Match(($lastlogcontent -match $getlastkey), "$getlastkey(.*)").Groups[1].value).tostring().trim()
           #$matchvalue= (($lastlogcontent|Select-String -Pattern "($getlastkey).*" -AllMatches |  ForEach-Object {$_.matches.value}).split($getlastkey))[-1].trim()
           $varhash+=@([PSCustomObject]@{           
            para_name = $paraname
            setvalue = $matchvalue
           })
          }
   
    }
    }  #test
              
  }
  }
