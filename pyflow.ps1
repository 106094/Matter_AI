$retesttimecheck=(get-content "C:\Matter_AI\settings\config_linux.txt"|Select-String "retest.") -match "\d+"
if($retesttimecheck){
  $retesttime= [int32]($matches[0])
}
else{
  $retesttime= [int32]1
}

if ($global:testtype -eq 1){

$caseids=$global:sels
$csvdata=import-csv C:\Matter_AI\settings\_py\py.csv | Where-Object {$_.TestCaseID -in $caseids -and $_.command.length -gt 0}
$settigns=import-csv C:\Matter_AI\settings\_py\settings.csv 
$headers=$settigns[0].PSObject.Properties.Name
$sound = New-Object -TypeName System.Media.SoundPlayer
$sound.SoundLocation = "C:\Windows\Media\notify.wav"
$logtc=(get-childitem -path "C:\Matter_AI\logs\_py" -directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1).fullname

foreach($csv in $csvdata){
    #$sound.Play()
    #([System.Media.SystemSounds]::Asterisk).Play()
    if($global:dutcontrol -eq 1){
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
   }
   if($global:dutcontrol -eq 2){
    $cycletime= ((get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "cycle" -and $_ -match "onoff"}) -split ":")[1]
    foreach($i in 1..$cycletime){ 
     dutcontrol -mode on
     dutcontrol -mode off
    }
    }
    if($global:dutcontrol -eq 3){
      $cycletime= ((get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "cycle" -and $_ -match "downup"}) -split ":")[1]
      foreach($i in 1..$cycletime){
        dutcontrol -mode donw
        dutcontrol -mode up   
      }
    }
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
  $lastcaseid=$null
  $specialsets=import-csv -path C:\Matter_AI\settings\*manual_special.csv
  $global:varhash=@()
  $caseids=$global:sels
  $csvdata=import-csv $global:csvfilename | Where-Object {$_.TestCaseID -in $caseids -and $_.cmd.length -gt 0}
  #$sound = New-Object -TypeName System.Media.SoundPlayer
  #$sound.SoundLocation = "C:\Windows\Media\notify.wav"
   $paring_thread="./chip-tool pairing ble-thread node-id operationalDataset --passcode --discriminator --paa-trust-store-path paapath --trace_decode 1"
   $paring_manual="./chip-tool pairing code-wifi node-id --wifi-ssid --wifi-passphrase --manual-code --paa-trust-store-path paapath --trace_decode 1"
   $paring_ble="./chip-tool pairing ble-wifi node-id --wifi-ssid --wifi-passphrase --passcode --discriminator --paa-trust-store-path paapath --trace_decode 1"
   
   if ($pairsettings."operationalDataset".length -gt 0){
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
    $puttyname0=$csv.session
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
 if($global:dutcontrol -eq 1){
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
 if($global:dutcontrol -eq 2){
    $cycletime= ((get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "cycle" -and $_ -match "onoff"}) -split ":")[1]
    foreach($i in 1..$cycletime){ 
     dutcontrol -mode on
     dutcontrol -mode off
    }
    }
 if($global:dutcontrol -eq 3){
      $cycletime= ((get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "cycle" -and $_ -match "downup"}) -split ":")[1]
      foreach($i in 1..$cycletime){
        dutcontrol -mode down
        dutcontrol -mode up   
      }
    }
   #check if putty session exist
   $puttyname=$puttyname0
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
        putty_paste -cmdline "rm -rf /tmp/chip_*" -skipcheck
        while (!$pairresult -and $k -lt $retesttime){
          $k++
          $pairresult=putty_paste -cmdline "$paringcmd" -checkline1 "Device commissioning completed with success"
          add-content -path $logpair -Value (get-content -path C:\Matter_AI\logs\lastlog.log )
          write-host "round $k"
        }

    }
    if($global:testing -eq 1){
      [int32]$pairresult=1
    }
    #start step cmd if connected pass
    if ($pairresult){ #test
      
      $k=0
      foreach($pyline in $pylines){
        $runflag=1
        $getlastkey=$null
        $addcmdall=@()
        $k++
        $puttyname=$puttyname0
          $waittime= [int64]($specialsets|Where-Object{$_.source -eq $excelfilename -and $_.TC -eq $caseid0 -and $_.step.trim() -eq $stepid -and $_.cmdline -eq $k -and $_.method -eq "waittime"}).waittime
          $specialset=$specialsets|Where-Object{$_.source -eq $excelfilename -and $_.TC -eq $caseid0 -and $_.step.trim() -eq $stepid -and $_.cmdline -eq $k}
          if ($specialset){
           foreach($special in $specialset){ 
             $puttyname=$puttyname0
             $method=$special."method"
             $newputtyname=$special."diff_session"
             $getlastkey=$special."lastlog_keyword"
             $paraname=$special."para_name"             
             $newwaittime=[int64]$special."waittime"
               if($method -match "replace"){
                 $keyword=$special."cmd_keyword"
                 $replaceby=$special."replace"
                 if($replaceby -match "var\:"){
                   $paraname=$replaceby.replace("var:","")
                   $replaceby=($global:varhash|Where-Object{$_.para_name -eq $paraname})."setvalue"
                 }
                 if($replaceby -match "py\:"){
                  $paraname=$replaceby.replace("py:","")
                  $replaceby=$pairsettings."$paraname"
                }
                 $pyline = $pyline.replace($keyword, $replaceby)
                }                           
                if($method -match "skip"){
                  $runflag=0
                }
                if($method -match "message"){
                  $message=$special."message"
                  $InfoParams = @{
                    Title = "INFORMATION" 
                    TitleFontSize = 22
                    ContentFontSize = 30
                    TitleBackground = 'LightSkyBlue'
                    ContentTextForeground = 'Red'
                    ButtonType = 'OK'
                      }
                    New-WPFMessageBox @InfoParams -Content $message
                }
                if($newputtyname.length -gt 0){
                  $puttyname=$newputtyname
                }
                if($newwaittime.length -gt 0){
                  $waittime=$newwaittime
                }
                if($method -match "add"){
                  $addcmd=$special."add_cmd"
                   #check if putty session exist
                  if($method -match "add_before"){
                    $waittime=[int64]$special."waittime"
                    $pycmd=putty_paste -cmdline "$addcmd" -check_sec $waittime -manual 
                    $lastlogcontent=get-content -path C:\Matter_AI\logs\lastlog.log
                    $datetime2=get-date -Format yyyyMMdd_HHmmss
                    $logtcstep="$tclogfd\$($datetime2)_$($caseid)_$($stepid)-$($k)_$($method).log"
                    if( $global:puttylogname.length -gt 0){
                      $logtcstep="$tclogfd\$($datetime2)_$($caseid)_$($stepid)-$($k)_$($method)_$($global:puttylogname).log"
                    }                     
                    new-item -ItemType File -Path $logtcstep | Out-Null
                    add-content -path $logtcstep -Value $lastlogcontent
                    if ($getlastkey.Length -gt 0){
                      getparameter -getlastkey $getlastkey
                      Write-Output "add var before"                     
                     }
                  }
                  if($method -match "add_after"){
                    $addcmdall+=@([PSCustomObject]@{
                      addcmdaf = $addcmd
                      puttysesstion = $puttyname
                      waittime =  $waittime
                    })
                  }


                }
              }
          }
          if($runflag -eq 1){
          $pycmd=putty_paste -cmdline "$pyline" -check_sec $waittime -manual 
          $lastlogcontent=get-content -path C:\Matter_AI\logs\lastlog.log
          
          $datetime2=get-date -Format yyyyMMdd_HHmmss
          $logtcstep="$tclogfd\$($datetime2)_$($caseid)_$($stepid)-$($k).log"
          if($global:puttylogname.length -gt 0){
            $logtcstep="$tclogfd\$($datetime2)_$($caseid)_$($stepid)-$($k)_$($global:puttylogname).log"
          }    
          new-item -ItemType File -Path $logtcstep | Out-Null
          add-content -path $logtcstep -Value $lastlogcontent

          if ($getlastkey.Length -gt 0){
           getparameter -getlastkey $getlastkey
           Write-Output "add var"
          }

         ## add after cmd
         if($addcmdall){
          $k=0
           foreach ($addcmd in $addcmdall){
            $k++
            $addcmdaf=$addcmd.addcmdaf
            #$puttysesstion=$addcmd.puttysesstion
            $waittime=[int64]$addcmd.waittime
            #$pycmd=putty_paste -cmdline "$addcmdaf" -puttyname $puttysesstion -check_sec $waittime -manual
            $pycmd=putty_paste -cmdline "$addcmdaf" -check_sec $waittime -manual
            $lastlogcontent=get-content -path C:\Matter_AI\logs\lastlog.log
            $datetime2=get-date -Format yyyyMMdd_HHmmss
            $logtcstep="$tclogfd\$($datetime2)_$($caseid)_$($stepid)-$($k)_$($method).log"
            if($global:puttylogname.length -gt 0){
              $logtcstep="$tclogfd\$($datetime2)_$($caseid)_$($stepid)-$($k)_$($method)_$($global:puttylogname).log"
            }
            new-item -ItemType File -Path $logtcstep | Out-Null
            add-content -path $logtcstep -Value $lastlogcontent
            if ($getlastkey.Length -gt 0){
              getparameter -getlastkey $getlastkey
              Write-Output "add var after"
             }
           }
         }
       }
   
    }
    }  #test
              
  }
  }
