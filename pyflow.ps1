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
    dutpower $global:dutcontrol
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
      $pycmd=putty_paste -cmdline "rm -f admin_storage.json && $pyline" -checkline1 "*Final result*pass*"
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
 
      dutpower $global:dutcontrol
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
                    $pycmd=putty_paste -cmdline "$addcmd" -check_sec $waittime -manual -puttyname $puttyname
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
            $pycmd=putty_paste -cmdline "$addcmdaf" -check_sec $waittime -manual -puttyname $puttyname
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

  if ($global:testtype -eq 3){
    webuiSelections
    dutpower $global:dutcontrol    
    Add-Type -AssemblyName System.Windows.Forms
  $logtc=(get-childitem -path "C:\Matter_AI\logs\_auto" -directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1).fullname
  $settings=get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "sship"}
  $sship=($settings.split(":"))[-1]
  $projname="MatterAI_"+"$($global:getproject)_"+(get-date -format "yyMMddHHmm")
  if($global:webuiselects -like "*for project*"){
    $projname=($global:webuiselects.split(":"))[-1]
    $xmlupdate=($global:webuiselects.split(":"))[0] -match "XML"
    $jsonupdate=($global:webuiselects.split(":"))[0] -match "JSON"
  }
  $fileContent=get-content C:\Matter_AI\settings\_auto\$($global:getproject)\json.txt 
  . C:\Matter_AI\cmdcollecting_tool\download_driver.ps1
  Get-ChildItem  "C:\Matter_AI\cmdcollecting_tool\tool\WebDriver.dll" |Unblock-File 
  Add-Type -Path "C:\Matter_AI\cmdcollecting_tool\tool\WebDriver.dll"

  #$driver = New-Object OpenQA.Selenium.Edge.EdgeDriver
  $optionsType = [OpenQA.Selenium.Edge.EdgeOptions]  
  $options = New-Object $optionsType
  $options.AddArgument("--disable-notifications") ## block the notification popup message
    $driverType = [OpenQA.Selenium.Edge.EdgeDriver]
      try{
          $driver = New-Object $driverType -ArgumentList $options
      }
      catch{
        write-output "fail to install web driver"
      }
    

# Create an Actions object

$actions = New-Object OpenQA.Selenium.Interactions.Actions($driver)
[OpenQA.Selenium.Interactions.Actions]$actions = New-Object OpenQA.Selenium.Interactions.Actions ($driver)
#$actions = New-Object OpenQA.Selenium.Interactions.Actions($driver)
$driver.Manage().Window.Maximize()
$driver.Navigate().GoToUrl("http://$sship")
$wait = [OpenQA.Selenium.Support.UI.WebDriverWait]::new($driver, [TimeSpan]::FromSeconds(60))
$waitfive = [OpenQA.Selenium.Support.UI.WebDriverWait]::new($driver, [TimeSpan]::FromSeconds(5))
$waitten = [OpenQA.Selenium.Support.UI.WebDriverWait]::new($driver, [TimeSpan]::FromSeconds(10))
if($global:webuiselects -eq "1"){
# Define a custom condition for the element's visibility using a ScriptBlock converted to a delegate
$addelement = $waitten.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]]{
    param ($driver)
      try{
        ($driver.FindElement([OpenQA.Selenium.By]::ClassName("icon-add-square")))
      }catch{
        return $null
      }

})

    # Perform actions on the element (if it was found)
    if ($addelement.Displayed) {
        $addelement.Click()
      }
      else{
        $addproject =  ($driver.FindElement([OpenQA.Selenium.By]::XPath('//span[text()="Add Project"]')))
        start-sleep -s 2
        $addproject.Click()
      }
    
    ($driver.FindElement([OpenQA.Selenium.By]::ClassName("p-inputtext"))).Clear()
    start-sleep -s 2
    ($driver.FindElement([OpenQA.Selenium.By]::ClassName("p-inputtext"))).SendKeys($projname)  #set project name
    start-sleep -s 5
     
     $createbt=($driver.FindElement([OpenQA.Selenium.By]::XPath('//span[text()="Create"]')))
     $driver.ExecuteScript("arguments[0].scrollIntoView(true);",  $createbt )
     start-sleep -s 5
     $createbt.click()
     start-sleep -s 10
    }
  
  if($jsonupdate -or $xmlupdate){
        $tdRow =  ($driver.FindElement([OpenQA.Selenium.By]::XPath('//td[contains(text(),$projname)]')))
        $actions.MoveToElement($tdRow).Perform() # hover to the project
         start-sleep -s 2  

    if($jsonupdate){
      ($driver.FindElement([OpenQA.Selenium.By]::XPath('//i[@ptooltip="Edit"]'))).click()
     start-sleep -s 5
      $textarea =  ($driver.FindElement([OpenQA.Selenium.By]::XPath('//textarea[@rows="10" and contains(@class, "p-inputtextarea")]')))
      start-sleep -s 2
     $textarea.Click()
     $textarea.Clear()
     Set-Clipboard -Value $fileContent
     start-sleep -s 5
     $textarea.SendKeys([OpenQA.Selenium.Keys]::Control + "v").Perform
      start-sleep -s 5
    }    
   if($xmlupdate){
     ($driver.FindElement([OpenQA.Selenium.By]::XPath('//i[@class="pi pi-upload"]'))).click()  #click update
     start-sleep -s 3     
     [System.Windows.Forms.SendKeys]::SendWait("^{l}")
     $xmlpath=join-path (join-path "C:\Matter_AI\settings\_auto" $global:getproject) "xml"
     Set-Clipboard -Value $xmlpath
     start-sleep -s 5
     [System.Windows.Forms.SendKeys]::SendWait("^{v}")
     start-sleep -s 2
     [System.Windows.Forms.SendKeys]::SendWait("~")
     start-sleep -s 2
     [System.Windows.Forms.SendKeys]::SendWait("{tab}")
     start-sleep -s 2
     [System.Windows.Forms.SendKeys]::SendWait("{tab}")
     start-sleep -s 2
     [System.Windows.Forms.SendKeys]::SendWait("{tab}")
      start-sleep -s 2
     [System.Windows.Forms.SendKeys]::SendWait(" ")
      start-sleep -s 2
     [System.Windows.Forms.SendKeys]::SendWait("s")
     start-sleep -s 2
    [System.Windows.Forms.SendKeys]::SendWait("%{o}")
    
    $xpath = '//div[contains(@class, "p-toast-top-right")]'
    # Monitor the 'z-index' value until it stops changing
    $timeout = [DateTime]::Now.AddSeconds(100)  # Set a 30-second timeout
    $currentZIndex = ""
     do{
        try{
            # Find the element
            $element =   ($driver.FindElement([OpenQA.Selenium.By]::XPath($xpath)))
            $zIndex = $element.GetAttribute("style")
            if ($zIndex -eq $currentZIndex) {
                break
            }
            $currentZIndex = $zIndex
    
        }catch {
            Write-Output "Element not found or inaccessible."
        }
        Start-Sleep -s 5
    
     }while([DateTime]::Now -lt $timeout)
    }  
    ($driver.FindElement([OpenQA.Selenium.By]::XPath('//span[text()="Update"]'))).click()
    Start-Sleep -s 10
   }
  
    # start loop
   $webv=($driver.FindElement([OpenQA.Selenium.By]::ClassName("sha-version"))).Text 
   foreach($webtc in $global:webuicases){

      if ( $global:webuicases.indexof($webtc) -ne 0){
        $driver.Navigate().GoToUrl("http://$sship")
        start-sleep -s 5
        dutpower $global:dutcontrol 
      }   
      $tclogfd="$logtc\$($webtc)"
      if (!(test-path $tclogfd)){
      new-item -ItemType Directory $tclogfd -Force|Out-Null
      }
     $webtcn=$webtc.Replace(".","_")
     start-sleep -s 5
     $tdRow =  ($driver.FindElement([OpenQA.Selenium.By]::XPath("//td[contains(text(),$projname)]")))
     $actions.MoveToElement($tdRow).Perform() # hover to the project
     start-sleep -s 2  
     ($driver.FindElement([OpenQA.Selenium.By]::XPath('//i[@ptooltip="Go To Test-Run"]'))).click()
      $element = $waitfive.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]]{
          try{
            ($driver.FindElement([OpenQA.Selenium.By]::ClassName("icon-add-square")))
          }catch{
            return $null
          }
  
      })
      
      if($element.Displayed){
        start-sleep -s 2
        $element.Click()
      }else{
        ($driver.FindElement([OpenQA.Selenium.By]::XPath('//span[text()="Add Test"]'))).click() # for 2nd run without add icon
       }
        start-sleep -s 2
        #set project name
        $testname=($driver.FindElement([OpenQA.Selenium.By]::XPath('//input[@ng-reflect-model="UI_Test_Run"]')))
        start-sleep -s 2
        $testname.Clear()
        start-sleep -s 2
        $testname.SendKeys($webtcn)  
        start-sleep -s 2
        #set tester name
        $testername=($driver.FindElement([OpenQA.Selenium.By]::XPath('//input[@placeholder="Enter Operator"]')))
        start-sleep -s 2
        $testername.Clear()
        start-sleep -s 2
        $testername.SendKeys("Allion")  
        start-sleep -s 2
        $testername.click() 
        $element = $waitfive.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]]{
          try{
            ($driver.FindElement([OpenQA.Selenium.By]::ClassName("icon-add-square")))
          }catch{
            return $null
          }
  
      })
       if($element.Displayed){
        start-sleep -s 2
        $testernadd=($driver.FindElement([OpenQA.Selenium.By]::ClassName("add-new-operator")))
        start-sleep -s 2
        $testernadd.click()
        start-sleep -s 2
      }else{
        start-sleep -s 2        
        $testername2=($driver.FindElement([OpenQA.Selenium.By]::ClassName("operator-item")))
        start-sleep -s 2
        $testername2.click()
        start-sleep -s 2
      }
       #select SDK YAML Tests   
       ($driver.FindElement([OpenQA.Selenium.By]::XPath('//span[contains(text(),"SDK YAML Tests")]'))).click()
        start-sleep -s 5
        #Clear Selection
        $clearbt=($driver.FindElement([OpenQA.Selenium.By]::XPath('//button[text()="Clear Selection "]')))
        $clearbt.Click() 
        start-sleep -s 10
        #$labelElement = ($driver.FindElement([OpenQA.Selenium.By]::XPath("//label[text()='FirstChipToolSuite']")))
        #start-sleep -s 5
        # Find the corresponding checkbox by navigating to its sibling elements
        #$checkSuite = ($labelElement.FindElement([OpenQA.Selenium.By]::XPath("//div[@class='p-checkbox-box']")))
        if($webv -like "*2a350c8*" ) {
        $checkSuite = ($driver.FindElement([OpenQA.Selenium.By]::XPath("//*[@id='p-tabpanel-2']/div/app-test-suites-list/div[2]/div[2]/div/p-checkbox")))
        $driver.ExecuteScript("arguments[0].scrollIntoView(true);", $checkSuite )
      }
      start-sleep -s 5
      $checkSuite.Click() 
      start-sleep -s 10
        #$cleartc = ($driver.FindElement([OpenQA.Selenium.By]::CssSelector("#p-tabpanel-7 > div > app-test-cases-list > div.test-name-master > div > p-checkbox > div > div.p-checkbox-box.p-highlight > span")))
        if($webv -like "*2a350c8*" ) {
        $cleartc = ($driver.FindElement([OpenQA.Selenium.By]::CssSelector("#p-tabpanel-2 > div > app-test-cases-list > div.test-name-master > div > p-checkbox > div > div.p-checkbox-box.p-highlight > span")))
        }
        $cleartc.Click()
        start-sleep -s 10
        #select TC
        #$checktc = ($driver.FindElement([OpenQA.Selenium.By]::XPath("//label[contains(text(), '$webtc')]"))) # Locate the label with the matching text
        if(-not $alltc){
         $labels = $driver.FindElements([OpenQA.Selenium.By]::XPath("//div[@class='test-name mb-10']//label"))
          foreach ($label in $labels) {
           if($label.Text.length -gt 0){
           $alltc+=@($(($label.Text.split(" "))[0]))
            }
          }
        }
        $ranktc=$alltc.IndexOf($webtc)-2
       
        $checktc= $driver.FindElement([OpenQA.Selenium.By]::CssSelector("#p-tabpanel-2 > div > app-test-cases-list > div.test-name.mb-10 > div:nth-child($ranktc) > div > p-checkbox > div > div.p-checkbox-box"))
        $driver.ExecuteScript("arguments[0].scrollIntoView(true);",  $checktc)
        $checktc.Click()
        start-sleep -s 10
        #start
        $startbt=($driver.FindElement([OpenQA.Selenium.By]::XPath('//button[text()="Start "]')))
      if($startbt.Enabled){
         $startbt.Click()
         start-sleep -s 60
        #wait run complete        
        $n=0
        do{                   
        $n++ 
          $checkcomplete = $wait.Until([System.Func[OpenQA.Selenium.IWebDriver, OpenQA.Selenium.IWebElement]]{
            try{
              ($driver.FindElement([OpenQA.Selenium.By]::ClassName("button-finish")))
            }catch{
              return $null
            }   
        }) 
        }until($checkcomplete.displayed -or $n -gt 360)
        #save log
        if($checkcomplete.displayed){
          start-sleep -s 10
        $checkcomplete.Click()
        }
        start-sleep -s 5
       #download json(report)/log
        remove-item $env:USERPROFILE\downloads\*.json -force -ea SilentlyContinue
        remove-item $env:USERPROFILE\downloads\*.log -force -ea SilentlyContinue
        $tdRow =  ($driver.FindElement([OpenQA.Selenium.By]::XPath("//td[contains(text(),$webtcn)]")))
        $actions.MoveToElement($tdRow).Perform() # hover to the project
        start-sleep -s 1
        ($driver.FindElement([OpenQA.Selenium.By]::XPath('//i[@ptooltip="Download Report"]'))).click()
        start-sleep -s 1
        ($driver.FindElement([OpenQA.Selenium.By]::XPath('//i[@ptooltip="Download Logs"]'))).click()
        do{
          start-sleep -s 1
          $dljson=(get-childitem $env:USERPROFILE\downloads\*.json).FullName         
          $dllog=(get-childitem $env:USERPROFILE\downloads\*.log).FullName
        }until($dljson -and $dllog)
        move-item -Path $dljson -Destination $tclogfd
        move-item -Path $dllog -Destination $tclogfd
       #download pdf
       $originalWindow = $driver.CurrentWindowHandle
       $tdRow =  ($driver.FindElement([OpenQA.Selenium.By]::XPath("//td[contains(text(),$webtcn)]")))
       $actions.MoveToElement($tdRow).Perform() # hover to the project
       start-sleep -s 1
      
       ($driver.FindElement([OpenQA.Selenium.By]::XPath('//i[@ptooltip="Show Report"]'))).click()
       start-sleep -s 5
       ($driver.FindElement([OpenQA.Selenium.By]::XPath('//span[text()="Print"]'))).click()
       start-sleep -s 10
       [System.Windows.Forms.SendKeys]::SendWait("^+p")
       start-sleep -s 2
       [System.Windows.Forms.SendKeys]::SendWait("{tab 6}")
       start-sleep -s 2
       [System.Windows.Forms.SendKeys]::SendWait(" ")
       start-sleep -s 5
       Set-Clipboard -Value   $webtcn
       start-sleep -s 5
       [System.Windows.Forms.SendKeys]::SendWait("^v")
       start-sleep -s 2
       [System.Windows.Forms.SendKeys]::SendWait("^{l}")
       Set-Clipboard -Value  $tclogfd
       start-sleep -s 5
       [System.Windows.Forms.SendKeys]::SendWait("^{v}")
       start-sleep -s 2
       [System.Windows.Forms.SendKeys]::SendWait("~")
       start-sleep -s 2
      [System.Windows.Forms.SendKeys]::SendWait("%{s}")
      #  close blank tab
       try {
        # Get all window handles
        $windowHandles = $driver.WindowHandles
        Write-Output "Open Window Handles: $($windowHandles -join ', ')"    
        foreach ($handle in $windowHandles) {
   
            # If the URL is 'about:blank', close this window
            if ($handle -ne  $originalWindow  ) {
                 $driver.SwitchTo().Window($handle)
                 $aittems+=@("Closing the about:blank window with handle: $handle")
               $driver.Close()
            }
        }
    
        # After closing, switch to the first remaining window
        $remainingWindowHandle = $driver.WindowHandles[0]
        $driver.SwitchTo().Window($remainingWindowHandle)
        Write-Output "Switched back to the remaining window with handle: $remainingWindowHandle"
        ($driver.FindElement([OpenQA.Selenium.By]::XPath('//button[.//span[contains(@class, "p-dialog-header-close-icon")]]'))).click()
       
    } catch {
        Write-Output "An error occurred: $($_.Exception.Message)"
    }

      }
      else{
        rename-item $tclogfd -NewName "$($webtc)_FailToStart"
      }
    



  }

  }
