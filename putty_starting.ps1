
$putty="C:\Matter_AI\putty.exe"

$settings=get-content C:\Matter_AI\settings\config_linux.txt
$pskey=($settings[2].split(":"))[-1]
#$sshpath=($settings[3].split(":"))[-1]
$fname=(Get-ChildItem $global:excelfile).name
$sshpath=(import-csv C:\Matter_AI\settings\filesettings.csv|Where-Object{$_.filename -eq $fname}).path

start-sleep -s 2
start-process $putty -ArgumentList '-load matter' -WindowStyle Maximized
#putty -load "101"
#&$putty -ssh $sshcmd -pw $sshpasswd
start-sleep -s 5
$wshell.SendKeys("raspberrypi")
start-sleep -s 2
$wshell.SendKeys("{enter}")
start-sleep -s 2
putty_paste -cmdline "sudo -s"
putty_paste -cmdline $pskey
putty_paste -cmdline "docker ps -a"
$idlogin=get-content "C:\Matter_AI\logs\lastlog.log"
$checkmatch=$idlogin -match "\/bin\/bash"
if($checkmatch){
  $ctnid= (($idlogin -match "\/bin\/bash").split(" "))[0]
}
else{
    puttyexit
}

putty_paste -cmdline "docker start $ctnid"
putty_paste -cmdline "docker exec -it $ctnid /bin/bash"
putty_paste -cmdline "cd $sshpath"

