
#region windows functions
$cSource = @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class Clicker
{
//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646270(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct INPUT
{ 
    public int        type; // 0 = INPUT_MOUSE,
                            // 1 = INPUT_KEYBOARD
                            // 2 = INPUT_HARDWARE
    public MOUSEINPUT mi;
}

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646273(v=vs.85).aspx
[StructLayout(LayoutKind.Sequential)]
struct MOUSEINPUT
{
    public int    dx ;
    public int    dy ;
    public int    mouseData ;
    public int    dwFlags;
    public int    time;
    public IntPtr dwExtraInfo;
}

//This covers most use cases although complex mice may have additional buttons
//There are additional constants you can use for those cases, see the msdn page
const int MOUSEEVENTF_MOVED      = 0x0001 ;
const int MOUSEEVENTF_LEFTDOWN   = 0x0002 ;
const int MOUSEEVENTF_LEFTUP     = 0x0004 ;
const int MOUSEEVENTF_RIGHTDOWN  = 0x0008 ;
const int MOUSEEVENTF_RIGHTUP    = 0x0010 ;
const int MOUSEEVENTF_MIDDLEDOWN = 0x0020 ;
const int MOUSEEVENTF_MIDDLEUP   = 0x0040 ;
const int MOUSEEVENTF_WHEEL      = 0x0080 ;
const int MOUSEEVENTF_XDOWN      = 0x0100 ;
const int MOUSEEVENTF_XUP        = 0x0200 ;
const int MOUSEEVENTF_ABSOLUTE   = 0x8000 ;

const int screen_length = 0x10000 ;

//https://msdn.microsoft.com/en-us/library/windows/desktop/ms646310(v=vs.85).aspx
[System.Runtime.InteropServices.DllImport("user32.dll")]
extern static uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

public static void LeftClickAtPoint(int x, int y)
{
    //Move the mouse
    INPUT[] input = new INPUT[3];
    input[0].mi.dx = x*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width);
    input[0].mi.dy = y*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height);
    input[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;
    //Left mouse button down
    input[1].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    //Left mouse button up
    input[2].mi.dwFlags = MOUSEEVENTF_LEFTUP;
    SendInput(3, input, Marshal.SizeOf(input[0]));
}
public static void rightClickAtPoint(int x, int y)
{
    //Move the mouse
    INPUT[] input = new INPUT[3];
    input[0].mi.dx = x*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width);
    input[0].mi.dy = y*(65535/System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height);
    input[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;
    //Left mouse button down
    input[1].mi.dwFlags = MOUSEEVENTF_RIGHTDOWN;
    //Left mouse button up
    input[2].mi.dwFlags = MOUSEEVENTF_RIGHTUP;
    SendInput(3, input, Marshal.SizeOf(input[0]));
}
}
'@
Add-Type -TypeDefinition $cSource -ReferencedAssemblies System.Windows.Forms,System.Drawing

Add-Type @"
              using System;
              using System.Runtime.InteropServices;
              public class Window {
                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
                [DllImport("User32.dll")]
                public extern static bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);
              }
              public struct RECT
              {
                public int Left;        // x position of upper-left corner
                public int Top;         // y position of upper-left corner
                public int Right;       // x position of lower-right corner
                public int Bottom;      // y position of lower-right corner
              }
"@
#endregion

#region putty cmd and check    
function putty_paste([string]$puttyname,[string]$cmdline,[int64]$check_sec,[int64]$line1,[string]$checkline1,[int64]$line2,[string]$checkline2,[switch]$manual,[switch]$skipcheck){
    if($check_sec -eq 0){$check_sec = 1}
    $global:puttylogname=$puttynamedest=$null
    if ($manual -and !$skipcheck){
        $lastword=$endpoint=$destid=$endpoint=$null
        $eplists=import-csv "C:\Matter_AI\settings\chip-tool_clustercmd - id_list.csv"
        $splitcmd=($cmdline.replace("./chip-tool ","")).split(" ")|where-object{$_.Length -gt 0}
        #checkif  attribute with ''
        if ($cmdline -match "'([^']*)'") {
            $attribute ="'"+ $matches[1]+"'"
        }
        #check if matched chiptool cmd
        $matchline=$eplists|Where-Object{$_.name -eq $splitcmd[0] -and $_.command -eq $splitcmd[1] -and $_.attribute -eq $splitcmd[2]}
        if($matchline){
            $endpoint=$matchline.endpoint
            $destid=$matchline."destination-id"
            $lastword=$splitcmd[2]
            $laststring= [regex]::Escape($lastword)
            #$pattern = "$laststring\s+(\S+)\s+(\S+)\s+(\S+)"
        }
        else{
         $matchline=$eplists|Where-Object{$_.name -eq $splitcmd[0] -and $_.command -eq $splitcmd[1]}
         $endpoint=$matchline.endpoint
         $destid=$matchline."destination-id"
         $lastword=$splitcmd[1]
         $laststring= [regex]::Escape($lastword)
         }
        $maxid=(@($endpoint,$destid)|Measure-Object -maximum).maximum
        $patterns="\s+(\S+)"*$maxid
        $pattern = "$laststring$($patterns)"

        if ($matchline){
            $matchData = @()  # Array to store match information
            $matches = [regex]::Match($cmdline +" abcd efgh 1234", $pattern) # for lack after id
            if($attribute){
                $matches = [regex]::Match( $cmdline.replace($attribute,"att1") +" abcd efgh 1234", $pattern) # for lack after id   
            }
                if ($matches.Success) {
                    foreach ($match in $matches.Groups) {
                        $matchInfo = [PSCustomObject]@{
                            Value = $match.Value
                            Index = $match.Index
                            Length = $match.length
                        }
                        $matchData += $matchInfo
                    }
                }
        
          if($destid.trim().length -gt 0){
             $destnodeid= $matchData[$destid].Value
             $puttynamedest="session$($destnodeid)"
           }
              if($endpoint){   
            $endpid0=((get-content C:\Matter_AI\settings\config_linux.txt | Select-String "endpoint0"|out-string).split(":"))[-1].trim()
            $endpid1=((get-content C:\Matter_AI\settings\config_linux.txt | Select-String "endpoint1"|out-string).split(":"))[-1].trim()
            $endpid2=((get-content C:\Matter_AI\settings\config_linux.txt | Select-String "endpoint2"|out-string).split(":"))[-1].trim()
            $endpid3=((get-content C:\Matter_AI\settings\config_linux.txt | Select-String "endpoint3"|out-string).split(":"))[-1].trim()
            if($endpid0 -ne 0 -or $endpid1 -ne 1 -or $endpid2 -ne 2){                  
             if ($matchData.Count -ge $endpoint) {
                $matched= $matchData[$endpoint].Value
                $numberIndex = $matchData[$endpoint].Index
                $numberLength = $matchData[$endpoint].Length
                if($endpid0 -and $endpid0 -ne 0 -and $matched -eq 0){
                $cmdline = $cmdline.Substring(0, $numberIndex) + $endpid0 + $cmdline.Substring($numberIndex + $numberLength)   
                }          
                if($endpid1 -and $endpid1 -ne 1 -and $matched -eq 1){
                $cmdline = $cmdline.Substring(0, $numberIndex) + $endpid1 + $cmdline.Substring($numberIndex + $numberLength)   
                } 
                if($endpid2 -and $endpid2 -ne 2 -and $matched -eq 2){
                    $cmdline = $cmdline.Substring(0, $numberIndex) + $endpid2 + $cmdline.Substring($numberIndex + $numberLength)   
                } 
                if($endpid3 -and $endpid3 -ne 3 -and $matched -eq 3){
                    $cmdline = $cmdline.Substring(0, $numberIndex) + $endpid3 + $cmdline.Substring($numberIndex + $numberLength)   
                } 
              }
            }
          }
        }
       #replace hardcode
        if($cmdline -like "*pairing*" -and $cmdline -like "*gamma*"){
            $pairsettings=import-csv C:\Matter_AI\settings\_manual\settings.csv
            $storepath=$pairsettings."paapath"
            $cmdline=$cmdline.replace("gamma","gamma --paa-trust-store-path $storepath --trace_decode 1")
        }
        
        #replace hardcode
        if($cmdline -like "*pairing*" -and $cmdline -like "*beta*"){
            $pairsettings=import-csv C:\Matter_AI\settings\_manual\settings.csv
            $storepath=$pairsettings."paapath"
            $cmdline=$cmdline.replace("beta","beta --paa-trust-store-path $storepath --trace_decode 1")
        }
    }   


if($puttynamedest.length -gt 0){
    $global:puttylogname=$puttynamedest
}
 
#for special cmd "interactive start" without session info
if($cmdline -match "avahi\-browse" ){
    $global:puttylogname="session1"
}
#for special cmd "interactive start" without session info
if( $cmdline -match "interactive\sstart"){
    $global:puttylogname="session1"
    if($cmdline -match "commissioner\-name\sbeta"){
        $global:puttylogname="session2"
    } 
    if($cmdline -match "commissioner\-name\sgamma"){
        $global:puttylogname="session3"
    }
}

#first priority
if($puttyname.length -gt 0){
    $global:puttylogname=$puttyname
}

puttystart -puttyname $global:puttylogname
    if($global:puttylogname.length -eq 0){
        #$pidd=(get-process putty|Sort-Object StartTime|Select-Object -Last 1).Id
        $logputty="C:\Matter_AI\logs\*putty.log"      
    }
    else{
        $logputty="C:\Matter_AI\logs\*putty_$($global:puttylogname).log"
    }


$pidd=($global:puttyset|Where-Object{$_.name -eq $global:puttylogname}|Select-Object -last 1).puttypid
if(!$pidd){$pidd=$global:puttyset[0].puttypid}
add-content C:\Matter_AI\logs\testing.log -value "$global:puttylogname (pid id: $pidd): $cmdline"
if($global:testing){
return
 }

$logfile=(Get-ChildItem $logputty|Sort-Object LastWriteTime|Select-Object -last 1).fullname
$checkend=((get-content $logfile)[-1]|Out-String).Trim()
#start-process notepad $logfile -WindowStyle Minimized
#start-sleep -s 3
#(get-process notepad).CloseMainWindow()|Out-Null
$lastlogline=(get-content $logfile).count -1
[Microsoft.VisualBasic.interaction]::AppActivate($pidd)|out-null
Set-Clipboard -Value $cmdline
start-sleep -s 3
$Handle= (get-process -id $pidd).MainWindowHandle 
$WindowRect = New-Object RECT
$GotWindowRect = [Window]::GetWindowRect($Handle, [ref]$WindowRect)
#Write-Host $WindowRect.Left $WindowRect.Top $WindowRect.Right $WindowRect.Bottom
##scale
#$bdh=(([System.Windows.Forms.Screen]::AllScreens|Select-Object Bounds).Bounds).Bottom 
#$height  = ([string]::Join("`n", (wmic path Win32_VideoController get CurrentVerticalResolution))).split("`n") -match "\d{1,}"
#$sacle=$height[0]/$bdh[0]
$sacle=1 ## for command use, no need to divided with scale ( don't know the reason yet)
$x1=[math]::Round(($WindowRect.Left + $WindowRect.Right)/2/$sacle,0)
$y1=[math]::Round(($WindowRect.Top + $WindowRect.Bottom)/2/$sacle,0)
[Clicker]::LeftClickAtPoint($x1, $y1)
start-sleep -s 1
if($checkend -eq ">>>" -and $cmdline -match "\./chip\-tool\s"){
    $wshell.SendKeys("^c")
    start-sleep -s 2
}

#interactive mode abnormal output prevention
if( !($cmdline -match "\./chip\-tool\s") -and !$skipcheck){
    $wshell.SendKeys("{enter}")
    start-sleep -s 1
}
[Clicker]::RightClickAtPoint($x1, $y1)
start-sleep -s 2
$wshell.SendKeys("{enter}")
start-sleep -s 2
#check log complete
if($cmdline -match "interactive\sstart" -or ($cmdline.split(" "))[0] -in $global:matchcmds){
    $wshell.SendKeys("{enter}")
    start-sleep -s 2
}

Start-Sleep -s $check_sec

do{
start-sleep -s 1
#interactive mode abnormal output prevention
if( !($cmdline -match "\./chip\-tool\s") -and !$skipcheck){
    $wshell.SendKeys("{enter}")
    start-sleep -s 1
}
$logfile=(Get-ChildItem $logputty|Sort-Object LastWriteTime|Select-Object -last 1).fullname
start-process notepad $logfile -WindowStyle Minimized
start-sleep -s 3
(get-process notepad).CloseMainWindow()|Out-Null
$checkend=((get-content $logfile)[-1]|Out-String).Trim()
$lastword=$checkend[-1]
}until($lastword -eq ":" -or $lastword -eq "$" -or $lastword -eq "#" -or $checkend -eq "logout" -or $checkend -eq ">>>")

$newlogline=(get-content $logfile).count -2
$alllog=get-content $logfile
$checklog=$alllog[$lastlogline..$newlogline]
set-content C:\Matter_AI\logs\lastlog.log -Value $checklog -Force
if($line1 -ne 0){
    $checklog=$checklog[$line1]
    set-content C:\Matter_AI\logs\lastlog_check.log -Value $checklog -Force
    }
if($checkline1.Length -gt 0){
$checkresult=$checklog -like "*$checkline1*"
if($checkresult -and $checkline2.Length -gt 0){
$checkresult=$checklog -like "*$checkline2*"
}
$checkresult
}

}    
#endregion

#region puttyexit
function puttyexit ([int32]$pidd){
$checkexit=get-process putty -ErrorAction SilentlyContinue
while($checkexit){
    if($pidd -eq 0){
     putty_paste -cmdline "exit"
    start-sleep -s 2
     $checkexit=get-process -name putty -ea SilentlyContinue
    }
    else{
        putty_paste -cmdline "exit" -pidd $pidd
        start-sleep -s 2
        $checkexit=get-process -id $pidd -ea SilentlyContinue
    }

}
start-sleep -s 10
}
#endregion

#region renamelog
function renamelog([string]$testid){
$logfile=(Get-ChildItem C:\Matter_AI\logs\*putty.log|Sort-Object LastWriteTime|Select-Object -last 1).fullname
$newname=$testid+"_"+(Get-ChildItem $logfile).name
rename-item -path $logfile -NewName $newname
}
#endregion

#region copyfiles
function copyfile([string]$filepath){
  $copyfilepath="C:\Matter_AI\logs\copyfiles"
  if(!(test-path $copyfilepath)){
    new-item $copyfilepath -ItemType Directory |Out-Null
  }
$settings=get-content C:\Matter_AI\settings\config_linux.txt
$sship=($settings[0].split(":"))[-1]
$sshusername=($settings[1].split(":"))[-1]
#$sshpasswd=($settings[2].split(":"))[-1]
$sshpath=(Split-Path -Parent (($settings[3].split(":"))[-1])).Replace("\","/")
$sourcefile="$($sshusername)@$($sship):$($sshpath)$($filepath)"
C:\Matter_AI\pscp.exe -r -pwfile C:\Matter_AI\psw.txt "$sourcefile" $copyfilepath
}
#endregion



#region WPFmessage
Function New-WPFMessageBox {

  # For examples for use, see my blog:
  # https://smsagent.wordpress.com/2017/08/24/a-customisable-wpf-messagebox-for-powershell/
  
  # CHANGES
  # 2017-09-11 - Added some required assemblies in the dynamic parameters to avoid errors when run from the PS console host.
  
  # Define Parameters
  [CmdletBinding()]
  Param
  (
      # The popup Content
      [Parameter(Mandatory=$True,Position=0)]
      [Object]$Content,

      # The window title
      [Parameter(Mandatory=$false,Position=1)]
      [string]$Title,

      # The buttons to add
      [Parameter(Mandatory=$false,Position=2)]
      [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
      [array]$ButtonType = 'OK',

      # The buttons to add
      [Parameter(Mandatory=$false,Position=3)]
      [array]$CustomButtons,

      # Content font size
      [Parameter(Mandatory=$false,Position=4)]
      [int]$ContentFontSize = 14,

      # Title font size
      [Parameter(Mandatory=$false,Position=5)]
      [int]$TitleFontSize = 14,

      # BorderThickness
      [Parameter(Mandatory=$false,Position=6)]
      [int]$BorderThickness = 1,

      # CornerRadius
      [Parameter(Mandatory=$false,Position=7)]
      [int]$CornerRadius = 8,

      # ShadowDepth
      [Parameter(Mandatory=$false,Position=8)]
      [int]$ShadowDepth = 3,

      # BlurRadius
      [Parameter(Mandatory=$false,Position=9)]
      [int]$BlurRadius = 20,

      # WindowHost
      [Parameter(Mandatory=$false,Position=10)]
      [object]$WindowHost,

      # Timeout in seconds,
      [Parameter(Mandatory=$false,Position=11)]
      [int]$Timeout,

      # Code for Window Loaded event,
      [Parameter(Mandatory=$false,Position=12)]
      [scriptblock]$OnLoaded,

      # Code for Window Closed event,
      [Parameter(Mandatory=$false,Position=13)]
      [scriptblock]$OnClosed

  )

  # Dynamically Populated parameters
  DynamicParam {
      
      # Add assemblies for use in PS Console 
      Add-Type -AssemblyName System.Drawing, PresentationCore
      
      # ContentBackground
      $ContentBackground = 'ContentBackground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute) 
      $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select-Object -ExpandProperty Name 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.ContentBackground = "White"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)
      

      # FontFamily
      $FontFamily = 'FontFamily'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute)  
      $arrSet = [System.Drawing.FontFamily]::Families.Name | Select-Object -Skip 1 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
      $AttributeCollection.Add($ValidateSetAttribute)
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
      $PSBoundParameters.FontFamily = "Segoe UI"

      # TitleFontWeight
      $TitleFontWeight = 'TitleFontWeight'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute) 
      $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select-Object -ExpandProperty Name 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.TitleFontWeight = "Normal"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

      # ContentFontWeight
      $ContentFontWeight = 'ContentFontWeight'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute) 
      $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select-Object -ExpandProperty Name 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.ContentFontWeight = "Normal"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)
      

      # ContentTextForeground
      $ContentTextForeground = 'ContentTextForeground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute) 
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select-Object -ExpandProperty Name 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.ContentTextForeground = "Black"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

      # TitleTextForeground
      $TitleTextForeground = 'TitleTextForeground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute) 
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select-Object -ExpandProperty Name 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.TitleTextForeground = "Black"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

      # BorderBrush
      $BorderBrush = 'BorderBrush'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute) 
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select-Object -ExpandProperty Name 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.BorderBrush = "Black"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


      # TitleBackground
      $TitleBackground = 'TitleBackground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute) 
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select-Object -ExpandProperty Name 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.TitleBackground = "White"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

      # ButtonTextForeground
      $ButtonTextForeground = 'ButtonTextForeground'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute) 
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select-Object -ExpandProperty Name 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.ButtonTextForeground = "Black"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

     
      # ButtonBorderThickness
      $ButtonBorderThickness = 'ButtonBorderThickness'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      $AttributeCollection.Add($ParameterAttribute) 
      $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select-Object -ExpandProperty Name 
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $PSBoundParameters.ButtonBorderThickness = "4"
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonBorderThickness, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($ButtonBorderThickness, $RuntimeParameter)

      # Sound
      $Sound = 'Sound'
      $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
      $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
      $ParameterAttribute.Mandatory = $False
      #$ParameterAttribute.Position = 14
      $AttributeCollection.Add($ParameterAttribute) 
      $arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select-Object -ExpandProperty Name).Replace('.wav','')
      $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
      $AttributeCollection.Add($ValidateSetAttribute)
      $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
      $RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)

      return $RuntimeParameterDictionary
  }

  Begin {
      Add-Type -AssemblyName PresentationFramework
  }
  
  Process {

# Define the XAML markup
[XML]$Xaml = @"
<Window 
      xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1">
  <Window.Resources>
      <Style TargetType="{x:Type Button}">
          <Setter Property="Template">
              <Setter.Value>
                  <ControlTemplate TargetType="Button">
                      <Border>
                          <Grid Background="LightGreen">
                              <ContentPresenter />
                          </Grid>
                      </Border>
                  </ControlTemplate>
              </Setter.Value>
          </Setter>
      </Style>
  </Window.Resources>
  <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
      <Border.Effect>
          <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
      </Border.Effect>
      <Border.Triggers>
          <EventTrigger RoutedEvent="Window.Loaded">
              <BeginStoryboard>
                  <Storyboard>
                      <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                      <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                  </Storyboard>
              </BeginStoryboard>
          </EventTrigger>
      </Border.Triggers>
      <Grid >
          <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
          <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
              <Grid.OpacityMask>
                  <VisualBrush Visual="{Binding ElementName=Mask}"/>
              </Grid.OpacityMask>
              <StackPanel Name="StackPanel" >                   
                  <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                  <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                  </DockPanel>
                  <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                  </DockPanel>
              </StackPanel>
          </Grid>
      </Grid>
  </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="$($PSBoundParameters.ButtonBorderThickness)" BorderBrush="Red" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" IsDefault="True" Cursor="Hand" Focusable="True"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

  # Load the window from XAML
  $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

  # Custom function to add a button
  Function Add-Button {
      Param($Content)
      $Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
      $ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
      $ButtonText.Text = "$Content"
      $Button.Content = $ButtonText
      $Button.Add_MouseEnter({
          $This.Content.FontSize = "17"
      })
      $Button.Add_MouseLeave({
          $This.Content.FontSize = "16"
      })
      $Button.Add_Click({
          New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -Scope Script -Force
          $Window.Close()
      })
      $Window.FindName('ButtonHost').AddChild($Button)
     
  }

  # Add buttons
  If ($ButtonType -eq "OK")
  {
      Add-Button -Content "OK"
  }

  If ($ButtonType -eq "OK-Cancel")
  {
      Add-Button -Content "OK"
      Add-Button -Content "Cancel"
  }

  If ($ButtonType -eq "Abort-Retry-Ignore")
  {
      Add-Button -Content "Abort"
      Add-Button -Content "Retry"
      Add-Button -Content "Ignore"
  }

  If ($ButtonType -eq "Yes-No-Cancel")
  {
      Add-Button -Content "Yes"
      Add-Button -Content "No"
      Add-Button -Content "Cancel"
  }

  If ($ButtonType -eq "Yes-No")
  {
      Add-Button -Content "Yes"
      Add-Button -Content "No"
  }

  If ($ButtonType -eq "Retry-Cancel")
  {
      Add-Button -Content "Retry"
      Add-Button -Content "Cancel"
  }

  If ($ButtonType -eq "Cancel-TryAgain-Continue")
  {
      Add-Button -Content "Cancel"
      Add-Button -Content "TryAgain"
      Add-Button -Content "Continue"
  }

  If ($ButtonType -eq "None" -and $CustomButtons)
  {
      Foreach ($CustomButton in $CustomButtons)
      {
          Add-Button -Content "$CustomButton"
      }
  }

  # Remove the title bar if no title is provided
  If ($Title -eq "")
  {
      $TitleBar = $Window.FindName('TitleBar')
      $Window.FindName('StackPanel').Children.Remove($TitleBar)
  }

  # Add the Content
  If ($Content -is [String])
  {
      # Replace double quotes with single to avoid quote issues in strings
      If ($Content -match '"')
      {
          $Content = $Content.Replace('"',"'")
      }
      
      # Use a text box for a string value...
      $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
      $Window.FindName('ContentHost').AddChild($ContentTextBox)
  }
  Else
  {
      # ...or add a WPF element as a child
      Try
      {
          $Window.FindName('ContentHost').AddChild($Content) 
      }
      Catch
      {
          $_
      }        
  }

  # Enable window to move when dragged
  $Window.FindName('Grid').Add_MouseLeftButtonDown({
      $Window.DragMove()
  })

  $window.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Space -or $e.Key -eq [System.Windows.Input.Key]::Enter ) {
        # Trigger the button click when "Space" is pressed
        #$button.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
        $Window.Close()
    }
})
  # Activate the window on loading
  If ($OnLoaded)
  {
      $Window.Add_Loaded({
          $This.Activate()
          Invoke-Command $OnLoaded
      })
  }
  Else
  {
      $Window.Add_Loaded({
          $This.Activate()
      })
  }
  

  # Stop the dispatcher timer if exists
  If ($OnClosed)
  {
      $Window.Add_Closed({
          If ($DispatcherTimer)
          {
              $DispatcherTimer.Stop()
          }
          Invoke-Command $OnClosed
      })
  }
  Else
  {
      $Window.Add_Closed({
          If ($DispatcherTimer)
          {
              $DispatcherTimer.Stop()
          }
      })
  }
  

  # If a window host is provided assign it as the owner
  If ($WindowHost)
  {
      $Window.Owner = $WindowHost
      $Window.WindowStartupLocation = "CenterOwner"
  }

  # If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
  If ($Timeout)
  {
      $Stopwatch = New-object System.Diagnostics.Stopwatch
      $TimerCode = {
          If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
          {
              $Stopwatch.Stop()
              $Window.Close()
          }
      }
      $DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
      $DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
      $DispatcherTimer.Add_Tick($TimerCode)
      $Stopwatch.Start()
      $DispatcherTimer.Start()
  }

  # Play a sound
  If ($($PSBoundParameters.Sound))
  {
      $SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
      $SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
      $SoundPlayer.Add_LoadCompleted({
          $This.Play()
          $This.Dispose()
      })
      $SoundPlayer.LoadAsync()
  }

  # Display the window
  $null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

  }
}

#endregion

#region manual GUI
function selection_manual($data, $column1, $column2) {
    # Load Windows Forms
    Add-Type -AssemblyName System.Windows.Forms
    
    # Initialize Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Matter Manul TC Selection"
    $form.Size = New-Object System.Drawing.Size(650,400)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    
    # Create Label for Column1 (above the first ListBox)
    $labelColumn1 = New-Object System.Windows.Forms.Label
    $labelColumn1.Text = $column1
    $labelColumn1.Location = New-Object System.Drawing.Point(10, 10)
    $labelColumn1.Size = New-Object System.Drawing.Size(150, 20)
    
    # Create Label for Column2 (above the second ListBox)
    $labelColumn2 = New-Object System.Windows.Forms.Label
    $labelColumn2.Text = $column2
    $labelColumn2.Location = New-Object System.Drawing.Point(200, 10)
    $labelColumn2.Size = New-Object System.Drawing.Size(150, 20)
    
    # Create Label for Column3 (above the second ListBox)
    $labelColumn3 = New-Object System.Windows.Forms.Label
    $labelColumn3.Text = "All TC"
    $labelColumn3.Location = New-Object System.Drawing.Point(400, 10)
    $labelColumn3.Size = New-Object System.Drawing.Size(150, 20)
    
    # Create ListBox for Unique Column1 values
    $uniqueListBox = New-Object System.Windows.Forms.ListBox
    $uniqueListBox.Location = New-Object System.Drawing.Point(10,40)
    $uniqueListBox.Size = New-Object System.Drawing.Size(150,300)
    $uniqueListBox.SelectionMode = 'MultiExtended'
    
    
    # Create ListBox for Column2 values filtered by Column1
    $filteredListBox = New-Object System.Windows.Forms.ListBox
    $filteredListBox.Location = New-Object System.Drawing.Point(200,40)
    $filteredListBox.Size = New-Object System.Drawing.Size(150,300)
    $filteredListBox.SelectionMode = 'MultiExtended'
    
    # Create ListBox to display the selected items from Column2
    $selectedListBox = New-Object System.Windows.Forms.ListBox
    $selectedListBox.Location = New-Object System.Drawing.Point(400,40)
    $selectedListBox.Size = New-Object System.Drawing.Size(150,300)
    
    # Button to move selected items from Column2 ListBox to Selected ListBox
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "+"
    $button.Size =  New-Object System.Drawing.Size(40,20)
    $button.Location = New-Object System.Drawing.Point(355,120)

    # Button to remove selected items from  final list ListBox to Column2 ListBox
    $rmbutton = New-Object System.Windows.Forms.Button
    $rmbutton.Text = "-"
    $rmbutton.Size =  New-Object System.Drawing.Size(40,20)
    $rmbutton.Location = New-Object System.Drawing.Point(355,150)
    
    # Create "Done" button to save the final list and close the form
    $doneButton = New-Object System.Windows.Forms.Button
    $doneButton.Text = "Done"
    $doneButton.Size =  New-Object System.Drawing.Size(40,20)
    $doneButton.Location = New-Object System.Drawing.Point(560, 120)
    
    
    # Populate Unique ListBox with unique Column1 values
    $uniqueValues = $data | ForEach-Object { $_.$column1 } | Sort-Object -Unique
    $uniqueListBox.Items.AddRange($uniqueValues)
    
    # Event: Update Column2 ListBox based on Column1 selection
    $uniqueListBox.add_SelectedIndexChanged({
        $filteredListBox.Items.Clear()
        $selectedValues = $uniqueListBox.SelectedItems
        if ($selectedValues.Count -gt 0) {
            $filteredData = $data | Where-Object { $selectedValues -contains $_.$column1 }
            $filteredColumn2 = $filteredData | ForEach-Object { $_.$column2 } | Sort-Object -Unique
            $filteredListBox.Items.AddRange($filteredColumn2)
        }
    })
    
    # Event: Add selected items from Column2 to the third ListBox
    $button.add_Click({
        $selectedItems = $filteredListBox.SelectedItems
        foreach ($item in $selectedItems) {
            if (-not $selectedListBox.Items.Contains($item)) {
                $selectedListBox.Items.Add($item)
            }
        }
    })
    
  # Event: remove selected items from the third ListBox to Column2 
    $rmbutton.add_Click({
        $rmselectedItems = @($selectedListBox.SelectedItems)
        $rmselectedItems
        foreach ($item in $rmselectedItems) {
            $selectedListBox.Items.Remove($item)            
        }
    })
    
    $doneButton.add_Click({
        $global:sels = @()
        foreach ($item in $selectedListBox.Items) {
            $global:sels += $item 
        }
        $form.Close()  # Close the form after saving the final list
    })
    
    # Add controls to the form
    $form.Controls.Add($labelColumn1)
    $form.Controls.Add($labelColumn2)
    $form.Controls.Add($labelColumn3)
    $form.Controls.Add($uniqueListBox)
    $form.Controls.Add($filteredListBox)
    $form.Controls.Add($selectedListBox)
    $form.Controls.Add($button)
    $form.Controls.Add($rmbutton)
    $form.Controls.Add($doneButton)
    
    # Show the form
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
    if($global:sels.count -eq 0){

        [System.Windows.Forms.MessageBox]::Show("Please select testcases","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        return 0
        
    }
    $global:sels
    }
    <#
    $data = @(
        @{Column1='A'; Column2='Apple'},
        @{Column1='B'; Column2='Banana'},
        @{Column1='A'; Column2='Avocado'},
        @{Column1='C'; Column2='Cherry'},
        @{Column1='B'; Column2='Blueberry'},
        @{Column1='A'; Column2='Apricot'},
        @{Column1='C'; Column2='Clementine'}
    )
    #>
   #endregion 

#region putty starting functions

function puttystart ([string]$puttyname) {
     
   $puttypid=($global:puttyset|Where-Object{$_.name -eq $puttyname}|Select-Object -last 1).puttypid
   if($puttypid){
    $checkpid=get-process -id $puttypid -ErrorAction SilentlyContinue
   }
   if(!$puttypid -or !$checkpid){
    $settings=get-content C:\Matter_AI\settings\config_linux.txt
    $sship=($settings[0].split(":"))[-1]
    $regfile="C:\Matter_AI\puttyreg.reg"
      if($sship -ne "192.168.2.201"){
        (get-content "C:\Matter_AI\puttyreg.reg").replace("192.168.2.201",$sship)|Set-Content "C:\Matter_AI\puttyreg1.reg"
        $regfile="C:\Matter_AI\puttyreg1.reg"
       }
    
    $sesname="matter"
    if($puttyname.length -gt 0){
      $sesname="matter_$($puttyname)"
      $newputtyreg="C:\Matter_AI\puttyreg_$($puttyname).reg"
      $sessionname="Sessions\matter_$($puttyname)"
        Copy-Item $regfile $newputtyreg -Force
        $puttylogpath="C:\\Matter_AI\\logs\\&Y&M&D&T_&H_putty.log"
        $puttylogpathnew="C:\\Matter_AI\\logs\\&Y&M&D&T_&H_putty_$($puttyname).log"
        ((get-content $newputtyreg).replace("Sessions\matter",$sessionname)).replace($puttylogpath,$puttylogpathnew)|Set-Content $newputtyreg
        $regfile=$newputtyreg
    }
    start-process reg -ArgumentList "import $regfile"

    $beforepid="na"
    if(get-process -name putty -ErrorAction SilentlyContinue){
    $beforepid=(get-process -name putty).id
    }
    $putty="C:\Matter_AI\putty.exe"
    start-sleep -s 2

    start-process $putty -ArgumentList "-load $sesname" -WindowStyle Maximized
    $afterpid=(get-process -name putty|Where-Object{$_.id -notin $beforepid}).id
    if($puttypid){
        ($global:puttyset|Where-Object{$_.name -eq $puttyname}).puttypid=$afterpid
    }else{
      $global:puttyset+=New-Object -TypeName PSObject -Property @{
          name=$puttyname
          puttypid=$afterpid
        }
     }
    $global:puttyset|Format-Table
     $puttypid=($global:puttyset|Where-Object{$_.name -eq $puttyname}|Select-Object -last 1).puttypid
    
       $settings=get-content C:\Matter_AI\settings\config_linux.txt
       $pskey=($settings[2].split(":"))[-1]
       #$sshpath=($settings[3].split(":"))[-1]
       $fname=(Get-ChildItem $global:excelfile).name
       $sshpath=(import-csv C:\Matter_AI\settings\filesettings.csv|Where-Object{$_.filename -eq $fname}).python_path
       $sshpathmanual=(import-csv C:\Matter_AI\settings\filesettings.csv|Where-Object{$_.filename -eq $fname}).manual_path
       $wshell = New-Object -ComObject WScript.Shell
       $wshell.AppActivate($puttypid)
       start-sleep -s 5
       $wshell.SendKeys("raspberrypi")
       start-sleep -s 2
       $wshell.SendKeys("{enter}")
       start-sleep -s 2

       
       if($global:testtype -eq 1){
        $global:failmsg=$null
        putty_paste -cmdline "sudo -s" -skipcheck
        putty_paste -cmdline $pskey  -skipcheck
        putty_paste -cmdline "docker ps -a" -skipcheck
        $idlogin=get-content "C:\Matter_AI\logs\lastlog.log" 
        $checkmatch=$idlogin -match "\/bin\/bash"
        if($checkmatch){
            $ctnid= (($idlogin -match "\/bin\/bash").split(" "))[0]
        }
        else{
            #puttyexit
            $global:failmsg="No found /bin/bash path"
            Write-Output $global:failmsg
        }
       #putty_paste -cmdline "docker start $ctnid"
       #putty_paste -cmdline "docker exec -it $ctnid /bin/bash"
       #putty_paste -cmdline "cd $sshpath"
        if(!$global:failmsg){
        putty_paste -cmdline "docker start $ctnid" -skipcheck
        putty_paste -cmdline "docker exec -it $ctnid /bin/bash" -skipcheck
        putty_paste -cmdline "cd $sshpath" -skipcheck
        }
       }
       if($global:testtype -eq 2){
       putty_paste -cmdline "sudo -s" -puttyname $puttyname -skipcheck
       putty_paste -cmdline $pskey -puttyname $puttyname -skipcheck
       putty_paste -cmdline "cd $($sshpathmanual)" -puttyname $puttyname -skipcheck
       }
    }
}

#endregion


function webdownload ([string]$goo_link,[string]$gid,[string]$sv_range,[string]$savepath,[string]$errormessage){
    
    Remove-Item "$ENV:UserProfile\downloads\*.csv" -force
    $link_save=$goo_link+"export?format=csv&gid=$($gid)&range=$($sv_range)"
    #$link_save
    $starttime=get-date
    Start-Process msedge $link_save
    
    do{
    Start-Sleep -s 2
    $lsnewc=(Get-ChildItem -path "$ENV:UserProfile\Downloads\*.csv" -file).count
    $timepassed=(new-timespan -start $starttime -end (get-date)).TotalSeconds
    }until($lsnewc -eq 1 -or $timepassed -gt 30)
    
    if($lsnewc){
    $downloadname= (Get-ChildItem -path "$ENV:UserProfile\Downloads\*.csv").FullName
    if(!(test-path $savepath)){
        New-Item -ItemType Directory $savepath -Force |Out-Null
     }
    copy-item $downloadname -Destination $savepath -Force  
    Remove-Item "$ENV:UserProfile\downloads\*.csv" -force
    return "Download ok"
    }
    else{
        $region = (Get-Culture).Name
        $ipAddress = (Invoke-RestMethod -Uri "http://ipinfo.io/json").ip
        
      $paramHash = @{
      To="shuningyu17120@allion.com.tw"
      from = 'Notioce <npl_siri@allion.com.tw>'
      BodyAsHtml = $True
      Subject = $errormessage
      Body = "Region: $region <br>IP: $ipAddress <br>$env:COMPUTERNAME/$env:UserName<br> Fail to download $goo_link"
     }
     
     Send-MailMessage @paramHash -Encoding utf8 -SmtpServer zimbra.allion.com.tw 
     return "Fail Download"
    }
    start-sleep -s 5
    (get-process -name "msedge" -ea SilentlyContinue).CloseMainWindow()|Out-Null   
  }
  

  function getparameter([string]$getlastkey){       
        $lastlogcontent=get-content -path C:\Matter_AI\logs\lastlog.log|Select-Object -skip 2
        $getlastkey=$getlastkey.replace("[","\[").replace("]","\]")
        $matchvalue= ([regex]::Match(($lastlogcontent -match $getlastkey), "$getlastkey(.*)").Groups[1].value).tostring().trim()
        $matchvalue=(($matchvalue.replace("[","")).replace("]","")).replace(",","")
        #$matchvalue= (($lastlogcontent|Select-String -Pattern "($getlastkey).*" -AllMatches |  ForEach-Object {$_.matches.value}).split($getlastkey))[-1].trim()
        $global:varhash+=@([PSCustomObject]@{           
         para_name = $paraname
         setvalue = $matchvalue
        })      
        $global:varhash
  }

        function selectcom{# Load the required .NET assembly for Windows Forms
            Add-Type -AssemblyName System.Windows.Forms
            
            # Create a new form
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Select COM Port"
            $form.Size = New-Object System.Drawing.Size(300, 150)
            $form.StartPosition = "CenterScreen"
            
            # Create a ComboBox for COM ports selection
            $comboBox = New-Object System.Windows.Forms.ComboBox
            $comboBox.Location = New-Object System.Drawing.Point(50, 20)
            $comboBox.Size = New-Object System.Drawing.Size(180, 20)
            
            # Populate the ComboBox with COM1 to COM256 without writing output
            $portnames=[System.IO.Ports.SerialPort]::getportnames()
            foreach ($portname in $portnames) {
                $comboBox.Items.Add($portname) | Out-Null
            }
            
            # Create a Button to confirm selection
            $button = New-Object System.Windows.Forms.Button
            $button.Location = New-Object System.Drawing.Point(100, 60)
            $button.Size = New-Object System.Drawing.Size(80, 30)
            $button.Text = "Select"
            
            # Add the ComboBox and Button to the form
            $form.Controls.Add($comboBox)
            $form.Controls.Add($button)
            
            # Add an event handler for the Button click event
            $button.Add_Click({
                if ($comboBox.SelectedItem) {
                    $global:comport = $comboBox.SelectedItem
                    $form.Close()
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Please select a COM port!", "Warning")
                }
            })
            
            # Show the form
            $form.Topmost = $true
            $form.ShowDialog()
       
   }
  function dutcontrol ([string]$mode){
    if ($mode.length -gt 0){
        $portid=((get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "serialport"}) -split ":")[1]
        if(!($portid -match "\d+")){
            selectcom
            if($global:comport -match "\d+"){
                $newcontent=get-content C:\Matter_AI\settings\config_linux.txt|ForEach-Object{
                    if($_ -match "serialport"){
                        $_="serialport:$global:comport"
                    }
                    $_
                }
                $newcontent|set-content C:\Matter_AI\settings\config_linux.txt
            }
            $portid=$global:comport
        }
       
        $port = New-Object System.IO.Ports.SerialPort
        $port.PortName = $portid
        $port.BaudRate = "9600"
        $port.Parity = "None"
        $port.DataBits = 8
        $port.StopBits = 1
        $port.ReadTimeout = 9000 # 9 seconds
        $port.DtrEnable = "true"

        $modes=@("on","off","up","down","testcom")
        $sendings=@("o","f","b","c","")
        $sending=$sendings[$modes.indexof($mode.ToLower())]
        $waittime=((get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "wait" -and $_ -match $mode}) -split ":")[1]

        do{
        $port.PortName = $portid    
        $port.open() #opens serial connection
        if($? -eq 0){
            add-content C:\Matter_AI\logs\testing.log -value "fail to open the serial port"
            do{
                selectcom
             }until($global:comport -match "\d+")
             $portid=$global:comport
             $newcontent=get-content C:\Matter_AI\settings\config_linux.txt|ForEach-Object{
                 if($_ -match "serialport"){
                     $_="serialport:$global:comport"
                 }
                 $_
             }
             $newcontent|set-content C:\Matter_AI\settings\config_linux.txt        
          }
          else{
            $Global:seialport="ok"    
          }
        }while(!$Global:seialport)

          Start-Sleep 2 # wait 2 seconds until Arduino is ready
             if ($sending.length -gt 0){
                $port.Write($sending) #writes your content to the serial connection
                if($? -eq 0){
                add-content C:\Matter_AI\logs\testing.log -value "fail to send signal to serial port"
                }
                else{
                    add-content C:\Matter_AI\logs\testing.log -value "$mode - action done $((get-date|Out-String).trim())"
                    start-sleep -Milliseconds $waittime
                }
                }
         
          $port.Close() #closes serial connection
          if($? -eq 0){
            add-content C:\Matter_AI\logs\testing.log -value "fail to close the serial port"
           }
           Start-Sleep 2

        }
    
     }


 function dutpower([int32]$mode){    
    if($mode -eq 1){
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
       if($mode -eq 2){
        $cycletime= ((get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "cycle" -and $_ -match "onoff"}) -split ":")[1]
        foreach($i in 1..$cycletime){ 
         dutcontrol -mode off
         dutcontrol -mode on
        }
        }
        if($mode -eq 3){
          $cycletime= ((get-content C:\Matter_AI\settings\config_linux.txt|Where-Object{$_ -match "cycle" -and $_ -match "downup"}) -split ":")[1]
          foreach($i in 1..$cycletime){
            dutcontrol -mode down
            dutcontrol -mode up   
          }
        }
    }

     function selguis ( [string[]]$Inputdata,[string]$instruction,[string]$errmessage) {

        Add-Type -AssemblyName System.Windows.Forms
         $newFont = New-Object System.Drawing.Font("Microsoft Sans Serif", 12)
        # Create the form
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "List Box Transfer UI"
        $form.Size = New-Object System.Drawing.Size(800, 600)  # Enlarged form size
        $form.StartPosition = "CenterScreen"
        
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(40,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
        $label.Text = $instruction
        $label.Font = $newFont
        # Left ListBox
        $leftListBox = New-Object System.Windows.Forms.ListBox
        $leftListBox.Location = New-Object System.Drawing.Point(40, 50)
        $leftListBox.Size = New-Object System.Drawing.Size(260, 400)  # Enlarged
        $leftListBox.SelectionMode = "MultiExtended" 
        $leftListBox.Items.AddRange($Inputdata)
        
        
        # Right ListBox
        $rightListBox = New-Object System.Windows.Forms.ListBox
        $rightListBox.Location = New-Object System.Drawing.Point(480, 50)
        $rightListBox.Size = New-Object System.Drawing.Size(260, 400)  # Enlarged
        $rightListBox.SelectionMode = "MultiExtended" 
        
        # Add Button
        $addButton = New-Object System.Windows.Forms.Button
        $addButton.Text = "+"
        $addButton.Location = New-Object System.Drawing.Point(370, 160)  # Adjusted for larger form
        $addButton.Size = New-Object System.Drawing.Size(60, 30)  # Enlarged button
        
        # Remove Button
        $removeButton = New-Object System.Windows.Forms.Button
        $removeButton.Text = "-"
        $removeButton.Location = New-Object System.Drawing.Point(370, 220)  # Adjusted for larger form
        $removeButton.Size = New-Object System.Drawing.Size(60, 30)  # Enlarged button
        
        # Done Button
        $doneButton = New-Object System.Windows.Forms.Button
        $doneButton.Text = "Done"
        $doneButton.Location = New-Object System.Drawing.Point(360, 450)  # Adjusted for larger form
        $doneButton.Size = New-Object System.Drawing.Size(80, 30)  # Enlarged button
        
        # Add Button Click Event
        $addButton.Add_Click({
            if ($leftListBox.SelectedItems.Count -gt 0) {
                foreach ($item in $leftListBox.SelectedItems) {
                    $rightListBox.Items.Add($item)
                }
                # Remove selected items from the left ListBox
                foreach ($item in @($leftListBox.SelectedItems)) {
                    $leftListBox.Items.Remove($item)
                }
            }
        })
        
        
        # Remove Button Click Event
        $removeButton.Add_Click({
            if ($rightListBox.SelectedItems.Count -gt 0) {
                foreach ($item in $rightListBox.SelectedItems) {
                    $leftListBox.Items.Add($item)
                }
                # Remove selected items from the left ListBox
                foreach ($item in @($rightListBox.SelectedItems)) {
                    $rightListBox.Items.Remove($item)
                }
            }
        })
        
        # Done Button Click Event
        $doneButton.Add_Click({
            $global:selss=$rightListBox.Items
            $form.Close()
        })
        
        # Add controls to the form
        
        $form.Controls.Add($label)
        $form.Controls.Add($leftListBox)
        $form.Controls.Add($rightListBox)
        $form.Controls.Add($addButton)
        $form.Controls.Add($removeButton)
        $form.Controls.Add($doneButton)
        
        # Show the form
        $global:selss=@()
        [void]$form.ShowDialog()
        
        if($global:selss.count -eq 0){
        
            [System.Windows.Forms.MessageBox]::Show($errmessage,"Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            return 0
            
        }
        
        $global:selss
        
        }

        
     function selgui ( [string[]]$Inputdata,[string]$instruction,[string]$errmessage) {

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        
        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Data Entry Form'
        $form.Size = New-Object System.Drawing.Size(300,250)
        $form.StartPosition = 'CenterScreen'
        
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(75,150)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = 'OK'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $OKButton
        $form.Controls.Add($OKButton)
        
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(150,150)
        $CancelButton.Size = New-Object System.Drawing.Size(75,23)
        $CancelButton.Text = 'Cancel'
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $CancelButton
        $form.Controls.Add($CancelButton)
        
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
        $label.Text = $instruction
        $form.Controls.Add($label)
        
        $listBox = New-Object System.Windows.Forms.Listbox
        $listBox.Location = New-Object System.Drawing.Point(10,40)
        $listBox.Size = New-Object System.Drawing.Size(260,20)
        
        $listBox.SelectionMode = 'MultiExtended'
        
        #select testcase
        $sels=@()
        $pylines=@()
        
        
        foreach ($select in $Inputdata){
        [void] $listBox.Items.Add($select)
        }
        
        $listBox.Height = 100
        $form.Controls.Add($listBox)
        $form.Topmost = $true
        
        $result = $form.ShowDialog()
        $global:sels=@()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $x = $listBox.SelectedItems
            if($x){
                $global:sels+=@($x.trim())
            }
            
        }
        
        if($global:sels.count -eq 0){
        
            [System.Windows.Forms.MessageBox]::Show($errmessage,"Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            return 0
            
        }
        
        
        $global:sels

     }

     function webuiSelections{

        Add-Type -AssemblyName System.Windows.Forms
        
        # Sample file list
        $list = (Get-ChildItem -path C:\Matter_AI\logs\_auto\ -Directory).Name
        if($list.count -eq 0){
          $global:webuiselects="Creating a new project"
          return
        }
        # Create form
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Project Selection"
        $form.Size = New-Object System.Drawing.Size(400, 300)
        $form.StartPosition = "CenterScreen"
        
        # Create "Create a New Project" RadioButton
        $rbNewProject = New-Object System.Windows.Forms.RadioButton
        $rbNewProject.Text = "Create a New Project"
        $rbNewProject.Location = New-Object System.Drawing.Point(20, 20)
        $rbNewProject.Size = New-Object System.Drawing.Size(200, 20)
        
        # Create "Use Existing Projects" RadioButton
        $rbExistingProjects = New-Object System.Windows.Forms.RadioButton
        $rbExistingProjects.Text = "Use Existing Projects"
        $rbExistingProjects.Location = New-Object System.Drawing.Point(20, 50)
        $rbExistingProjects.Size = New-Object System.Drawing.Size(200, 20)
        
        # Create ListBox for existing project files
        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point(40, 80)
        $listBox.Size = New-Object System.Drawing.Size(300, 80)
        $listBox.Items.AddRange($list)
        $listBox.Enabled = $false  # Initially disabled
        
        # Create "Update JSON" Checkbox
        $cbUpdateJson = New-Object System.Windows.Forms.CheckBox
        $cbUpdateJson.Text = "Update JSON"
        $cbUpdateJson.Location = New-Object System.Drawing.Point(40, 170)
        $cbUpdateJson.Size = New-Object System.Drawing.Size(200, 20)
        $cbUpdateJson.Enabled = $false  # Initially disabled
        
        # Create "Update XML" Checkbox
        $cbUpdateXml = New-Object System.Windows.Forms.CheckBox
        $cbUpdateXml.Text = "Update XML"
        $cbUpdateXml.Location = New-Object System.Drawing.Point(40, 200)
        $cbUpdateXml.Size = New-Object System.Drawing.Size(200, 20)
        $cbUpdateXml.Enabled = $false  # Initially disabled
        
        # Create OK button
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "OK"
        $btnOK.Location = New-Object System.Drawing.Point(150, 230)
        $btnOK.Size = New-Object System.Drawing.Size(100, 30)
        $btnOK.Add_Click({
            # Define button action here
            if ($rbNewProject.Checked) {
                #[System.Windows.Forms.MessageBox]::Show("Creating a new project...")
                $global:webuiselects="1"
            } elseif ($rbExistingProjects.Checked) {
                if ($listBox.SelectedItem -ne $null) {
                    $selectedUpdates = @()
                    if ($cbUpdateJson.Checked) { $selectedUpdates += "JSON" }
                    if ($cbUpdateXml.Checked) { $selectedUpdates += "XML" }
        
                    $global:webuiselects="Updating $($selectedUpdates -join ', ') for project: $($listBox.SelectedItem)"
                } else {
                    #[System.Windows.Forms.MessageBox]::Show("Please select an existing project.")
                    return  # Exit click event without closing the form
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please select an option.")
                return  # Exit click event without closing the form
            }
            
            # Close the form after successful action
            $form.Close()
          })
        
        # Event handlers to enable/disable controls based on the selected option
        $rbNewProject.Add_CheckedChanged({
            if ($rbNewProject.Checked) {
                $listBox.Enabled = $false
                $cbUpdateJson.Enabled = $false
                $cbUpdateXml.Enabled = $false
            }
        })
        
        $rbExistingProjects.Add_CheckedChanged({
            if ($rbExistingProjects.Checked) {
                $listBox.Enabled = $true
                $cbUpdateJson.Enabled = $true
                $cbUpdateXml.Enabled = $true
            }
        })
        
        # Add controls to form
        $form.Controls.Add($rbNewProject)
        $form.Controls.Add($rbExistingProjects)
        $form.Controls.Add($listBox)
        $form.Controls.Add($cbUpdateJson)
        $form.Controls.Add($cbUpdateXml)
        $form.Controls.Add($btnOK)
        
        # Show form
        $form.Add_Shown({$form.Activate()})
        [void] $form.ShowDialog()
        }
        
       