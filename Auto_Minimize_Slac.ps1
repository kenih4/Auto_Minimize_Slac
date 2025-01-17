#   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process



# MEMO  
# Arg1: Miniute
# Arg2: Target Window title Name
# .\Auto_Minimize_Slac.ps1 0.5 general
# Get-Process | Where-Object {$_.MainWindowTitle -ne ""} | Select-Object MainWindowTitle
#
# IExpressでexe化するとき、 残念ながら引数は埋め込み　なんのために引数にしたんだ！
# PowerShell.exe -ExecutionPolicy Bypass -File C:\Users\kenichi\Dropbox\gitdir\Auto_Minimize_Slac\Auto_Minimize_Slac.ps1 1.0 general

#  for Debug output 
#  $DebugPreference = 'Continue'
#  or
#  Auto_Minimize_Slac.ps1 -Debug 1.0 general
#

# Parameters
Param (
   $Arg1,$Arg2,
   [Parameter(HelpMessage="Timeout minutes until minimuze")][alias("Timeout","t")][ValidateRange(1,1440)][Int]$TimeoutMin = 1,
   [alias("Force","f")][Switch]$ForceFlag
)

$ActionInt = 0
if ($ForceFlag) {
   $ActionInt = $ActionInt + 4
}
  
# Modules
Add-Type @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
  
namespace PInvoke.Win32 {
  
   public static class UserInput {
  
      [DllImport("user32.dll", SetLastError=false)]
      private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
  
      [StructLayout(LayoutKind.Sequential)]
      private struct LASTINPUTINFO {
         public uint cbSize;
         public int dwTime;
      }
  
      public static DateTime LastInput {
         get {
            DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
            DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
            return lastInput;
         }
      }
  
      public static TimeSpan IdleTime {
         get {
            return DateTime.UtcNow.Subtract(LastInput);
         }
      }
  
      public static int LastInputTicks {
         get {
            LASTINPUTINFO lii = new LASTINPUTINFO();
            lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
            GetLastInputInfo(ref lii);
            return lii.dwTime;
         }
      }
   }
}
'@
  


# Modules  for get current window title
$code = @'
    [DllImport("user32.dll")]
     public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
'@

Add-Type $code -Name Utils -Namespace Win32
$myPid = [IntPtr]::Zero;

# Modules  for Window minimize
Add-Type -AssemblyName UIAutomationClient












function zzz($timeup, [string]$msg = "")
{
#    Clear-Host
#    Write-Host "@zzz    timeup:" + $timeup
#    Write-Host "@zzz    timeup KATA:" + $timeup.GetType().FullName

        while (1) 
        {
#            Write-Progress -Activity "w..." -PercentComplete ([int]((100 * $i))) -Status ("[Now:" + (Get-Date -Format "HH:mm") + "]  <time:" + $timeup + ">")
            $now = Get-Date
            Start-Sleep -Second 3
            if (($now - $timeup) -le 0)
            {
#               Write-Host ("MADA Now:" + (Get-Date -Format "HH:mm") + "]  <time:" + $timeup + ">    " + ($now - $timeup))  -ForegroundColor DarkGray
            }else{
#                Write-Host "Finish"
                break
            }
        }
}










Write-Debug $Arg1
Write-Debug $Arg2

$TargetWindowName = $Arg2
Write-Debug "* * * TargetWindowName: $TargetWindowName" 

$nowH = (Get-Date).Hour
if($nowH -ge 1 -And $nowH -lt 9){
    $str = [string](Get-Date).ToString("yyyy/MM/dd")  + " 09:00:00"
}elseif($nowH -gt 9 -And $nowH -lt 17){
#    $str = [string](Get-Date).ToString("yyyy/MM/dd")  + " 17:00:00"
    $str = [string](Get-Date).ToString("yyyy/MM/dd")  + " 17:00:00"
}elseif($nowH -gt 17 -And $nowH -lt 24){
    $tommorow = (Get-Date).AddDays(1)
    $str = [string]($tommorow).ToString("yyyy/MM/dd")  + " 01:00:00"
}elseif($nowH -eq 0){
   $str = [string](Get-Date).ToString("yyyy/MM/dd")  + " 01:00:00"
}else{
   return
}

Write-Debug "str:  $str" 
$timeup = [DateTime]::ParseExact($str,"yyyy/MM/dd HH:mm:ss", $null);
Write-Debug "timeup:  $timeup" 



# Main Routine
#  zzz $timeup Timeup


#$Timeout = [TimeSpan]::FromMinutes($TimeoutMin)
$Timeout = [TimeSpan]::FromMinutes($Arg1)
do {

#   $NowTotalWindowTItle = Get-Process | Where-Object {$_.MainWindowTitle -ne ""} | Select-Object MainWindowTitle
#DAME   Write-Debug "* * * WindowTitle of ALL GUI : $NowTotalWindowTItle" 
#   Write-Host "* * * WindowTitle of ALL GUI : " $NowTotalWindowTItle

   Start-Sleep -Milliseconds 5000

   $now = Get-Date

   if (($now - $timeup) -le 0)
   {
      Write-Debug ("MADA.....    Now:" + (Get-Date -Format "HH:mm") + "]  <time:" + $timeup + ">    " + ($now - $timeup))
#      Write-Host ("MADA.....    Now:" + (Get-Date -Format "HH:mm") + "]  <time:" + $timeup + ">    " + ($now - $timeup))  -ForegroundColor DarkGray
   }else{
      Write-Debug "Finish"
      break
   }

   $IdleTime = [PInvoke.Win32.UserInput]::IdleTime
#   Write-Debug "<IdleTime:" + $IdleTime + "<Timeout:" + $Timeout + ">  " + " TimeoutPercent:" + $TimeoutPercent
   $TimeoutPercent = ($IdleTime.TotalSeconds / $Timeout.TotalSeconds) * 100
   if($TimeoutPercent -ge 100){
      $TimeoutPercent=100
   }


   $hwnd = [Win32.Utils]::GetForegroundWindow()
   $null = [Win32.Utils]::GetWindowThreadProcessId($hwnd, [ref] $myPid)
   $WindowTitle = [string](Get-Process | Where-Object ID -eq $myPid | Select-Object MainWindowTitle)      
   Write-Debug "- - - WindowTitle of Current GUI : $WindowTitle"



   Write-Progress -Activity "Show desktop until:"  -CurrentOperation ("[IdleTime:" + $IdleTime + "]  <Timeout:" + $Timeout + ">")  -Status ("Active Window Title [$WindowTitle]") -PercentComplete $TimeoutPercent

   if ($IdleTime -gt $Timeout) {
   #     Write-Debug "A     <IdleTime:" + $IdleTime + "<Timeout:" + $Timeout
   #    #   @{Name=slack; ProcessName=slack; Id=19168; Path=C:\Users\saclalog1\AppData\Local\slack\app-4.26.3\slack.exe; MainWindowTitle=Slack | general | 1_StudyLawlerHorio_20220609}

#      if(    $WindowTitle.Contains("Slack | general") ){
      if(    $WindowTitle.Contains($TargetWindowName) ){
#         Write-Debug "Minumize 1"
#      if(    $WindowTitle.Contains("Chrome") ){
#         Start-Process "Shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}"      Show Desktop DAME   
         $window=[System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
         $windowPattern=$window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
         $windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Minimized)
      }
   }
} while ($true)

#Write-Host "Do exit"

exit