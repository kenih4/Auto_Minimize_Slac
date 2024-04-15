#   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process



# MEMO  
# �E�B���h�E�^�C�g���̈ꗗ��\��
# Get-Process | Where-Object {$_.MainWindowTitle -ne ""} | Select-Object MainWindowTitle
#


# Parameters
Param (
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



$nowH = (Get-Date).Hour
if($nowH -ge 1 -And $nowH -lt 9){
    $str = [string](Get-Date).ToString("yyyy/MM/dd")  + " 09:00:00"
}elseif($nowH -gt 9 -And $nowH -lt 17){
#    $str = [string](Get-Date).ToString("yyyy/MM/dd")  + " 17:00:00"
    $str = [string](Get-Date).ToString("yyyy/MM/dd")  + " 17:00:00"
}elseif($nowH -gt 17 -And $nowH -lt 24){
    $tommorow = (Get-Date).AddDays(1)
    $str = [string]($tommorow).ToString("yyyy/MM/dd")  + " 1:00:00"
}elseif($nowH -eq 0){
   $str = [string](Get-Date).ToString("yyyy/MM/dd")  + " 1:00:00"
}else{
   return
}

#Write-Host "str:  " + [string]$str
$timeup = [DateTime]::ParseExact($str,"yyyy/MM/dd HH:mm:ss", $null);
#Write-Host "START   timeup:   " + $timeup



# Main Routine
#  zzz $timeup Timeup


$Timeout = [TimeSpan]::FromMinutes($TimeoutMin)
#$Timeout = [TimeSpan]::FromMinutes(0.05)
do {

   Start-Sleep -Milliseconds 3000

   $now = Get-Date
   if (($now - $timeup) -le 0)
   {
#      Write-Host ("MADA Now:" + (Get-Date -Format "HH:mm") + "]  <time:" + $timeup + ">    " + ($now - $timeup))  -ForegroundColor DarkGray
   }else{
#      Write-Host "Finish"
      break
   }



   $IdleTime = [PInvoke.Win32.UserInput]::IdleTime
#   Write-Host "<IdleTime:" + $IdleTime + "<Timeout:" + $Timeout + ">  " + " TimeoutPercent:" + $TimeoutPercent
   $TimeoutPercent = ($IdleTime.TotalSeconds / $Timeout.TotalSeconds) * 100
   if($TimeoutPercent -ge 100){
      $TimeoutPercent=100
   }
   Write-Progress -Activity "Show desktop until:" -Status ("[IdleTime:" + $IdleTime + "]  <Timeout:" + $Timeout + ">") -PercentComplete $TimeoutPercent

   if ($IdleTime -gt $Timeout) {
   #     Write-Host "A     <IdleTime:" + $IdleTime + "<Timeout:" + $Timeout
   #    #   @{Name=slack; ProcessName=slack; Id=19168; Path=C:\Users\saclalog1\AppData\Local\slack\app-4.26.3\slack.exe; MainWindowTitle=Slack | general | 1_StudyLawlerHorio_20220609}
      $hwnd = [Win32.Utils]::GetForegroundWindow()
      $null = [Win32.Utils]::GetWindowThreadProcessId($hwnd, [ref] $myPid)
      $WindowTitle = [string](Get-Process | Where-Object ID -eq $myPid | Select-Object MainWindowTitle)
      
      Write-Host "WindowTitle"
      Write-Host $WindowTitle

#      if(    $WindowTitle.Contains("Slack | general") ){
      if(    $WindowTitle.Contains("general") ){
#         Write-Host "Minumize 1"
#      if(    $WindowTitle.Contains("Chrome") ){
#         Start-Process "Shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}"      Show Desktop DAME   
         $window=[System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
         $windowPattern=$window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
         $windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Minimized)
#         Write-Host "Minumize 2"
      }else{
#         Write-Host "Not Slac"
      }
   }else{
   #     Write-Host "B     <IdleTime:" + $IdleTime + "<Timeout:" + $Timeout
   }
#   Write-Host "<IdleTime:" + $IdleTime + "<Timeout:" + $Timeout + ">  "
} while ($true)

#Write-Host "Do exit"

exit