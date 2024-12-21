@set "light=1"&for /f "tokens=3" %%i in ('2^>nul reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -v "AppsUseLightTheme"') do @set /a "light=%%i"
@if %light% equ 0 (set "dark=1"&set "fBC='#3F4144'"&set "cBC='#1F2023'"&set "cFC='White'"&set "cTBC='#CFD1D4'"&set "cTFC='Black'")^
 else set "dark=0"&set "fBC='#DFE0E3'"&set "cBC='#EFF0F2'"&set "cFC='Black'"&set "cTBC='#3F4144'"&set "cTFC='White'"
@for %%i in ("pwsh.exe") do @if "%%~$PATH:i"=="" (set "ps=powershell") else set "ps=pwsh"
@start /min conhost %ps% -w Hidden -c ^"^
 Add-Type -AN System.Windows.Forms;$w=Add-Type -Name W '[DllImport(\"dwmapi\")]public static extern void DwmSetWindowAttribute(IntPtr w,int a,ref int v,int s);' -PassThru;^
 $f=New-Object Windows.Forms.Form -Property @{AutoSize=$true;BackColor=%fBC%;StartPosition='CenterScreen';Text='Calendar'};^
 $c=New-Object Windows.Forms.MonthCalendar -Property @{BackColor=%cBC%;CalendarDimensions='3,4';ForeColor=%cFC%;ScrollChange=3;ShowToday=$false;ShowWeekNumbers=$true;TitleBackColor=%cTBC%;TitleForeColor=%cTFC%};^
 $f.add_Shown({$w32bool=%dark%;$w::DwmSetWindowAttribute($this.Handle,20,[ref]$w32bool,4);$c.Left=($this.ClientSize.Width-$c.Width)/2;$c.Top=($this.ClientSize.Height-$c.Height)/2});$f.Controls.Add($c);$f.ShowDialog();^"
