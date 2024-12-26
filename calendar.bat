@setlocal DisableDelayedExpansion&set "dark=0"&for /f "tokens=3" %%i in ('2^>nul reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -v "AppsUseLightTheme"') do @set /a "dark=!%%i"
@if %dark%==1 (set "fBC='#3F4144'"&set "cBC='#1F2023'"&set "cFC='White'"&set "tBC='#CFD1D4'"&set "tFC='Black'") else set "fBC='#DFE0E3'"&set "cBC='#EFF0F2'"&set "cFC='Black'"&set "tBC='#3F4144'"&set "tFC='White'"
@for %%i in ("pwsh.exe") do @if "%%~$PATH:i"=="" (set "ps=powershell") else set "ps=pwsh"
@start /min conhost.exe %ps%.exe -w Hidden -c ^"Add-Type -AN System.Windows.Forms;$w=Add-Type -Name W '^
 [DllImport(\"dwmapi\")]public static extern void DwmSetWindowAttribute(IntPtr w,int a,ref int v,int s);[DllImport(\"user32\")]public static extern void SetWindowPos(IntPtr w,int i,int X,int Y,int x,int y,int f);' -PassThru;^
 $f=New-Object Windows.Forms.Form -Property @{AutoSize=$true;BackColor=%fBC%;StartPosition='CenterScreen';Text='Calendar'};^
 $ck=New-Object Windows.Forms.CheckBox -Property @{AutoSize=$true;ForeColor=%cFC%;Padding='20,10,0,0';Text='Always on top'};^
 $c=New-Object Windows.Forms.MonthCalendar -Property @{BackColor=%cBC%;CalendarDimensions='3,4';ForeColor=%cFC%;ScrollChange=3;ShowToday=$false;ShowWeekNumbers=$true;TitleBackColor=%tBC%;TitleForeColor=%tFC%};^
 $f.add_Shown({$useDark=%dark%;$w::DwmSetWindowAttribute($this.Handle,20,[ref]$useDark,4);$c.Left=($this.ClientSize.Width-$c.Width)/2;$c.Top=($this.ClientSize.Height-$c.Height)/2+$ck.Height});^
 $ck.add_CheckedChanged({$w::SetWindowPos($this.Parent.Handle,$this.Checked-2,0,0,0,0,3);});$f.Controls.AddRange(@($ck,$c));$f.ShowDialog();^"
