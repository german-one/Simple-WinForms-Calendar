@set p=&for %%i in (pwsh.exe powershell.exe) do @if not defined p set p=%%~$PATH:i
@start /min conhost "%p%" -w Hidden -c ^"Add-Type -AN System.Windows.Forms;Add-Type i '[DllImport(\"dwmapi\")]public static extern void DwmSetWindowAttribute(IntPtr w,int a,int[] v,int s);'-NS p;^
$d=-not(gp HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize).AppsUseLightTheme;if($d){$a='#3F4144','#1F2023','White','#CFD1D4','Black'}else{$a='#DFE0E3','#EFF0F2','Black','#3F4144','White'}^
$f=[Windows.Forms.Form]@{AutoSize=1;BackColor=$a[0];FormBorderStyle=2;Icon=[Drawing.Icon]::ExtractAssociatedIcon('%p%');MaximizeBox=0;Text='Calendar'};^
$f.add_Shown({[p.i]::DwmSetWindowAttribute($f.Handle,20,@($d),4);$s=$f.ClientSize;$c.Left=($s.Width-$c.Width)/2;$c.Top=($s.Height-$c.Height)/2+$b.Height});^
$c=[Windows.Forms.MonthCalendar]@{BackColor=$a[1];CalendarDimensions='3,4';ForeColor=$a[2];ScrollChange=3;ShowToday=0;ShowWeekNumbers=1;TitleBackColor=$a[3];TitleForeColor=$a[4]};^
$b=[Windows.Forms.CheckBox]@{AutoSize=1;ForeColor=$a[2];Padding='20,10,0,0';Text='Always on top'};$b.add_CheckedChanged({$f.TopMost=$b.Checked});$f.Controls.AddRange(@($b,$c));$f.ShowDialog();^"
