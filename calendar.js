var i=0, f=new ActiveXObject('Scripting.FileSystemObject'), s=new ActiveXObject('WScript.Shell'), p=s.Environment('PROCESS')('PATH').split(';'), ps='powershell.exe';
for(; i<p.length; ++i) if(p[i]!='' && f.FileExists(p[i]+'\\pwsh.exe')){ps='pwsh.exe'; break;}
s.Run("conhost.exe "+ps+" -c \""+
"Add-Type -AN System.Windows.Forms;"+
"if($host.Version.Major -lt 7){$ref='System.Windows.Forms'}"+
"else{$ref='Microsoft.Win32.Registry','System.Windows.Forms','System.ComponentModel.Primitives','System.Windows.Forms.Primitives','System.Threading.Thread'}"+
"Add-Type '"+
"using Microsoft.Win32;"+
"using System;"+
"using System.Runtime.InteropServices;"+
"using System.Threading.Tasks;"+
"using System.Windows.Forms;"+
"public class DLForm:Form{"+
"  [DllImport(\\\"dwmapi.dll\\\")] static extern int DwmSetWindowAttribute(IntPtr w,int a,ref int v,int s);"+
"  [DllImport(\\\"user32.dll\\\")] static extern int SetThreadDpiAwarenessContext(int c);"+
"  public static void SetThreadDpiAware(){"+
"    SetThreadDpiAwarenessContext(-4);"+
"  }"+
"  public void SuspendRedrawWhile(Action act){"+
"    Message m=Message.Create(Handle,11,IntPtr.Zero,IntPtr.Zero);"+
"    DefWndProc(ref m);"+
"    m.WParam=(IntPtr)1;"+
"    try{act();}"+
"    finally{DefWndProc(ref m); Invalidate(); Refresh();}"+
"  }"+
"  bool dM;"+
"  void UpdtDM(){"+
"    DarkMode=(0==(int)Registry.GetValue(@\\\"HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\\\",\\\"AppsUseLightTheme\\\",1));"+
"  }"+
"  public bool DarkMode{"+
"    get{return dM;}"+
"    private set{dM=value;}"+
"  }"+
"  bool dTB;"+
"  public bool DarkTitleBar{"+
"    get{return dTB;}"+
"    set{int v=value?1:0; if(0==DwmSetWindowAttribute(Handle,20,ref v,4)) dTB=value;}"+
"  }"+
"  public event Action DarkModeChanged=delegate{};"+
"  protected override void WndProc(ref Message m){"+
"    base.WndProc(ref m);"+
"    if(m.Msg==26 && m.WParam==IntPtr.Zero && Marshal.PtrToStringAuto(m.LParam)==\\\"ImmersiveColorSet\\\") Task.Run(()=>{UpdtDM(); DarkModeChanged();});"+
"  }"+
"  public DLForm(){"+
"    UpdtDM();"+
"  }"+
"}' -R $ref;"+
"function Show-DLForm{"+
"  param([Parameter(Mandatory=$true)][DLForm]$sdlf_DLForm,"+
"        [ScriptBlock]$sdlf_DarkModeChangedAction={})"+
"  $sdlf_Rs=[RunspaceFactory]::CreateRunspace($Host);"+
"  $sdlf_Rs.Open();"+
"  gv|%{try{$sdlf_Rs.SessionStateProxy.PSVariable.Set($_)}catch{}};"+
"  $sdlf_Ps=[PowerShell]::Create().AddScript({Register-ObjectEvent $sdlf_DLForm DarkModeChanged -A $sdlf_DarkModeChangedAction});"+
"  $sdlf_Ps.Runspace=$sdlf_Rs;"+
"  $sdlf_Res=$sdlf_Ps.BeginInvoke();"+
"  $sdlf_DLForm.ShowDialog();"+
"  [void]$sdlf_Ps.EndInvoke($sdlf_Res);"+
"  $sdlf_Rs.Close();"+
"  $sdlf_Rs.Dispose();"+
"  $sdlf_Ps.Dispose();"+
"}"+
"[DLForm]::SetThreadDpiAware();"+
"$wnd=[DLForm]@{AutoSize=$true; FormBorderStyle='Fixed3D'; Icon=[Drawing.Icon]::ExtractAssociatedIcon((gps -Id $PID).Path); MaximizeBox=$false; Text='Calendar'};"+
"$wnd.add_Shown({$cal.Left=($wnd.ClientSize.Width-$cal.Width)/2; $cal.Top=$chk.Height; $wnd.Activate()});"+
"$chk=[Windows.Forms.CheckBox]@{AutoSize=$true; Font='Microsoft Sans Serif,8'; Padding='20,10,0,20'; Text='Always on top'};"+
"$chk.add_CheckedChanged({$wnd.TopMost=$chk.Checked});"+
"$cal=[Windows.Forms.MonthCalendar]@{CalendarDimensions='3,4'; Margin='0,0,0,20'; ScrollChange=3; ShowToday=$false; ShowWeekNumbers=$true};"+
"$cal.Font=\\\"Microsoft Sans Serif,$($cal.Font.Size)\\\";"+
"$wnd.Controls.AddRange(@($chk,$cal));"+
"$settheme={"+
"  if($wnd.DarkMode){$thm='#3F4144','#1F2023','White','#CFD1D4','Black'}else{$thm='#DFE0E3','#EFF0F2','Black','#3F4144','White'}"+
"  $wnd.SuspendRedrawWhile({"+
"    $cal.ForeColor=$thm[2]; $cal.TitleForeColor=$thm[4]; $chk.ForeColor=$thm[2];"+
"    $cal.BackColor=$thm[1]; $cal.TitleBackColor=$thm[3]; $wnd.BackColor=$thm[0]; $wnd.DarkTitleBar=$wnd.DarkMode"+
"  });"+
"};"+
"&$settheme;"+
"Show-DLForm $wnd $settheme;\"",0,false)
