@set ps=&for %%i in (pwsh.exe powershell.exe) do @if not defined ps set ps=%%~$PATH:i
@start /min conhost.exe "%ps%" -w Hidden -c ^"^
$zoom = 1.0;^
Add-Type -AN System.Windows.Forms;^
if ($host.Version.Major -lt 7) {^
  $ref = 'System.Windows.Forms';^
  $fct = 0.7;^
}^
else {^
  $ref = 'Microsoft.Win32.Registry', 'System.ComponentModel.Primitives', 'System.Threading.Thread', 'System.Windows.Forms', 'System.Windows.Forms.Primitives';^
  $fct = 0.62;^
}^
Add-Type '^
using Microsoft.Win32;^
using System;^
using System.Runtime.InteropServices;^
using System.Threading.Tasks;^
using System.Windows.Forms;^
public class DLForm : Form {^
  [DllImport(\"dwmapi.dll\")] private static extern int DwmSetWindowAttribute(IntPtr w, int a, ref int v, int s);^
  [DllImport(\"user32.dll\")] private static extern int SetThreadDpiAwarenessContext(int c);^
  private const string customThemeKey = @\"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\";^
  private const int DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = -2;^
  private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;^
  private const int FALSE = 0;^
  private const int S_OK = 0;^
  private const int TRUE = 1;^
  private const int WM_SETREDRAW = 11;^
  private const int WM_SETTINGCHANGE = 26;^
  private static readonly IntPtr wpFALSE = IntPtr.Zero;^
  private static readonly IntPtr wpTRUE = new IntPtr(1);^
  public static void SetThreadDpiAware() {^
    SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_SYSTEM_AWARE);^
  }^
  public void SuspendRedrawWhile(Action act) {^
    Message msg = Message.Create(Handle, WM_SETREDRAW, wpFALSE, IntPtr.Zero);^
    DefWndProc(ref msg);^
    msg.WParam = wpTRUE;^
    try {^
      act();^
    }^
    finally {^
      DefWndProc(ref msg);^
      Invalidate();^
      Refresh();^
    }^
  }^
  private bool dM;^
  private void UpdtDM() {^
    DarkMode = (FALSE == (int)Registry.GetValue(customThemeKey, \"AppsUseLightTheme\", TRUE));^
  }^
  public bool DarkMode {^
    get {return dM;}^
    private set {dM = value;}^
  }^
  private bool dTB;^
  public bool DarkTitleBar {^
    get {return dTB;}^
    set {^
      int doDark = value ? TRUE : FALSE;^
      if (S_OK == DwmSetWindowAttribute(Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref doDark, Marshal.SizeOf(doDark))) {^
        dTB = value;^
      }^
    }^
  }^
  public event Action DarkModeChanged = delegate {};^
  protected override void WndProc(ref Message msg) {^
    base.WndProc(ref msg);^
    if (msg.Msg == WM_SETTINGCHANGE ^&^&^
        msg.WParam == IntPtr.Zero ^&^&^
        Marshal.PtrToStringAuto(msg.LParam) == \"ImmersiveColorSet\") {^
      Task.Run(() =^> {^
        UpdtDM();^
        DarkModeChanged();^
      });^
    }^
  }^
  public DLForm() {^
    UpdtDM();^
  }^
}' -ReferencedAssemblies $ref;^
function Show-DLForm {^
  param(^
    [Parameter(Mandatory = $true)][DLForm]$sdlf_DLForm,^
    [ScriptBlock]$sdlf_DarkModeChangedAction = {}^
  )^
  $sdlf_Rs = [RunspaceFactory]::CreateRunspace($Host);^
  $sdlf_Rs.Open();^
  Get-Variable ^| ForEach-Object{^
    try {^
      $sdlf_Rs.SessionStateProxy.PSVariable.Set($_);^
    }^
    catch {}^
  };^
  $sdlf_Ps = [PowerShell]::Create().AddScript({^
    Register-ObjectEvent $sdlf_DLForm DarkModeChanged -Action $sdlf_DarkModeChangedAction;^
  });^
  $sdlf_Ps.Runspace = $sdlf_Rs;^
  $sdlf_Res = $sdlf_Ps.BeginInvoke();^
  $sdlf_DLForm.ShowDialog();^
  [void]$sdlf_Ps.EndInvoke($sdlf_Res);^
  $sdlf_Rs.Close();^
  $sdlf_Rs.Dispose();^
  $sdlf_Ps.Dispose();^
}^
[DLForm]::SetThreadDpiAware();^
$wnd = [DLForm]@{^
  AutoSize = $true;^
  AutoSizeMode = 'GrowAndShrink';^
  FormBorderStyle = 'Fixed3D';^
  Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $PID).Path);^
  MaximizeBox = $false;^
  Text = 'Calendar';^
};^
$wnd.add_Shown({^
  $rba.Left = 10;^
  $rbl.Left = $rba.Right + 10;^
  $rbd.Left = $rbl.Right + 10;^
  $rbd.Top = $rbl.Top = $rba.Top = 20;^
  $grp.Location = \"$($cal.Right - $grp.Width - 20), 10\";^
  $cal.Top = $grp.Bottom + 10;^
  $wnd.Activate();^
});^
$chk = [Windows.Forms.CheckBox]@{^
  AutoSize = $true;^
  Font = \"Segoe UI, $(6 * $zoom)\";^
  Padding = '20, 10, 0, 20';^
  Text = 'Always on top';^
};^
$chk.add_CheckedChanged({^
  $wnd.TopMost = $chk.Checked;^
});^
$cal = [Windows.Forms.MonthCalendar]@{^
  CalendarDimensions = '3, 4';^
  Margin = '0, 0, 0, 20';^
  ScrollChange = 3;^
  ShowToday = $false;^
  ShowWeekNumbers = $true;^
};^
$cal.Font = \"$($chk.Font.FontFamily), $($cal.Font.Size * $fct * $zoom)\";^
$grp = [Windows.Forms.GroupBox]@{^
  AutoSize = $true;^
  AutoSizeMode = 'GrowAndShrink';^
  Font =$cal.Font;^
  Text = 'Theme';^
};^
$rba = [Windows.Forms.RadioButton]@{^
  AutoSize = $true;^
  Checked = $true;^
  Font = $chk.Font;^
  Margin = '0, 0, 0, 0';^
  Text = 'Auto';^
};^
$rba.add_CheckedChanged({^
  if ($rba.Checked) {^&$settheme};^
});^
$rbl = [Windows.Forms.RadioButton]@{^
  AutoSize = $true;^
  Font = $chk.Font;^
  Margin = '0, 0, 0, 0';^
  Text = 'Light';^
};^
$rbl.add_CheckedChanged({^
  if ($rbl.Checked) {^&$settheme};^
});^
$rbd = [Windows.Forms.RadioButton]@{^
  AutoSize = $true;^
  Font = $chk.Font;^
  Margin = '0, 0, 0, 0';^
  Text = 'Dark';^
};^
$rbd.add_CheckedChanged({^
  if ($rbd.Checked) {^&$settheme};^
});^
$grp.Controls.AddRange(@($rba, $rbl, $rbd));^
$wnd.Controls.AddRange(@($chk, $grp, $cal));^
$settheme = {^
  $dark = $(if ($rba.Checked) {$wnd.DarkMode} else {$rbd.Checked});^
  if ($dark) {^
    $pal = 'White', 'Black', '#1F2023', '#CFD1D4', '#3F4144';^
  }^
  else {^
    $pal = 'Black', 'White', '#EFF0F2', '#3F4144', '#DFE0E3';^
  }^
  $wnd.SuspendRedrawWhile({^
    $cal.ForeColor = $grp.ForeColor = $chk.ForeColor = $pal[0];^
    $cal.TitleForeColor = $pal[1];^
    $cal.BackColor = $pal[2];^
    $cal.TitleBackColor = $pal[3];^
    $wnd.BackColor = $pal[4];^
    $wnd.DarkTitleBar = $dark;^
  });^
};^
&$settheme;^
Show-DLForm $wnd $settheme;^"
