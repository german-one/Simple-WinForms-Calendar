var i = 0, fs = new ActiveXObject('Scripting.FileSystemObject'), ws = new ActiveXObject('WScript.Shell'), ap = ws.Environment('PROCESS')('PATH').split(';'), ps = 'powershell.exe';
for (; i < ap.length; ++i) {if (ap[i] != '' && fs.FileExists(ap[i] + '\\pwsh.exe')) {ps = 'pwsh.exe'; break;}}
ws.Run("conhost.exe " + ps + " -c \"" +
"$zoom = 1.0;" +
"Add-Type -AN System.Windows.Forms;" +
"if ($host.Version.Major -lt 7) {" +
"  $ref = 'System.Windows.Forms';" +
"  $fct = 0.7;" +
"}" +
"else {" +
"  $ref = 'Microsoft.Win32.Registry', 'System.ComponentModel.Primitives', 'System.Threading.Thread', 'System.Windows.Forms', 'System.Windows.Forms.Primitives';" +
"  $fct = 0.62;" +
"}" +
"Add-Type '" +
"using Microsoft.Win32;" +
"using System;" +
"using System.Runtime.InteropServices;" +
"using System.Threading.Tasks;" +
"using System.Windows.Forms;" +
"public class DLForm : Form {" +
"  [DllImport(\\\"dwmapi.dll\\\")] private static extern int DwmSetWindowAttribute(IntPtr w, int a, ref int v, int s);" +
"  [DllImport(\\\"user32.dll\\\")] private static extern int SetThreadDpiAwarenessContext(int c);" +
"  private const string customThemeKey = @\\\"HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\\\";" +
"  private const int DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = -2;" +
"  private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;" +
"  private const int FALSE = 0;" +
"  private const int S_OK = 0;" +
"  private const int TRUE = 1;" +
"  private const int WM_SETREDRAW = 11;" +
"  private const int WM_SETTINGCHANGE = 26;" +
"  private static readonly IntPtr wpFALSE = IntPtr.Zero;" +
"  private static readonly IntPtr wpTRUE = new IntPtr(1);" +
"  public static void SetThreadDpiAware() {" +
"    SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_SYSTEM_AWARE);" +
"  }" +
"  public void SuspendRedrawWhile(Action act) {" +
"    Message msg = Message.Create(Handle, WM_SETREDRAW, wpFALSE, IntPtr.Zero);" +
"    DefWndProc(ref msg);" +
"    msg.WParam = wpTRUE;" +
"    try {" +
"      act();" +
"    }" +
"    finally {" +
"      DefWndProc(ref msg);" +
"      Invalidate();" +
"      Refresh();" +
"    }" +
"  }" +
"  private bool dM;" +
"  private void UpdtDM() {" +
"    DarkMode = (FALSE == (int)Registry.GetValue(customThemeKey, \\\"AppsUseLightTheme\\\", TRUE));" +
"  }" +
"  public bool DarkMode {" +
"    get {return dM;}" +
"    private set {dM = value;}" +
"  }" +
"  private bool dTB;" +
"  public bool DarkTitleBar {" +
"    get {return dTB;}" +
"    set {" +
"      int doDark = value ? TRUE : FALSE;" +
"      if (S_OK == DwmSetWindowAttribute(Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref doDark, Marshal.SizeOf(doDark))) {" +
"        dTB = value;" +
"      }" +
"    }" +
"  }" +
"  public event Action DarkModeChanged = delegate {};" +
"  protected override void WndProc(ref Message msg) {" +
"    base.WndProc(ref msg);" +
"    if (msg.Msg == WM_SETTINGCHANGE &&" +
"        msg.WParam == IntPtr.Zero &&" +
"        Marshal.PtrToStringAuto(msg.LParam) == \\\"ImmersiveColorSet\\\") {" +
"      Task.Run(() => {" +
"        UpdtDM();" +
"        DarkModeChanged();" +
"      });" +
"    }" +
"  }" +
"  public DLForm() {" +
"    UpdtDM();" +
"  }" +
"}' -ReferencedAssemblies $ref;" +
"function Show-DLForm {" +
"  param(" +
"    [Parameter(Mandatory = $true)][DLForm]$sdlf_DLForm," +
"    [ScriptBlock]$sdlf_DarkModeChangedAction = {}" +
"  )" +
"  $sdlf_Rs = [RunspaceFactory]::CreateRunspace($Host);" +
"  $sdlf_Rs.Open();" +
"  Get-Variable | ForEach-Object{" +
"    try {" +
"      $sdlf_Rs.SessionStateProxy.PSVariable.Set($_);" +
"    }" +
"    catch {}" +
"  };" +
"  $sdlf_Ps = [PowerShell]::Create().AddScript({" +
"    Register-ObjectEvent $sdlf_DLForm DarkModeChanged -Action $sdlf_DarkModeChangedAction;" +
"  });" +
"  $sdlf_Ps.Runspace = $sdlf_Rs;" +
"  $sdlf_Res = $sdlf_Ps.BeginInvoke();" +
"  $sdlf_DLForm.ShowDialog();" +
"  [void]$sdlf_Ps.EndInvoke($sdlf_Res);" +
"  $sdlf_Rs.Close();" +
"  $sdlf_Rs.Dispose();" +
"  $sdlf_Ps.Dispose();" +
"}" +
"[DLForm]::SetThreadDpiAware();" +
"$wnd = [DLForm]@{" +
"  AutoSize = $true;" +
"  AutoSizeMode = 'GrowAndShrink';" +
"  FormBorderStyle = 'Fixed3D';" +
"  Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $PID).Path);" +
"  MaximizeBox = $false;" +
"  Text = 'Calendar';" +
"};" +
"$wnd.add_Shown({" +
"  $rba.Left = 10;" +
"  $rbl.Left = $rba.Right + 10;" +
"  $rbd.Left = $rbl.Right + 10;" +
"  $rbd.Top = $rbl.Top = $rba.Top = 20;" +
"  $grp.Location = \\\"$($wnd.ClientSize.Width - $grp.Width - 25), 5\\\";" +
"  $cal.Top = $grp.Bottom + 15;" +
"  $wnd.Activate();" +
"});" +
"$btn = [Windows.Forms.Button]@{" +
"  AutoSize = $true;" +
"  AutoSizeMode = 'GrowAndShrink';" +
"  FlatStyle = 'Flat';" +
"  Font = \\\"Segoe UI, $(7 * $zoom)\\\";" +
"  Location = '25, 20';" +
"  Text = 'This Year';" +
"};" +
"$btn.add_Click({" +
"  $now = Get-Date;" +
"  $first = Get-Date -Year $now.Year -Month 1 -Day 1;" +
"  $last = Get-Date -Year $now.Year -Month 12 -Day 31;" +
"  $cal.SelectionRange = @{Start = $first; End = $first;};" +
"  $cal.SelectionRange = @{Start = $last; End = $last;};" +
"  $cal.SelectionRange = @{Start = $now; End = $now;};" +
"});" +
"$chk = [Windows.Forms.CheckBox]@{" +
"  AutoSize = $true;" +
"  FlatStyle = 'Flat';" +
"  Font = $btn.Font;" +
"  Location = '200, 25';" +
"  Text = 'Always on top';" +
"};" +
"$chk.add_CheckedChanged({" +
"  $wnd.TopMost = $chk.Checked;" +
"});" +
"$grp = [Windows.Forms.GroupBox]@{" +
"  AutoSize = $true;" +
"  AutoSizeMode = 'GrowAndShrink';" +
"  Text = 'Theme';" +
"};" +
"$grp.Font = \\\"$($btn.Font.FontFamily), $($grp.Font.Size * $fct * $zoom)\\\";" +
"$rbprop = @{" +
"  AutoSize = $true;" +
"  FlatStyle = 'Flat';" +
"  Font = $btn.Font;" +
"  Margin = '0, 0, 0, 0';" +
"};" +
"$rbact = {" +
"  if ($this.Checked) {&$settheme}" +
"};" +
"$rba = [Windows.Forms.RadioButton]$rbprop;" +
"$rba.Checked = $true;" +
"$rba.Text = 'Auto';" +
"$rba.add_CheckedChanged($rbact);" +
"$rbl = [Windows.Forms.RadioButton]$rbprop;" +
"$rbl.Text = 'Light';" +
"$rbl.add_CheckedChanged($rbact);" +
"$rbd = [Windows.Forms.RadioButton]$rbprop;" +
"$rbd.Text = 'Dark';" +
"$rbd.add_CheckedChanged($rbact);" +
"$grp.Controls.AddRange(@($rba, $rbl, $rbd));" +
"$cal = [Windows.Forms.MonthCalendar]@{" +
"  CalendarDimensions = '3, 4';" +
"  Font =$grp.Font;" +
"  Margin = '0, 0, 0, 15';" +
"  ScrollChange = 3;" +
"  ShowToday = $false;" +
"  ShowWeekNumbers = $true;" +
"};" +
"$wnd.Controls.AddRange(@($btn, $chk, $grp, $cal));" +
"$settheme = {" +
"  $dark = $(if ($rba.Checked) {$wnd.DarkMode} else {$rbd.Checked});" +
"  if ($dark) {" +
"    $pal = 'White', 'Black', '#1F2023', '#CFD1D4', '#3F4144'" +
"  }" +
"  else {" +
"    $pal = 'Black', 'White', '#EFF0F2', '#3F4144', '#DFE0E3'" +
"  }" +
"  $wnd.SuspendRedrawWhile({" +
"    $cal.ForeColor = $grp.ForeColor = $chk.ForeColor = $btn.ForeColor = $pal[0];" +
"    $cal.TitleForeColor = $pal[1];" +
"    $cal.BackColor = $btn.BackColor = $pal[2];" +
"    $cal.TitleBackColor = $pal[3];" +
"    $wnd.BackColor = $pal[4];" +
"    $wnd.DarkTitleBar = $dark;" +
"  });" +
"};" +
"&$settheme;" +
"Show-DLForm $wnd $settheme;\"", 0, false);
