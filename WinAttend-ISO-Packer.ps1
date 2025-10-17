#requires -version 5.1
<#
  Windows ISO Repacker — WPF Dark UI (WPF-native dialogs)
  - WPF UI with dark theme
  - WPF OpenFileDialog (no WinForms)
  - oscdimg.exe at %TEMP%\WinRepacker\oscdimg.exe (or side-by-side / ADK)
  - Logs -> %TEMP%\WinRepacker\logs
  - Output ISO next to the script/EXE
  - STA/admin relaunch; ps2exe-friendly
#>

param(
  [string] $IsoPath,
  [string] $UnattendXmlPath,
  [switch] $IncludeOEM,
  [switch] $SkipOEM,
  [string] $OscdimgPath,
  [string] $AnswerFilesDir,
  [switch] $ForceDark,
  [switch] $ForceLight
)

# ---------- App metadata ----------
$AppTitle   = 'WinAttend ISO Packer'
$AppVersion = 'v1'
$AppAuthor  = 'Forz'

# ---------- Relaunch STA/Admin ----------
function Get-BaseDir { if ($PSScriptRoot -and (Test-Path -LiteralPath $PSScriptRoot)) { $PSScriptRoot } else { [AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\','/') } }
function Test-IsCompiled { try { $p=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName; ($p -like '*.exe' -and -not ($PSCommandPath -like '*.ps1')) } catch { $false } }
function Get-ForwardArgs { @($args) }
function Restart-InStaAndOrAdmin {
  $needsAdmin = -not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  $needsSta   = [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA'
  if (-not $needsAdmin -and -not $needsSta) { return }
  $fwd = Get-ForwardArgs
  if (Test-IsCompiled) {
    $exe = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $sp = @{ FilePath=$exe; ArgumentList=$fwd }; if ($needsAdmin) { $sp.Verb='RunAs' }; Start-Process @sp | Out-Null
  } else {
    $sh = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    $al = @('-NoProfile','-ExecutionPolicy','Bypass'); if ($needsSta){$al+='-STA'}; $al += @('-File',"`"$PSCommandPath`"") + $fwd
    $sp = @{ FilePath=$sh; ArgumentList=$al }; if ($needsAdmin){$sp.Verb='RunAs'}; Start-Process @sp | Out-Null
  }
  exit
}
Restart-InStaAndOrAdmin

# ---------- WPF assemblies ----------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ---------- WPF DoEvents helper ----------
function Invoke-DoEventsWpf {
  $frame = New-Object System.Windows.Threading.DispatcherFrame
  [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
    [System.Windows.Threading.DispatcherPriority]::Background,
    [System.Action]{ $frame.Continue = $false }
  ) | Out-Null
  [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

# ---------- Console show/hide helpers (valid MemberDefinition) ----------
try {
  Add-Type -Namespace Win32 -Name Native -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern System.IntPtr GetConsoleWindow();

[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@ -ErrorAction Stop
} catch {}
function Hide-ConsoleWindow {
  try { $h=[Win32.Native]::GetConsoleWindow(); if($h -ne [IntPtr]::Zero){ [Win32.Native]::ShowWindow($h,0) | Out-Null } } catch {}
}
function Minimize-ConsoleWindow {
  try { $h=[Win32.Native]::GetConsoleWindow(); if($h -ne [IntPtr]::Zero){ [Win32.Native]::ShowWindow($h,6) | Out-Null } } catch {}
}

# ---------- Theme ----------
function Get-SystemDarkPref {
  if ($ForceDark) { return $true }
  if ($ForceLight){ return $false }
  try {
    $reg='HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
    (Get-ItemProperty -Path $reg -Name AppsUseLightTheme -ErrorAction Stop).AppsUseLightTheme -eq 0
  } catch { $true }
}
$IsDark = Get-SystemDarkPref
$bg     = if ($IsDark) { '#202225' } else { '#FFFFFFFF' }
$panel  = if ($IsDark) { '#2B2D31' } else { '#FFF5F5F5' }
$input  = if ($IsDark) { '#32353B' } else { '#FFFFFFFF' }
$border = if ($IsDark) { '#444850' } else { '#FFCCCCCC' }
$text   = if ($IsDark) { '#FFFFFFFF' } else { '#FF111111' }
$muted  = if ($IsDark) { '#FFB6B6B6' } else { '#FF444444' }
$accent = '#FF0078D7'

# ---------- Main XAML ----------
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$AppTitle"
        Height="640" Width="960" MinHeight="560" MinWidth="860"
        WindowStartupLocation="CenterScreen"
        Background="$bg" Foreground="$text" FontFamily="Segoe UI">
  <Window.Resources>
    <Style TargetType="GroupBox">
      <Setter Property="Foreground" Value="$text"/>
      <Setter Property="BorderBrush" Value="$border"/>
      <Setter Property="Margin" Value="12,8,12,8"/>
      <Setter Property="Padding" Value="8"/>
      <Setter Property="Background" Value="$panel"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="GroupBox">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Background="$bg" Padding="4,0,4,0" HorizontalAlignment="Left">
                <TextBlock Text="{TemplateBinding Header}" FontWeight="SemiBold" Foreground="$muted"/>
              </Border>
              <Border Grid.Row="1" BorderBrush="$border" BorderThickness="1" CornerRadius="6" Background="$panel" Padding="8">
                <ContentPresenter/>
              </Border>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="$input"/>
      <Setter Property="Foreground" Value="$text"/>
      <Setter Property="BorderBrush" Value="$border"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="$text"/>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Margin" Value="8,0,0,0"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="Background" Value="$panel"/>
      <Setter Property="Foreground" Value="$text"/>
      <Setter Property="BorderBrush" Value="$border"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="$accent"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderBrush" Value="$accent"/>
    </Style>
    <Style TargetType="ScrollViewer">
      <Setter Property="Background" Value="$panel"/>
    </Style>
  </Window.Resources>

  <DockPanel>
    <Border DockPanel.Dock="Top" Background="$panel" BorderBrush="$border" BorderThickness="0,0,0,1" Padding="12,10">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
        <TextBlock Text="$AppTitle" FontWeight="SemiBold" FontSize="16"/>
        <TextBlock Text="$AppVersion  -  $AppAuthor" Foreground="$muted" Margin="8,0,0,0"/>
      </StackPanel>
    </Border>

    <Grid Margin="0,8,0,8">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

        <GroupBox Header="Inputs" Grid.Row="0">
        <Grid>
            <!-- Five rows: ISO, ISO hint, XML, XML hint, OEM -->
            <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>  <!-- ISO row -->
            <RowDefinition Height="Auto"/>  <!-- ISO hint -->
            <RowDefinition Height="Auto"/>  <!-- XML row -->
            <RowDefinition Height="Auto"/>  <!-- XML hint -->
            <RowDefinition Height="Auto"/>  <!-- OEM block -->
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
            <ColumnDefinition Width="220"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="140"/>
            </Grid.ColumnDefinitions>

            <!-- ISO -->
            <TextBlock Grid.Row="0" Grid.Column="0" Text="Windows ISO:" VerticalAlignment="Center" Margin="0,8,0,0"/>
            <TextBox   Grid.Row="0" Grid.Column="1" Name="IsoPathBox" Margin="0,8,0,0"
                    IsReadOnly="True" Focusable="False" IsHitTestVisible="False" Cursor="Arrow"/>
            <Button    Grid.Row="0" Grid.Column="2" Content="Browse..." Name="IsoBrowseBtn" Margin="8,8,0,0" ToolTip="Browse for a Windows .iso"/>

            <!-- ISO hint -->
            <TextBlock Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2"
                    Text="Pick a Windows .iso. Contents are copied to a working folder next to the EXE/PS1 (you can reuse it on the next run)."
                    Foreground="$muted" TextWrapping="Wrap" Margin="0,6,0,0"/>

            <!-- XML -->
            <TextBlock Grid.Row="2" Grid.Column="0" Text="Answer file (autounattend):" VerticalAlignment="Center" Margin="0,8,0,0"/>
            <TextBox   Grid.Row="2" Grid.Column="1" Name="XmlPathBox" Margin="0,8,0,0"
                    IsReadOnly="True" Focusable="False" IsHitTestVisible="False" Cursor="Arrow"/>
            <Button    Grid.Row="2" Grid.Column="2" Content="Select..." Name="XmlSelectBtn" Margin="8,8,0,0" ToolTip="Pick your autounattend.xml"/>

            <!-- XML hint -->
            <TextBlock Grid.Row="3" Grid.Column="1" Grid.ColumnSpan="2"
                    Text="If an 'answerfiles' folder exists next to the EXE/PS1, the picker opens there. Your selection is copied as 'autounattend.xml' to the ISO root."
                    Foreground="$muted" TextWrapping="Wrap" Margin="0,6,0,0"/>

            <!-- OEM -->
            <StackPanel Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="3" Orientation="Vertical" Margin="0,12,0,0">
            <CheckBox Name="OemCheck" Content="Include `$OEM$ folder"/>
            <TextBlock Text="Place a folder named `$OEM$ next to the EXE/PS1. Inside, the following are mapped: `$`$ -> C:\Windows  `$1 -> C:\"
                        Foreground="$muted" TextWrapping="Wrap" Margin="24,4,0,0"/>
            </StackPanel>
        </Grid>
        </GroupBox>


      <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="12,4,12,8" HorizontalAlignment="Left">
        <Button Name="StartBtn" Content="Start" Style="{StaticResource PrimaryButton}" Width="160" Height="34"/>
      </StackPanel>

      <GroupBox Header="Log" Grid.Row="2" Margin="12,0,12,8">
        <Grid>
          <RichTextBox Name="LogBox" VerticalScrollBarVisibility="Auto" IsReadOnly="True" FontFamily="Consolas" FontSize="12" Background="$input" Foreground="$text"/>
        </Grid>
      </GroupBox>
    </Grid>
  </DockPanel>
</Window>
"@

# ---------- Build window ----------
$Window = [Windows.Markup.XamlReader]::Parse($xaml)

# Hide/minimize console now that WPF is up (optional)
Hide-ConsoleWindow   # or Minimize-ConsoleWindow

# ---------- Find controls ----------
$IsoPathBox  = $Window.FindName('IsoPathBox')
$IsoBrowseBtn= $Window.FindName('IsoBrowseBtn')
$XmlPathBox  = $Window.FindName('XmlPathBox')
$XmlSelectBtn= $Window.FindName('XmlSelectBtn')
$OemCheck    = $Window.FindName('OemCheck')
$StartBtn    = $Window.FindName('StartBtn')
$LogBox      = $Window.FindName('LogBox')

# ---------- Paths & logging ----------
$BaseDir       = Get-BaseDir
$TempRoot      = Join-Path $env:TEMP 'WinRepacker'
$EmbeddedOscd  = Join-Path $TempRoot 'tools\oscdimg.exe'
$LogsDir       = Join-Path $TempRoot 'logs'
$null = New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null

function Append-Log([string]$text,[string]$color='White'){
  $para = New-Object Windows.Documents.Paragraph
  $run  = New-Object Windows.Documents.Run($text)
  $brush = New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($color))
  $run.Foreground = $brush
  $para.Inlines.Add($run) | Out-Null
  $LogBox.Document.Blocks.Add($para)
  $LogBox.ScrollToEnd()
  Write-Host $text
}
function Show-Error { param([string]$m) Append-Log "[FAIL] $m" 'Tomato' }
function Show-Info  { param([string]$m) Append-Log "[INFO] $m" 'SkyBlue' }
function Show-Ok    { param([string]$m) Append-Log "[ OK ] $m" 'PaleGreen' }
function Show-Warn  { param([string]$m) Append-Log "[WARN] $m" 'Khaki' }

# ---------- WPF OpenFileDialog ----------
function Show-OpenFile([string]$title,[string]$filter,[string]$initialDir) {
  $ofd = New-Object Microsoft.Win32.OpenFileDialog
  $ofd.Title = $title
  $ofd.Filter = $filter
  if ($initialDir) { $ofd.InitialDirectory = $initialDir }
  $res = $ofd.ShowDialog()
  if ($res) { return $ofd.FileName } else { return $null }
}

# ---------- oscdimg discovery ----------
function Get-OscdimgPath {
  param([string]$UserPath)
  if ($UserPath) { if (Test-Path -LiteralPath $UserPath) { return (Resolve-Path -LiteralPath $UserPath).Path } throw "Specified Oscdimg.exe not found: $UserPath" }

  if (Test-Path -LiteralPath $EmbeddedOscd) { return (Resolve-Path -LiteralPath $EmbeddedOscd).Path }
  foreach ($c in @((Join-Path $BaseDir 'oscdimg.exe'), (Join-Path $BaseDir 'tools\oscdimg.exe'))) {
    if (Test-Path -LiteralPath $c) { return (Resolve-Path -LiteralPath $c).Path }
  }
  $fromPath = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
  if ($fromPath) { return $fromPath.Source }
  foreach ($root in @(
    "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools",
    "C:\Program Files\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools"
  )) {
    if (Test-Path -LiteralPath $root) {
      $found = Get-ChildItem -LiteralPath $root -Recurse -Filter oscdimg.exe -ErrorAction SilentlyContinue |
               Select-Object -First 1 -ExpandProperty FullName
      if ($found) { return $found }
    }
  }
  throw "Oscdimg.exe not found. Install Windows ADK, place it beside the script/EXE, or compile with -embedFiles to %TEMP%\WinRepacker\oscdimg.exe."
}

# ---------- oscdimg progress ----------
function New-OscdimgProgressWindow {
@"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Creating ISO..." Height="160" Width="460" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner"
        Background="$panel" Foreground="$text" FontFamily="Segoe UI">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock x:Name="StepText" Text="Initializing oscdimg..." Margin="0,0,0,10" />
    <ProgressBar x:Name="Bar" Grid.Row="1" Height="18" Minimum="0" Maximum="100" IsIndeterminate="True"/>
    <TextBlock x:Name="Pct" Grid.Row="2" Text="0 %" HorizontalAlignment="Right" Margin="0,8,0,0" />
  </Grid>
</Window>
"@
}


function Invoke-OscdimgWithProgress {
  param(
    [Parameter(Mandatory)][string]$OscdimgPath,
    [Parameter(Mandatory)][string]$WorkDir,
    [Parameter(Mandatory)][string]$OutIso,
    [Parameter(Mandatory)][string]$BootData
  )

  # Build oscdimg argument string
  $argList = @('-m','-o','-u2','-udfver102',"-bootdata:$BootData","`"$WorkDir`"","`"$OutIso`"")
  $cmdArgsEscaped = ($argList -join ' ')

  # Progress window
  $win = [Windows.Markup.XamlReader]::Parse((New-OscdimgProgressWindow))
  $bar = $win.FindName('Bar'); $txt = $win.FindName('StepText'); $pct = $win.FindName('Pct')
  $win.Owner = $Window
  $win.Show(); Invoke-DoEventsWpf

  # Write STDERR to a temp file so we can safely parse it from the UI thread
  $progLog = Join-Path $LogsDir ("oscdimg_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

  # Use cmd.exe for redirection; escape > for PowerShell, and quote paths
  $cmdLine = "/d /c `"`"$OscdimgPath`" $cmdArgsEscaped 2> `"$progLog`"`""

  # Start the process (no handlers; no runspace issues)
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = 'cmd.exe'
  $psi.Arguments = $cmdLine
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow   = $true
  $proc = [System.Diagnostics.Process]::Start($psi)

  # Tail loop: parse "NN% complete" lines from the log
  $lastLen = 0
  try {
    $bar.IsIndeterminate = $true
    while (-not $proc.HasExited) {
      if (Test-Path -LiteralPath $progLog) {
        # read just the last few lines to reduce churn
        $lines = Get-Content -LiteralPath $progLog -Tail 8 -ErrorAction SilentlyContinue
        foreach ($line in $lines) {
          if ($line -match '(\d{1,3})%\s+complete') {
            $p = [int]$matches[1]; if ($p -gt 100) { $p = 100 }
            if ($bar.IsIndeterminate) { $bar.IsIndeterminate = $false }
            $bar.Value = $p
            $pct.Text  = "$p %"
          }
          if ($line) { $txt.Text = $line }
        }
      }
      Invoke-DoEventsWpf
      Start-Sleep -Milliseconds 120
    }

    # One final parse after exit to catch the last line(s)
    if (Test-Path -LiteralPath $progLog) {
      $lines = Get-Content -LiteralPath $progLog -Tail 15 -ErrorAction SilentlyContinue
      foreach ($line in $lines) {
        if ($line -match '(\d{1,3})%\s+complete') {
          $p = [int]$matches[1]; if ($p -gt 100) { $p = 100 }
          $bar.IsIndeterminate = $false
          $bar.Value = $p
          $pct.Text  = "$p %"
        }
        if ($line) { $txt.Text = $line }
      }
    }
  }
  finally {
    try { $proc.WaitForExit() } catch {}
    $win.Close()
  }

  return $proc.ExitCode
}


# ---------- Robocopy + Activity dialog ----------
function New-ActivityDialogWpf([string]$Title='Working...',[string]$Message='Please wait while files are copied...'){
  $x = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Height="140" Width="380"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize" WindowStyle="ToolWindow"
        Background="$panel" Foreground="$text" FontFamily="Segoe UI">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="$Message" TextWrapping="Wrap"/>
    <ProgressBar Grid.Row="1" IsIndeterminate="True" Height="16" Margin="0,12,0,0"/>
  </Grid>
</Window>
"@
  [Windows.Markup.XamlReader]::Parse($x)
}
function Invoke-RobocopyQuiet {
  param([Parameter(Mandatory)][string]$Source,[Parameter(Mandatory)][string]$Destination,[Parameter(Mandatory)][string]$LogPath,[int]$Threads=16)
  if ($Source -notmatch '[\\/]$') { $Source += '\' }
  $core = @($Source,$Destination,'*','/E','/COPY:DAT','/DCOPY:DAT','/R:0','/W:0','/XJ',("/MT:{0}" -f ([math]::Min([math]::Max($Threads,1),128))),'/NFL','/NDL','/NJH','/NJS','/NP')
  $line = ($core | % { if ($_ -match '\s'){'"{0}"' -f $_}else{$_} }) -join ' '
  $cmd  = "/c robocopy $line 1>> `"$LogPath`" 2>>&1"
  $dlg  = New-ActivityDialogWpf -Title 'Copying files...' -Message "Copying files from ISO to working folder.`nLog: $LogPath"
  $dlg.Owner = $Window
  $p = Start-Process -FilePath 'cmd.exe' -ArgumentList $cmd -WindowStyle Hidden -PassThru
  try {
    $dlg.Show()
    while (-not $p.HasExited) { Invoke-DoEventsWpf; Start-Sleep -Milliseconds 120 }
  } finally { $dlg.Close() }
  $p.ExitCode
}

# ---------- Answerfile helper ----------
function Get-AutounattendPath {
  param([string] $UnattendXmlPath,[string] $AnswerFilesDir,[switch] $AllowBrowseFallback = $false)
  if ($UnattendXmlPath) { if (Test-Path -LiteralPath $UnattendXmlPath) { return (Resolve-Path -LiteralPath $UnattendXmlPath).Path }; throw "Specified autounattend path not found: $UnattendXmlPath" }
  $rootXml = Join-Path $BaseDir 'autounattend.xml'
  if (Test-Path -LiteralPath $rootXml) { return (Resolve-Path -LiteralPath $rootXml).Path }
  if (-not $AnswerFilesDir) { $AnswerFilesDir = Join-Path $BaseDir 'answerfiles' }
  if (Test-Path -LiteralPath $AnswerFilesDir) {
    $picked = Show-OpenFile -title 'Select autounattend.xml' -filter 'XML (*.xml)|*.xml' -initialDir $AnswerFilesDir
    if ($picked) { return (Resolve-Path -LiteralPath $picked).Path }
    throw "No autounattend XML selected."
  }
  if ($AllowBrowseFallback) {
    $picked = Show-OpenFile -title 'Select autounattend.xml' -filter 'XML (*.xml)|*.xml' -initialDir $BaseDir
    if ($picked) { return (Resolve-Path -LiteralPath $picked).Path }
  }
  throw "No autounattend XML selected."
}

# ---------- Early startup oscdimg check (disable Start if missing) ----------
try {
  [void](Get-OscdimgPath -UserPath $OscdimgPath)
} catch {
  Show-Error $_.Exception.Message
  $StartBtn.IsEnabled = $false
}

# ---------- UI events ----------
$IsoBrowseBtn.Add_Click({
  $sel = Show-OpenFile -title 'Select Windows ISO' -filter 'ISO (*.iso)|*.iso' -initialDir $BaseDir
  if ($sel) { $IsoPathBox.Text = $sel }
})
$XmlSelectBtn.Add_Click({
  try {
    $dir = if ($AnswerFilesDir) { $AnswerFilesDir } else { Join-Path $BaseDir 'answerfiles' }
    $picked = if (Test-Path -LiteralPath $dir) {
      Show-OpenFile -title 'Select autounattend.xml' -filter 'XML (*.xml)|*.xml' -initialDir $dir
    } else {
      Show-OpenFile -title 'Select autounattend.xml' -filter 'XML (*.xml)|*.xml' -initialDir $BaseDir
    }
    if ($picked) { $XmlPathBox.Text = $picked }
  } catch { Show-Error $_.Exception.Message }
})

# ---------- Worker ----------
$StartBtn.Add_Click({
  $Window.Cursor = 'Wait'
  $StartBtn.IsEnabled = $false
  try {
    $isoSel = if ($IsoPathBox.Text) { $IsoPathBox.Text } elseif ($IsoPath) { $IsoPath } else { '' }
    if (-not $isoSel) { throw "Please select an ISO." }
    if (-not (Test-Path -LiteralPath $isoSel)) { throw "ISO not found: $isoSel" }

    $xmlSel = if ($XmlPathBox.Text) { $XmlPathBox.Text } else {
      try { Get-AutounattendPath -UnattendXmlPath $UnattendXmlPath -AnswerFilesDir $AnswerFilesDir -AllowBrowseFallback:$false }
      catch { throw "Please select an autounattend XML." }
    }
    if (-not (Test-Path -LiteralPath $xmlSel)) { throw "autounattend XML not found: $xmlSel" }

    $includeOem = if ($OemCheck.IsChecked) { $true } elseif ($SkipOEM) { $false } elseif ($IncludeOEM) { $true } else { $false }

    $isoName = [IO.Path]::GetFileNameWithoutExtension($isoSel)
    $workDir = Join-Path $BaseDir $isoName
    Show-Ok  "Selected ISO: $isoName"

    $skip = $false
    if (Test-Path -LiteralPath $workDir) {
      $res = [System.Windows.MessageBox]::Show("Folder '$workDir' exists. Reuse its contents?","Reuse extracted files?",'YesNo','Question')
      if ($res -eq 'Yes') { Show-Info "Reusing existing folder contents."; $skip=$true }
      else { Show-Warn "Deleting existing folder..."; Remove-Item -LiteralPath $workDir -Recurse -Force; Show-Ok "Existing folder deleted." }
    }

    # Mount & copy
    $img = $null
    try {
      if (-not $skip) {
        $img = Mount-DiskImage -ImagePath $isoSel -PassThru
        $drive = ($img | Get-Volume).DriveLetter + ':'
        Show-Ok "Mounted ISO at $drive\"
        [IO.Directory]::CreateDirectory($workDir) | Out-Null

        $log = Join-Path $LogsDir ("robocopy_$(${isoName})_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
        Show-Info "Copying files from ISO (quiet, multi-threaded)..."
        $rc = Invoke-RobocopyQuiet -Source "$drive\" -Destination $workDir -LogPath $log -Threads 16
        if ($rc -gt 7) { throw "Robocopy failed (exit code $rc). See log: $log" }
        Show-Ok "Files copied. Full robocopy log saved to:`n$log"
      }
    } finally {
      if ($img) { try { Dismount-DiskImage -ImagePath $isoSel | Out-Null } catch {} }
    }

    # autounattend.xml
    Show-Ok  "Selected XML: $xmlSel"
    Show-Info "Copying autounattend.xml to $workDir"
    Copy-Item -LiteralPath $xmlSel -Destination (Join-Path $workDir 'autounattend.xml') -Force
    try { [void][xml](Get-Content -LiteralPath $xmlSel -Raw) } catch { Show-Warn "The selected file is not well-formed XML: $xmlSel" }

    # $OEM$ (optional)
    if ($includeOem) {
      $oemDefault = Join-Path $BaseDir '$OEM$'
      $oemSrc = if (Test-Path -LiteralPath $oemDefault) { $oemDefault } else { $null }
      if (-not $oemSrc) { Show-Warn "Default `$OEM$ folder not found at: $oemDefault (skipping in WPF build)" }
      if ($oemSrc) {
        $dest = Join-Path $workDir '$OEM$'
        $null = New-Item -ItemType Directory -Force -Path $dest
        $contents = Get-ChildItem -LiteralPath $oemSrc -Force -ErrorAction SilentlyContinue
        if ($contents) { Get-ChildItem -LiteralPath $oemSrc -Force | Copy-Item -Destination $dest -Recurse -Force; Show-Ok "`$OEM$ folder included." }
        else { Show-Warn "`$OEM$ folder is empty; nothing to copy." }
      }
    } else { Show-Info "Skipping `$OEM$ folder." }

    # Build ISO
    $oscd   = Get-OscdimgPath -UserPath $OscdimgPath
    $outIso = Join-Path $BaseDir ($isoName + '.custom.iso')
    $bootEtfs = Join-Path $workDir 'boot\etfsboot.com'
    $bootEfi  = Join-Path $workDir 'efi\microsoft\boot\efisys.bin'
    $bootData = '2#p0,e,b"{0}"#pEF,e,b"{1}"' -f $bootEtfs, $bootEfi

    Show-Info "Creating custom ISO..."
    $exit = Invoke-OscdimgWithProgress -OscdimgPath $oscd -WorkDir $workDir -OutIso $outIso -BootData $bootData
    if ($exit -ne 0 -or -not (Test-Path -LiteralPath $outIso)) { throw "Failed to create the custom ISO (oscdimg exit $exit)." }

    Show-Ok "Custom ISO generated: $outIso"
    $sizeMB = [math]::Round(((Get-Item -LiteralPath $outIso).Length/1MB),2)
    $sizeGB = [math]::Round(((Get-Item -LiteralPath $outIso).Length/1GB),2)
    Show-Ok "Custom ISO size: $sizeMB MB (~$sizeGB GB)"

    # Cleanup
    $ans = [System.Windows.MessageBox]::Show("Delete the extraction folder '$workDir'?","Cleanup",'YesNo','Question')
    if ($ans -eq 'Yes') { Remove-Item -LiteralPath $workDir -Recurse -Force; Show-Ok "Extraction folder deleted." }
    else { Show-Info "Kept extraction folder: $workDir" }

    Show-Ok "Done."
  } catch {
    Show-Error $_.Exception.Message
  } finally {
    $Window.Cursor = 'Arrow'
    $StartBtn.IsEnabled = $true
  }
})

# ---------- Show WPF window (does not auto-close) ----------
$Window.ShowDialog() | Out-Null
