#Requires -RunAsAdministrator

# Force loading of necessary .NET assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

Write-Host "Checking for PSWindowsUpdate module..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Host "Installing PSWindowsUpdate module..." -ForegroundColor DarkYellow
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
    Install-Module PSWindowsUpdate -Force -AllowClobber -Scope AllUsers
}
Import-Module PSWindowsUpdate

Write-Host "Adding Microsoft Update Service Manager..." -ForegroundColor Cyan
Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue | Out-Null

# 0. Add C# type for DWM API to enable Dark Mode on the window title bar
if (-not ('DwmApi' -as [type])) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class DwmApi {
        [DllImport("dwmapi.dll", PreserveSig = true)]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    }
"@
}

# 1. Define Static Dark Theme Color Palette
$WindowBg = "#1E1E1E"
$ControlBg = "#2D2D2D"
$TextFg = "#F1F1F1"
$Border = "#3F3F46"
$Accent = "#007ACC"
$AccentHover = "#0097FF"
$DropBg = "#004C7A"

# 2. Build the XAML UI integrating the Master Theme
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows Update Manager" Height="550" Width="850"
        Background="$WindowBg" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <!-- Base Button Style -->
        <Style TargetType="{x:Type Button}">
            <Setter Property="Background" Value="$ControlBg" />
            <Setter Property="Foreground" Value="$TextFg" />
            <Setter Property="BorderBrush" Value="$Border" />
            <Setter Property="Padding" Value="8" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="$Border" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Accent Button Style (for Convert/Install) -->
        <Style x:Key="AccentButton" TargetType="{x:Type Button}" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="$Accent" />
            <Setter Property="Foreground" Value="#FFFFFF" />
            <Setter Property="BorderBrush" Value="$Accent" />
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="$AccentHover" />
                    <Setter Property="BorderBrush" Value="$AccentHover" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- ScrollViewer Style -->
        <Style TargetType="{x:Type ScrollViewer}">
            <Setter Property="Margin" Value="0,10,0,10"/>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="Available Windows Updates" FontSize="20" FontWeight="SemiBold" Foreground="$TextFg" />

        <Border Grid.Row="1" BorderBrush="$Border" BorderThickness="1" Background="$ControlBg" Margin="0,10,0,10">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="UpdateList" Margin="5" />
            </ScrollViewer>
        </Border>

        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <TextBlock x:Name="StatusText" Grid.Column="0" Foreground="$TextFg" VerticalAlignment="Center" Text="Ready to scan." TextWrapping="Wrap" Margin="0,0,15,0" />
            <Button x:Name="ScanBtn" Grid.Column="1" Content="Scan for Updates" Width="140" Margin="0,0,10,0" />
            <Button x:Name="InstallBtn" Grid.Column="2" Content="Install Selected" Width="140" Style="{StaticResource AccentButton}" IsEnabled="False" />
        </Grid>
    </Grid>
</Window>
"@

# Load the Window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Connect UI Elements
$UpdateList = $Window.FindName("UpdateList")
$ScanBtn = $Window.FindName("ScanBtn")
$InstallBtn = $Window.FindName("InstallBtn")
$StatusText = $Window.FindName("StatusText")

# 3. Apply Dark Mode to Title Bar
$hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($Window)).EnsureHandle()
$trueVal = 1
[DwmApi]::DwmSetWindowAttribute($hwnd, 20, [ref]$trueVal, 4) | Out-Null # Windows 10 1903+ and Windows 11
[DwmApi]::DwmSetWindowAttribute($hwnd, 19, [ref]$trueVal, 4) | Out-Null # Windows 10 1809

# Helper function to keep the UI from freezing during PS commands
function Do-Events {
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    $dispatcher = [System.Windows.Threading.Dispatcher]::CurrentDispatcher
    $action = [Action] { $frame.Continue = $false }
    $dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, $action) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

# --- EVENTS ---

$ScanBtn.Add_Click({
        $StatusText.Text = "Scanning for updates... Please wait."
        $UpdateList.Children.Clear()
        $ScanBtn.IsEnabled = $false
        $InstallBtn.IsEnabled = $false
        Do-Events

        Write-Host "`nScanning for available Windows Updates..." -ForegroundColor Cyan

        # Execute get updates
        $updates = Get-WindowsUpdate

        if (-not $updates) {
            $StatusText.Text = "Your device is up to date. No updates found."
            Write-Host "No updates found." -ForegroundColor DarkGray
        }
        else {
            foreach ($u in $updates) {
                $kb = $u.KBArticleID
                if ($kb -is [array]) { $kb = $kb[0] }

                $title = $u.Title
                $cat = $u.Categories -join ', '

                # Calculate Size
                $sizeStr = ""
                if ($null -ne $u.Size -and $u.Size -gt 0) {
                    if ($u.Size -ge 1GB) {
                        $sizeVal = [math]::Round($u.Size / 1GB, 2)
                        $sizeStr = " - ${sizeVal} GB"
                    }
                    else {
                        $sizeVal = [math]::Round($u.Size / 1MB, 2)
                        $sizeStr = " - ${sizeVal} MB"
                    }
                }

                # Get Release Date
                $dateStr = ""
                if ($null -ne $u.LastDeploymentChangeTime) {
                    $dateStr = " - Released: $("{0:yyyy-MM-dd}" -f $u.LastDeploymentChangeTime)"
                }
                $title = "$title$sizeStr$dateStr"

                # Determine colors based on categories
                if ($cat -match 'Upgrade' -or $title -match 'Upgrade') {
                    $uiColor = "Red"
                    $consoleColor = "Red"
                    $prefix = "[FULL VERSION UPDATE]"
                    $title = "$title (Full Version)"
                }
                elseif ($cat -match 'Feature' -or $title -match 'Feature Update' -or $title -match 'version \d{2}[Hh]\d') {
                    $uiColor = "Orange"
                    $consoleColor = "DarkYellow"
                    $prefix = "[FEATURE UPDATE]"
                    $title = "$title (Feature)"
                }
                elseif ($cat -match 'Security' -or $title -match 'Security') {
                    $uiColor = "LightGreen"
                    $consoleColor = "Green"
                    $prefix = "[SECURITY UPDATE]"
                    $title = "$title (Security)"
                }
                else {
                    $uiColor = $TextFg
                    $consoleColor = "White"
                    $prefix = "[OTHER UPDATE]"
                }

                # Display in terminal
                if ($kb) {
                    Write-Host "$prefix $kb - $title" -ForegroundColor $consoleColor
                }
                else {
                    Write-Host "$prefix $title" -ForegroundColor $consoleColor
                }

                # Create Checkbox and TextBlock for the UI
                $tb = New-Object System.Windows.Controls.TextBlock
                if ($kb) {
                    $tb.Text = "[$kb] $title"
                }
                else {
                    $tb.Text = $title
                }
                $tb.Foreground = $uiColor
                $tb.TextWrapping = "Wrap"

                $cb = New-Object System.Windows.Controls.CheckBox
                $cb.Content = $tb
                $cb.Tag = $u # Store the entire update object, not just the KB ID
                $cb.Margin = "0,5,0,5"

                $UpdateList.Children.Add($cb) | Out-Null
            }
            $StatusText.Text = "Scan complete. $($UpdateList.Children.Count) update(s) found. Select the ones you wish to install."
            $InstallBtn.IsEnabled = $true
        }

        $ScanBtn.IsEnabled = $true
    })

$InstallBtn.Add_Click({
        $StatusText.Text = "Installing selected updates... Check terminal for progress."
        $ScanBtn.IsEnabled = $false
        $InstallBtn.IsEnabled = $false
        Do-Events

        Write-Host "`n--- Beginning Update Installation ---" -ForegroundColor Cyan
        foreach ($cb in $UpdateList.Children) {
            if ($cb.IsChecked -eq $true) {
                $updateObject = $cb.Tag
                $updateTitle = $cb.Content.Text

                if ($updateObject) {
                    Write-Host "Installing $updateTitle ..." -ForegroundColor Yellow
                    Do-Events
                    # Pass the entire update object to the cmdlet for robust installation
                    $updateObject | Install-WindowsUpdate -AcceptAll -Confirm:$false
                    Write-Host "Finished processing '$($updateObject.Title)'." -ForegroundColor Green
                }
            }
        }

        $StatusText.Text = "Installation cycle complete. You may want to reboot."
        Write-Host "--- Installation Cycle Complete ---" -ForegroundColor Cyan

        $ScanBtn.IsEnabled = $true
        $InstallBtn.IsEnabled = $true
    })

# Show the application
$Window.ShowDialog() | Out-Null
