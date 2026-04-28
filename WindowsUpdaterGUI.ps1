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

        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Available Windows Updates" FontSize="20" FontWeight="SemiBold" Foreground="$TextFg" VerticalAlignment="Center" />
        </StackPanel>

        <Grid Grid.Row="1" Margin="0,10,0,10">
            <ListView x:Name="UpdateList" Background="$ControlBg" Foreground="$TextFg" BorderBrush="$Border" BorderThickness="1">
                <ListView.Resources>
                    <!-- Style to make GridView headers match the Dark Theme -->
                    <Style TargetType="GridViewColumnHeader">
                        <Setter Property="Background" Value="$ControlBg"/>
                        <Setter Property="Foreground" Value="$TextFg"/>
                        <Setter Property="FontWeight" Value="SemiBold"/>
                        <Setter Property="Padding" Value="5,5,5,5"/>
                        <Setter Property="HorizontalContentAlignment" Value="Left"/>
                    </Style>
                </ListView.Resources>
                <ListView.ItemContainerStyle>
                    <Style TargetType="ListViewItem">
                        <Setter Property="Foreground" Value="{Binding Color}"/>
                    </Style>
                </ListView.ItemContainerStyle>
                <ListView.View>
                    <GridView>
                        <GridViewColumn Width="40">
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <CheckBox IsChecked="{Binding IsChecked, Mode=TwoWay}" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn Header="Title" DisplayMemberBinding="{Binding Title}" Width="350"/>
                        <GridViewColumn Header="KB Article" DisplayMemberBinding="{Binding KBArticle}" Width="100"/>
                        <GridViewColumn Header="Type" DisplayMemberBinding="{Binding TypeStr}" Width="90"/>
                        <GridViewColumn Header="Release Date" DisplayMemberBinding="{Binding ReleaseDate}" Width="100"/>
                        <GridViewColumn Header="Size" DisplayMemberBinding="{Binding SizeStr}" Width="80"/>
                    </GridView>
                </ListView.View>
            </ListView>
        </Grid>

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
        $UpdateList.ItemsSource = $null
        $ScanBtn.IsEnabled = $false
        $InstallBtn.IsEnabled = $false

        $StatusText.Text = "Scanning for Windows updates... Please wait."
        Do-Events
        Write-Host "`nScanning for available Windows Updates..." -ForegroundColor Cyan

        # Execute get updates
        $updates = Get-WindowsUpdate

        if (-not $updates) {
            $StatusText.Text = "Your device is up to date. No updates found."
            Write-Host "No updates found." -ForegroundColor DarkGray
        }
        else {
            $UpdateData = New-Object System.Collections.ArrayList
            foreach ($u in $updates) {
                $kb = if ($u.KBArticleIDs -and $u.KBArticleIDs.Count -gt 0) {
                    "KB$($u.KBArticleIDs[0])"
                }
                elseif ($u.KBArticleID) {
                    "KB$($u.KBArticleID -replace '^KB', '')"
                }
                else {
                    "N/A"
                }

                $title = $u.Title

                # Microsoft often embeds the KB article directly in the update title.
                # Let's strip it out so it doesn't appear redundantly in the main body column.
                if ($kb -ne "N/A") {
                    $escapedKb = [regex]::Escape($kb)
                    $title = $title -replace "\(\s*$escapedKb\s*\)", ""
                    $title = $title -replace $escapedKb, ""
                    $title = $title.Trim()
                }

                $cat = $u.Categories -join ', '

                # Calculate Size
                $sizeStr = ""
                if ($null -ne $u.Size -and $u.Size -gt 0) {
                    if ($u.Size -ge 1GB) {
                        $sizeVal = [math]::Round($u.Size / 1GB, 2)
                        $sizeStr = "${sizeVal} GB"
                    }
                    else {
                        $sizeVal = [math]::Round($u.Size / 1MB, 2)
                        $sizeStr = "${sizeVal} MB"
                    }
                }

                # Get Release Date
                $dateStr = ""
                if ($null -ne $u.LastDeploymentChangeTime) {
                    $dateStr = "{0:yyyy-MM-dd}" -f $u.LastDeploymentChangeTime
                }

                # Determine colors based on categories
                if ($cat -match 'Upgrade' -or $title -match 'Upgrade') {
                    $uiColor = "Red"
                    $consoleColor = "Red"
                    $prefix = "[FULL VERSION UPDATE]"
                    $typeStr = "Full Version"
                }
                elseif ($cat -match 'Feature' -or $title -match 'Feature Update' -or $title -match 'version \d{2}[Hh]\d') {
                    $uiColor = "Orange"
                    $consoleColor = "DarkYellow"
                    $prefix = "[FEATURE UPDATE]"
                    $typeStr = "Feature"
                }
                elseif ($cat -match 'Security' -or $title -match 'Security') {
                    $uiColor = "LightGreen"
                    $consoleColor = "Green"
                    $prefix = "[SECURITY UPDATE]"
                    $typeStr = "Security"
                }
                else {
                    $uiColor = $TextFg
                    $consoleColor = "White"
                    $prefix = "[OTHER UPDATE]"
                    $typeStr = "Other"
                }

                # Keep the original formatting for the terminal logs
                $consoleInfo = "$title"
                if ($sizeStr) { $consoleInfo += " - $sizeStr" }
                if ($dateStr) { $consoleInfo += " - Released: $dateStr" }

                # Display in terminal
                if ($kb) {
                    Write-Host "$prefix $kb - $consoleInfo" -ForegroundColor $consoleColor
                }
                else {
                    Write-Host "$prefix $consoleInfo" -ForegroundColor $consoleColor
                }

                # Default all items to unchecked.
                $item = [PSCustomObject]@{
                    IsChecked    = $false
                    Title        = $title
                    KBArticle    = $kb
                    TypeStr      = $typeStr
                    ReleaseDate  = $dateStr
                    SizeStr      = $sizeStr
                    Color        = $uiColor
                    RawObject    = $u
                }
                $UpdateData.Add($item) | Out-Null
            }
            $UpdateList.ItemsSource = $UpdateData
            $StatusText.Text = "Scan complete. $($UpdateData.Count) update(s) found. Select the ones you wish to install."
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
        foreach ($item in $UpdateList.ItemsSource) {
            if ($item.IsChecked -eq $true) {
                $updateObject = $item.RawObject
                $updateTitle = $item.Title

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
