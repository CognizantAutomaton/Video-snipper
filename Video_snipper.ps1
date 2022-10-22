using namespace System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore, PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

$ErrorActionPreference = "Stop"

function New-Window {
    param(
        [Parameter(Mandatory=$true)]
        [Xml]$XAML
    )

    $XAML.Window.RemoveAttribute('x:Class')
    $XAML.Window.RemoveAttribute('mc:Ignorable')

    $WpfNs = New-Object -TypeName Xml.XmlNamespaceManager -ArgumentList $XAML.NameTable
    $WpfNs.AddNamespace('x', $XAML.DocumentElement.x)
    $WpfNs.AddNamespace('d', $XAML.DocumentElement.d)
    $WpfNs.AddNamespace('mc', $XAML.DocumentElement.mc)

    # Read XAML markup
    try {
        $Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAML))
    } catch {
        Write-Host $_ -ForegroundColor Red
        Exit
    }

    # provide a reference to each GUI form control
    $Gui = @{}
    $XAML.SelectNodes('//*[@x:Name]', $WpfNs) | ForEach-Object {
        $Gui.Add($_.Name, $Window.FindName($_.Name))
    }

    return @{ Window = $Window; Gui = $Gui }
}

[Xml]$WpfXml = @"
<Window x:Class="ffmpeg_script_gui.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:ffmpeg_script_gui"
        mc:Ignorable="d"
        Title="Video snipper" Height="280" Width="500"
        ResizeMode="CanMinimize"
        WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style x:Key="DarkGrey" TargetType="FrameworkElement">
            <Setter Property="Control.Background" Value="#111111" />
            <Setter Property="Control.BorderBrush" Value="Black" />
        </Style>
        <Style x:Key="Dark" TargetType="FrameworkElement">
            <Setter Property="Control.Background" Value="Black" />
            <Setter Property="Control.BorderBrush" Value="Black" />
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Control.Background" Value="#333333" />
            <Setter Property="Control.Foreground" Value="#425595" />
        </Style>
        <Style TargetType="Label" BasedOn="{StaticResource DarkGrey}">
            <Setter Property="Control.Foreground" Value="#AAAAAA" />
        </Style>
        <Style TargetType="TextBlock" BasedOn="{StaticResource DarkGrey}">
            <Setter Property="Control.Foreground" Value="#AAAAAA" />
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Control.Background" Value="#111111" />
            <Setter Property="Control.Foreground" Value="#AAAAAA" />
            <Setter Property="Control.BorderBrush" Value="#333333" />
        </Style>
        <Style TargetType="ListBox" BasedOn="{StaticResource Dark}">
        </Style>
        <Style TargetType="GroupBox" BasedOn="{StaticResource Dark}">
            <Setter Property="Control.Foreground" Value="#AAAAAA" />
        </Style>
        <Style TargetType="Button">
            <Setter Property="Control.Background" Value="#222222" />
            <Setter Property="Control.Foreground" Value="#AAAAAA" />
        </Style>
        <Style TargetType="{x:Type RadioButton}">
            <Setter Property="Control.Foreground" Value="#AAAAAA" />
        </Style>
        <Style TargetType="Menu" BasedOn="{StaticResource DarkGrey}"/>
        <Style TargetType="MenuItem">
            <Setter Property="Control.Foreground" Value="#AAAAAA" />
            <Setter Property="Control.Background" Value="#111111" />
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Control.Background" Value="Yellow" />
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Control.Background" Value="Yellow" />
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="Window" BasedOn="{StaticResource DarkGrey}">
        </Style>
        <ControlTemplate x:Key="CustomButton" TargetType="{x:Type Button}">
            <Border Background="#222222" BorderBrush="Gray" BorderThickness="1" x:Name="Border">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
            </Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter TargetName="Border" Property="Background" Value="#777777" />
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#000055" TargetName="Border" />
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>
        <Style TargetType="{x:Type DataGridRow}">
            <Style.Setters>
                <Setter Property="Background" Value="{Binding Path=Code}"></Setter>
            </Style.Setters>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="30"/>
            <RowDefinition/>
            <RowDefinition Height="80"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
        </Grid.RowDefinitions>
        <Grid
            Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition/>
                <ColumnDefinition Width="300"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock
                Grid.Column="0"
                Text="File path: "
                HorizontalAlignment="Right"
                VerticalAlignment="Center"/>
            <TextBox
                Grid.Column="1"
                x:Name="txtSourceFilePath"
                HorizontalAlignment="Stretch"
                VerticalAlignment="Center"/>
            <Button
                Grid.Column="2"
                x:Name="btnBrowse"
                Content="Browse"
                HorizontalAlignment="Left"
                VerticalAlignment="Center"
                Template="{StaticResource CustomButton}"/>
        </Grid>
        <GroupBox
            Grid.Row="1"
            Header="Video clipping"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Stretch">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition/>
                    <RowDefinition/>
                </Grid.RowDefinitions>
                <Grid
                    Grid.Row="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <TextBlock
                        Grid.Column="0"
                        Text="From: "
                        HorizontalAlignment="Right"
                        VerticalAlignment="Center"/>
                    <TextBox
                        Grid.Column="1"
                        x:Name="txtFrom"
                        Text="00:00:00.000"
                        HorizontalAlignment="Left"
                        VerticalAlignment="Center"
                        Width="100"/>
                </Grid>
                <Grid
                    Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <TextBlock
                        Grid.Column="0"
                        Text="To: "
                        HorizontalAlignment="Right"
                        VerticalAlignment="Center"/>
                    <TextBox
                        Grid.Column="1"
                        x:Name="txtTo"
                        Text="23:59:59.999"
                        HorizontalAlignment="Left"
                        VerticalAlignment="Center"
                        Width="100"/>
                </Grid>
            </Grid>
        </GroupBox>
        <GroupBox
            Grid.Row="2"
            Header="Output format"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Stretch">
            <Grid
                HorizontalAlignment="Center">
                <Grid.RowDefinitions>
                    <RowDefinition/>
                    <RowDefinition/>
                </Grid.RowDefinitions>
                <Grid
                    Grid.Row="0"
                    HorizontalAlignment="Center">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <RadioButton
                        Grid.Column="0"
                        x:Name="rdoVideo"
                        GroupName="grpFormat"
                        Content="Video"
                        Margin="0,0,20,0"
                        IsChecked="True"/>
                    <RadioButton
                        Grid.Column="1"
                        x:Name="rdoGIF"
                        GroupName="grpFormat"
                        Content="GIF"
                        Margin="20,0,0,0"/>
                </Grid>
                <Grid
                    Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Label
                        x:Name="lblEncoding"
                        Grid.Column="0"
                        Content="Encoding method: "/>
                    <ComboBox
                        Grid.Column="1"
                        IsEditable="False"
                        x:Name="dropdownEncodingMethod"
                        Background="#111111"
                        VerticalAlignment="Top"
                        Height="25"
                        Width="200">
                        <ComboBox.Resources>
                            <SolidColorBrush x:Key="{x:Static SystemColors.WindowBrushKey}" Color="DarkBlue" />
                        </ComboBox.Resources>
                        <ComboBox.ItemContainerStyle>
                            <Style TargetType="ComboBoxItem">
                                <Setter Property="Background" Value="#111111"/>
                                <Setter Property="Foreground" Value="#AAAAAA"/>
                                <Setter Property="BorderBrush" Value="Gray"/>
                            </Style>
                        </ComboBox.ItemContainerStyle>
                        <ComboBoxItem x:Name="itemNearestIframe" Content="Nearest i-frame (fast)" IsSelected="True"/>
                        <ComboBoxItem x:Name="itemReEncode" Content="Re-encode (slower, more accurate)"/>
                    </ComboBox>
                </Grid>
            </Grid>
        </GroupBox>
        <Button
            Grid.Row="3"
            x:Name="btnStart"
            Content="Start"
            HorizontalAlignment="Center"
            VerticalAlignment="Stretch"
            Width="60"
            Template="{StaticResource CustomButton}"/>
        <ProgressBar
            Grid.Row="4"
            x:Name="progressBar"/>
    </Grid>
</Window>
"@

$Sync = [Hashtable]::Synchronized((New-Window -XAML $WpfXml))

[Hashtable]$Commands = @{
    "TOTAL_LENGTH" = 'ffprobe -v quiet -print_format compact=print_section=0:nokey=1:escape=csv -show_entries format=duration "{0}"'
    "FAST_NEAREST_IFRAME" = 'ffmpeg -i "{0}" -ss {1} -to {2} -c:v copy -c:a copy "{3}"'
    "ACCURATE_REENCODE" = 'ffmpeg -i "{0}" -ss {1} -to {2} "{3}"'
    "SNIP_GIF" = 'ffmpeg -ss {0} -t {1} -i "{2}" "{3}"' #0=from_time, 1=num_secs, 2=input_vid, 3=output_gif
}
[double]$global:TotalLength = [TimeSpan]::Parse("23:59:59.999").TotalSeconds

$Sync.Gui.btnBrowse.add_Click({
    $ofd = New-Object OpenFileDialog

    if ($ofd.ShowDialog() -eq "OK") {
        $Sync.Gui.txtSourceFilePath.Text = $ofd.FileName

        $global:TotalLength = Invoke-Expression ($Commands.TOTAL_LENGTH -f $ofd.FileName)
        [TimeSpan]$span = [TimeSpan]::FromSeconds($global:TotalLength)
        $Sync.Gui.txtTo.Text = "{0:hh\:mm\:ss\.fff}" -f $span
    }
})

function Get-TimeFormat([string]$InputObject) {
    $lst = New-Object System.Collections.Generic.List[object]
    $lst.AddRange([object[]]($InputObject -split ':'))

    for ([int]$n = 0; $n -lt $lst.Count; $n++) {
        if ($n -eq $lst.Count - 1) {
            $lst[$n] = [string]::Format("{0:00.000}", ([double]$lst[$n]))
        } else {
            $lst[$n] = "{0:00}" -f ([int]$lst[$n])
        }
    }

    $joined = $lst -join ':'

    if ($lst.Count -eq 1) {
        $joined = "00:00:$joined"
    } elseif ($lst.Count -eq 2) {
        $joined = "00:$joined"
    }

    [TimeSpan]$result = [TimeSpan]::Parse($joined)

    return "{0:hh\:mm\:ss\.fff}" -f $result
}

$Sync.Gui.txtFrom.add_LostFocus({
    if ($this.Text.Length -eq 0) {
        $this.Text = "0"
    }

    try {
        $this.Text = Get-TimeFormat $this.Text

        if ([TimeSpan]::Parse($this.Text).TotalSeconds -gt [TimeSpan]::Parse($Sync.Gui.txtTo.Text).TotalSeconds) {
            [TimeSpan]$span = [TimeSpan]::FromSeconds([TimeSpan]::Parse($Sync.Gui.txtTo.Text).TotalSeconds - 1)
            $Sync.Gui.txtFrom.Text = "{0:hh\:mm\:ss\.fff}" -f $span
        }
    } catch {
        $this.Text = "00:00:00.000"
        [System.Windows.Forms.MessageBox]::Show("Input must be hh:mm:ss format.")
    }
})

$Sync.Gui.txtTo.add_LostFocus({
    if ($this.Text.Length -eq 0) {
        $this.Text = "0"
    }

    try {
        $this.Text = Get-TimeFormat $this.Text

        if ([TimeSpan]::Parse($this.Text).TotalSeconds -gt $global:TotalLength) {
            [TimeSpan]$span = [TimeSpan]::FromSeconds($global:TotalLength)
            $Sync.Gui.txtTo.Text = "{0:hh\:mm\:ss\.fff}" -f $span
        }
    } catch {
        [TimeSpan]$span = [TimeSpan]::FromSeconds($global:TotalLength)
        $Sync.Gui.txtTo.Text = "{0:hh\:mm\:ss\.fff}" -f $span
        [System.Windows.Forms.MessageBox]::Show("Input must be hh:mm:ss format.")
    }
})

$Sync.Gui.btnStart.add_Click({
    [string]$src = $Sync.Gui.txtSourceFilePath.Text
    [string]$dest = [string]::Empty
    [double]$start = [TimeSpan]::Parse($Sync.Gui.txtFrom.Text).TotalSeconds
    [double]$end = [TimeSpan]::Parse($Sync.Gui.txtTo.Text).TotalSeconds
    [SaveFileDialog]$sfd = $null
    [string]$SoundFile = "C:\Windows\Media\Alarm05.wav"
    $Sync.Gui.progressBar.Value = 0

    if (Test-Path -LiteralPath $src) {
        $item = Get-Item -LiteralPath $src

        switch ($true) {
            $Sync.Gui.rdoVideo.IsChecked {
                $sfd = New-Object SaveFileDialog -Property @{
                    FileName = Join-Path $item.DirectoryName "$($item.BaseName)_edit$($item.Extension)"
                    Filter = "$($item.Extension.Trim(".").ToUpper()) file|*$($item.Extension)|All files|*.*"
                }
            }

            $Sync.Gui.rdoGIF.IsChecked {
                $sfd = New-Object SaveFileDialog -Property @{
                    FileName = Join-Path $item.DirectoryName "$($item.BaseName).gif"
                    Filter = "Animated GIF|*.gif"
                }
            }
        }

        if ($sfd.ShowDialog() -eq "OK") {
            $dest = $sfd.FileName

            switch ($true) {
                $Sync.Gui.rdoVideo.IsChecked {
                    if ($Sync.Gui.dropdownEncodingMethod.SelectedItem.Name -eq "itemNearestIframe") {
                        # this method seeks to the nearest i-frame and copies ... fast
                        Invoke-Expression ($Commands.FAST_NEAREST_IFRAME -f $src, $start, $end, $dest)
                    } elseif ($Sync.Gui.dropdownEncodingMethod.SelectedItem.Name -eq "itemReEncode") {
                        # this method re-encodes the file ... slow, but more accurate adherence to a time specification
                        Invoke-Expression ($Commands.ACCURATE_REENCODE -f $src, $start, $end, $dest)
                    }
                }

                $Sync.Gui.rdoGIF.IsChecked {
                    $end -= $start
                    $MyCommand = ($Commands.SNIP_GIF -f $start, $end, $src, $dest)
                    Invoke-Expression ($MyCommand)
                }
            }

            if (Test-Path -LiteralPath $SoundFile) {
                $NotifySound = New-Object System.Media.SoundPlayer
                $NotifySound.SoundLocation = $SoundFile
                $NotifySound.PlaySync()
            }

            $Sync.Gui.progressBar.Value = 100
            Write-Host "Finished!"
        } else {
            Write-Host "Cancelled"
        }
    } else {
        Write-Host "Source file path is not valid"
    }
})

$Sync.Gui.rdoVideo.add_Checked({
    $Sync.Gui.lblEncoding.Visibility = "Visible"
    $Sync.Gui.dropdownEncodingMethod.Visibility = "Visible"
})

$Sync.Gui.rdoGIF.add_Checked({
    $Sync.Gui.lblEncoding.Visibility = "Hidden"
    $Sync.Gui.dropdownEncodingMethod.Visibility = "Hidden"
})

[void]$Sync.Window.ShowDialog()
