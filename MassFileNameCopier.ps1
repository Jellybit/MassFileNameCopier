Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Global variables
$Global:DroppedFiles = @()
$Global:PathDelimiter = "|||" # Delimiter for joining multiple file paths; change if needed

# Function to show message box
function Show-MessageBox {
    param (
        [string]$Message,
        [string]$Title = "Mass File Name Copier",
        [string]$Type = "Info"
    )
    $image = switch ($Type) {
        "Error" { [System.Windows.MessageBoxImage]::Error }
        "Warning" { [System.Windows.MessageBoxImage]::Warning }
        "Question" { [System.Windows.MessageBoxImage]::Question }
        default { [System.Windows.MessageBoxImage]::Information }
    }
    $button = if ($Type -eq "Question") { [System.Windows.MessageBoxButton]::YesNo } else { [System.Windows.MessageBoxButton]::OK }
    [System.Windows.MessageBox]::Show($Global:Window, $Message, $Title, $button, $image)
}

# Function to show OpenFileDialog for multiple files
function Show-OpenFileDialog {
    param (
        [string]$InitialPathHint
    )
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "All Files (*.*)|*.*"
    $dialog.Multiselect = $true
    if ($InitialPathHint -and (Test-Path $InitialPathHint)) {
        $dialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($InitialPathHint)
    }
    if ($dialog.ShowDialog() -eq "OK") { return $dialog.FileNames }
    return $null
}

# Function to show FolderBrowserDialog
function Show-FolderBrowserDialog {
    param (
        [string]$InitialPathHint
    )
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select input folder"
    $dialog.ShowNewFolderButton = $true
    if ($InitialPathHint -and (Test-Path $InitialPathHint)) {
        $dialog.SelectedPath = $InitialPathHint
    }
    if ($dialog.ShowDialog() -eq "OK") { return $dialog.SelectedPath }
    return $null
}

# XAML for the GUI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        Title="Mass File Name Copier" Height="450" Width="600"
        WindowStartupLocation="CenterScreen"
        AllowsTransparency="True" WindowStyle="None" Background="Transparent"
        Name="MainWindow" ResizeMode="CanResize">
    <shell:WindowChrome.WindowChrome>
        <shell:WindowChrome ResizeBorderThickness="8"
                            CaptionHeight="30"
                            CornerRadius="0"
                            GlassFrameThickness="0"
                            UseAeroCaptionButtons="False"/>
    </shell:WindowChrome.WindowChrome>
    <Border BorderBrush="#FF007ACC" BorderThickness="1" Background="#FF2D2D30" CornerRadius="0">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border x:Name="CustomTitleBar" Grid.Row="0" Background="#FF252526" Height="30" CornerRadius="0,0,0,0">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Text="Mass File Name Copier" VerticalAlignment="Center" HorizontalAlignment="Left" Margin="10,0,0,0" Foreground="White" FontWeight="SemiBold"/>
                    <Button x:Name="CloseButton" Grid.Column="1" Content="✕" Width="40" Height="30"
                            shell:WindowChrome.IsHitTestVisibleInChrome="True">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Background" Value="#FF252526"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter Property="BorderThickness" Value="0"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="Button">
                                            <Border Background="{TemplateBinding Background}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Border>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                                <Style.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#FFE81123"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    </Button>
                </Grid>
            </Border>
            <Grid Grid.Row="1" Margin="10">
                <Grid.Resources>
                    <Style x:Key="BrowseButtonStyle" TargetType="Button">
                        <Setter Property="Background" Value="#FF3E3E42"/>
                        <Setter Property="Foreground" Value="White"/>
                        <Setter Property="BorderBrush" Value="#FF007ACC"/>
                        <Setter Property="BorderThickness" Value="1"/>
                        <Setter Property="Padding" Value="8,3"/>
                        <Setter Property="Margin" Value="5,5,0,5"/>
                        <Setter Property="VerticalContentAlignment" Value="Center"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="Button">
                                    <Border Background="{TemplateBinding Background}"
                                            BorderBrush="{TemplateBinding BorderBrush}"
                                            BorderThickness="{TemplateBinding BorderThickness}"
                                            CornerRadius="2">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter Property="Background" Value="#FF4F4F53"/>
                                        </Trigger>
                                        <Trigger Property="IsPressed" Value="True">
                                            <Setter Property="Background" Value="#FF007ACC"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                    <Style x:Key="ActionButtonStyle" TargetType="Button">
                        <Setter Property="Background" Value="#FF007ACC"/>
                        <Setter Property="Foreground" Value="White"/>
                        <Setter Property="BorderThickness" Value="0"/>
                        <Setter Property="FontWeight" Value="Bold"/>
                        <Setter Property="Padding" Value="10,7"/>
                        <Setter Property="Margin" Value="5,15,5,5"/>
                        <Setter Property="Height" Value="60"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="Button">
                                    <Border Background="{TemplateBinding Background}" CornerRadius="2">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter Property="Background" Value="#FF005A9E"/>
                                        </Trigger>
                                        <Trigger Property="IsPressed" Value="True">
                                            <Setter Property="Background" Value="#FF004C87"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </Grid.Resources>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Label Content="Input Path:" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" Foreground="White" Margin="0,0,10,0"/>
                <TextBox Name="TextBoxInputPath" Grid.Row="0" Grid.Column="1" Margin="5" Padding="5" Background="#FF3E3E42" Foreground="Gray" BorderBrush="#FF007ACC" VerticalContentAlignment="Center" Text="Drag in file(s) or folder" AllowDrop="True" IsReadOnly="True"/>
                <Button Name="ButtonBrowseFolder" Grid.Row="0" Grid.Column="2" Content="Browse..." Style="{StaticResource BrowseButtonStyle}"/>
                <Grid Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="3" Margin="5,15,5,5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Button Name="ButtonGetFileNames" Grid.Column="0" Content="Get File Names" Style="{StaticResource ActionButtonStyle}" Margin="0,0,2.5,0"/>
                    <Button Name="ButtonRenameFiles" Grid.Column="1" Content="Rename Files" Style="{StaticResource ActionButtonStyle}" Margin="2.5,0,0,0"/>
                </Grid>
                <TextBox Name="TextBoxFileNames" Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="3" Margin="5" Padding="5" Background="#FF3E3E42" Foreground="White" BorderBrush="#FF007ACC" AcceptsReturn="True" AcceptsTab="False" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" AllowDrop="True"/>
            </Grid>
            <Grid x:Name="ResizeGripVisual" Grid.Row="1" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="18" Height="18" Margin="0,0,0,0" IsHitTestVisible="False">
                <Path Stroke="#888888" StrokeThickness="1.5">
                    <Path.Data>
                        <LineGeometry StartPoint="5,13" EndPoint="13,5"/>
                    </Path.Data>
                </Path>
                <Path Stroke="#888888" StrokeThickness="1.5">
                    <Path.Data>
                        <LineGeometry StartPoint="9,13" EndPoint="13,9"/>
                    </Path.Data>
                </Path>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

# Load XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $Global:Window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Error "Error loading XAML: $($_.Exception.Message)"
    Write-Host "--- XAML Content ---"; Write-Host $xaml; Write-Host "--- End XAML Content ---"
    exit 1
}

# Find controls
$CloseButton = $Global:Window.FindName("CloseButton")
$TextBoxInputPath = $Global:Window.FindName("TextBoxInputPath")
$ButtonBrowseFolder = $Global:Window.FindName("ButtonBrowseFolder")
$ButtonGetFileNames = $Global:Window.FindName("ButtonGetFileNames")
$ButtonRenameFiles = $Global:Window.FindName("ButtonRenameFiles")
$TextBoxFileNames = $Global:Window.FindName("TextBoxFileNames")

# Drag-and-drop handlers
$ScriptBlock_PreviewDragEnterOver = {
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $e.Effects = [System.Windows.DragDropEffects]::Copy
    } else {
        $e.Effects = [System.Windows.DragDropEffects]::None
    }
    $e.Handled = $true
}

$ScriptBlock_Drop_InputPath = {
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $Global:DroppedFiles = @()
        $TextBoxInputPath.Text = "Drag in file(s) or folder"
        $TextBoxInputPath.Foreground = "Gray"
        
        $files = @($e.Data.GetData([System.Windows.DataFormats]::FileDrop) | ForEach-Object { [System.IO.Path]::GetFullPath($_) })
        if ($files.Count -eq 1 -and [System.IO.Directory]::Exists($files[0])) {
            $TextBoxInputPath.Text = $files[0]
            $TextBoxInputPath.Foreground = "White"
            $Global:DroppedFiles = @($files[0])
        } else {
            $validFiles = @($files | Where-Object { $_ -and [System.IO.File]::Exists($_) })
            if ($validFiles.Count -eq 0) {
                Show-MessageBox -Message "Please drop valid files or a single folder." -Type "Error"
                return
            }
            $TextBoxInputPath.Text = $validFiles -join $Global:PathDelimiter
            $TextBoxInputPath.Foreground = "White"
            $Global:DroppedFiles = $validFiles
        }
    }
    $e.Handled = $true
}

$ScriptBlock_Drop_FileNames = {
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::Text)) {
        $text = $e.Data.GetData([System.Windows.DataFormats]::Text)
        $TextBoxFileNames.Text = $text
    } elseif ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $files = @($e.Data.GetData([System.Windows.DataFormats]::FileDrop) | ForEach-Object { [System.IO.Path]::GetFullPath($_) })
        $validFiles = @($files | Where-Object { $_ -and [System.IO.File]::Exists($_) })
        if ($validFiles.Count -eq 0) {
            Show-MessageBox -Message "Please drop valid files." -Type "Error"
            return
        }
        $fileNames = @($validFiles | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_) })
        $TextBoxFileNames.Text = $fileNames -join "`n"
    }
    $e.Handled = $true
}

# Attach drag-drop event handlers
$TextBoxInputPath.Add_PreviewDragEnter($ScriptBlock_PreviewDragEnterOver)
$TextBoxInputPath.Add_PreviewDragOver($ScriptBlock_PreviewDragEnterOver)
$TextBoxInputPath.Add_Drop($ScriptBlock_Drop_InputPath)
$TextBoxFileNames.Add_PreviewDragEnter($ScriptBlock_PreviewDragEnterOver)
$TextBoxFileNames.Add_PreviewDragOver($ScriptBlock_PreviewDragEnterOver)
$TextBoxFileNames.Add_Drop($ScriptBlock_Drop_FileNames)

# Button event handlers
$CloseButton.Add_Click({
    $Global:Window.Close()
})

$ButtonBrowseFolder.Add_Click({
    $currentPath = $TextBoxInputPath.Text.Trim()
    $pathHint = if ($currentPath -and (Test-Path $currentPath)) { $currentPath } else { $Global:DroppedFiles[0] }
    $selectedPath = Show-FolderBrowserDialog -InitialPathHint $pathHint
    if ($selectedPath) {
        $TextBoxInputPath.Text = $selectedPath
        $TextBoxInputPath.Foreground = "White"
        $Global:DroppedFiles = @($selectedPath)
    }
})

$ButtonGetFileNames.Add_Click({
    $inputPath = $TextBoxInputPath.Text.Trim()
    $TextBoxFileNames.Text = ""
    $files = @()
    
    # Check if inputPath is empty
    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        Show-MessageBox -Message "Input path is empty. Please provide a valid folder or file paths." -Type "Error"
        return
    }
    
    # Check if inputPath contains the delimiter (indicating a list of files)
    if ($inputPath -like "*$($Global:PathDelimiter)*") {
        # Split inputPath by delimiter, trim whitespace, and remove control characters
        $potentialFiles = $inputPath -split [regex]::Escape($Global:PathDelimiter) | ForEach-Object { $_.Trim() -replace '[\x00-\x1F]', '' }
        foreach ($file in $potentialFiles) {
            if ([string]::IsNullOrWhiteSpace($file)) {
                Show-MessageBox -Message "Empty file path detected in the input list." -Type "Error"
                return
            }
            try {
                $normalizedFile = [System.IO.Path]::GetFullPath($file)
                if (Test-Path -LiteralPath $normalizedFile -PathType Leaf -ErrorAction Stop) {
                    $files += $normalizedFile
                } else {
                    Show-MessageBox -Message "File does not exist or is not a valid file: $file" -Type "Error"
                    return
                }
            } catch {
                Show-MessageBox -Message "Error validating file path '$file': $_" -Type "Error"
                return
            }
        }
    } else {
        # Treat inputPath as a directory
        try {
            $normalizedPath = [System.IO.Path]::GetFullPath($inputPath)
            if (Test-Path -LiteralPath $normalizedPath -PathType Container -ErrorAction Stop) {
                $files = Get-ChildItem -LiteralPath $normalizedPath -File
            } else {
                Show-MessageBox -Message "Invalid directory path: $inputPath" -Type "Error"
                return
            }
        } catch {
            Show-MessageBox -Message "Error validating directory path '$inputPath': $_" -Type "Error"
            return
        }
    }
    
    if ($files.Count -eq 0) {
        Show-MessageBox -Message "No files found in the specified input." -Type "Error"
        return
    }
    
    $fileNames = @()
    foreach ($file in $files) {
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $fileNames += $fileName
    }
    $TextBoxFileNames.Text = $fileNames -join "`n"
})

$ButtonRenameFiles.Add_Click({
    $inputPath = $TextBoxInputPath.Text.Trim()
    $newFileNames = $TextBoxFileNames.Text -split "`n" | ForEach-Object { $_.Trim() }
    $files = @()
    
    # Check if inputPath is empty
    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        Show-MessageBox -Message "Input path is empty. Please provide a valid folder or file paths." -Type "Error"
        return
    }
    
    # Check if inputPath contains the delimiter (indicating a list of files)
    if ($inputPath -like "*$($Global:PathDelimiter)*") {
        # Split inputPath by delimiter, trim whitespace, and remove control characters
        $potentialFiles = $inputPath -split [regex]::Escape($Global:PathDelimiter) | ForEach-Object { $_.Trim() -replace '[\x00-\x1F]', '' }
        foreach ($file in $potentialFiles) {
            if ([string]::IsNullOrWhiteSpace($file)) {
                Show-MessageBox -Message "Empty file path detected in the input list." -Type "Error"
                return
            }
            try {
                $normalizedFile = [System.IO.Path]::GetFullPath($file)
                if (Test-Path -LiteralPath $normalizedFile -PathType Leaf -ErrorAction Stop) {
                    $files += $normalizedFile
                } else {
                    Show-MessageBox -Message "File does not exist or is not a valid file: $file" -Type "Error"
                    return
                }
            } catch {
                Show-MessageBox -Message "Error validating file path '$file': $_" -Type "Error"
                return
            }
        }
    } else {
        # Treat inputPath as a directory
        try {
            $normalizedPath = [System.IO.Path]::GetFullPath($inputPath)
            if (Test-Path -LiteralPath $normalizedPath -PathType Container -ErrorAction Stop) {
                $files = Get-ChildItem -LiteralPath $normalizedPath -File
            } else {
                Show-MessageBox -Message "Invalid directory path: $inputPath" -Type "Error"
                return
            }
        } catch {
            Show-MessageBox -Message "Error validating directory path '$inputPath': $_" -Type "Error"
            return
        }
    }
    
    if ($files.Count -ne $newFileNames.Count) {
        Show-MessageBox -Message "The number of new file names ($($newFileNames.Count)) does not match the number of files ($($files.Count))." -Type "Error"
        return
    }
    
    for ($i = 0; $i -lt $files.Count; $i++) {
        $file = $files[$i]
        $newFileName = $newFileNames[$i]
        if ([string]::IsNullOrWhiteSpace($newFileName)) {
            Show-MessageBox -Message "Empty file name provided for file: $file" -Type "Error"
            return
        }
        $newFilePath = Join-Path (Split-Path $file -Parent) "$newFileName$([System.IO.Path]::GetExtension($file))"
        try {
            Rename-Item -LiteralPath $file -NewName $newFilePath -ErrorAction Stop
        } catch {
            Show-MessageBox -Message "Error renaming file '$file' to '$newFilePath': $_" -Type "Error"
            return
        }
    }
    Show-MessageBox -Message "Files renamed successfully." -Type "Info"
})

# Show GUI
$Global:Window.ShowDialog() | Out-Null