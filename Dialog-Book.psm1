###############################################################################
# Dialog Display Script Book
# GUI dialogs framework for common shell libraries
###############################################################################

Add-Type -AssemblyName 'PresentationFramework', 'System.Windows.Forms', 'System.Drawing', 'Microsoft.VisualBasic'

# Definition of color palette for host messages
# (c) 30.10.2020 khristal
$PalInfo    = @{ForegroundColor = [System.ConsoleColor]::White}
$PalDebug   = @{ForegroundColor = [System.ConsoleColor]::Gray}
$PalSuccess = @{ForegroundColor = [System.ConsoleColor]::Green}
$PalWarning = @{ForegroundColor = [System.ConsoleColor]::Yellow}
$PalError   = @{ForegroundColor = [System.ConsoleColor]::Red}
$PalFault   = @{ForegroundColor = [System.ConsoleColor]::Red; BackgroundColor = [System.ConsoleColor]::White}

# Show InputBox <VB> and return entered [string] text or [string[]] list
# (c) aug 2018 khristal
# (c) 30.10.2020 khristal - added empty list item filter
function Input-String
{
    param (
        [string]$Text   = "",                             # initial text
        [string]$Prompt = "Enter a list comma delimited", # dialog prompt
        [string]$Title  = "Input",                        # dialog title
        [string]$Delim  = "`t;,"                          # list delimiter(s) (set empty string for disable)
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic'))
    {
#        [Microsoft.VisualBasic.Interaction]::InputBox($prompt, $title, $text).Split($delim).Trim()
        [Microsoft.VisualBasic.Interaction]::InputBox($prompt, $title, $text).Split($delim) | % {$_.Trim() | ? {$_}}
    }
}

# Show MsgBox <VB> and return pressed [Microsoft.VisualBasic.MsgBoxResult] button
# (c) jun 2020 khristal
function Display-Message
{
    param (
        [string]$Msg                               = "Hello!",                                         # message text
        [string]$Title                             = "Info",                                           # dialog title
        [Microsoft.VisualBasic.MsgBoxStyle]$Button = [Microsoft.VisualBasic.MsgBoxStyle]::OkOnly,      # dialog button
        [Microsoft.VisualBasic.MsgBoxStyle]$Icon   = [Microsoft.VisualBasic.MsgBoxStyle]::Information, # dialog icon
        [switch]$Top                               = $true                                             # display topmost
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic'))
    {
        [Microsoft.VisualBasic.Interaction]::MsgBox($msg, $button + $icon +
         $(if ($top) {[Microsoft.VisualBasic.MsgBoxStyle]::SystemModal} else {[Microsoft.VisualBasic.MsgBoxStyle]::ApplicationModal}), $title)
    }
}

# Show Popup message <WS> and return pressed [Microsoft.VisualBasic.MsgBoxResult] button or -1 = timeout
# (c) jun 2020 khristal
function Popup-Message
{
    param (
        [string]$Msg                               = "Hello!",                                         # message text
        [string]$Title                             = "Info",                                           # dialog title
        [Microsoft.VisualBasic.MsgBoxStyle]$Button = [Microsoft.VisualBasic.MsgBoxStyle]::OkOnly,      # dialog button
        [Microsoft.VisualBasic.MsgBoxStyle]$Icon   = [Microsoft.VisualBasic.MsgBoxStyle]::Information, # dialog icon
        [int32]$Timeout                            = 0                                                 # dialog timeout in seconds (set 0 for disable)
    )

    if ($shell = $(New-Object -ComObject 'WScript.Shell'))
    {
        $shell.Popup($msg, $timeout, $title, $button + $icon)
    }
}

# Show MessageBox <WPF> and return pressed [System.Windows.MessageBoxResult] button
# (c) jun 2020 khristal
function View-Message
{
    param (
        [string]$Msg                             = "Hello!",                                      # message text
        [string]$Title                           = "Info",                                        # dialog title
        [System.Windows.MessageBoxButton]$Button = [System.Windows.MessageBoxButton]::OK,         # dialog button
        [System.Windows.MessageBoxImage]$Icon    = [System.Windows.MessageBoxImage]::Information, # dialog icon
        [switch]$Top                             = $true                                          # display topmost
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework'))
    {
        if ($win = New-Object -TypeName 'System.Windows.Window')
        {
            $win.TopMost = $top
            $null = $win.Activate()
            [System.Windows.MessageBox]::Show($win, $msg, $title, $button, $icon)
        }
    }
}

# Show MessageBox <WFS> and return pressed [System.Windows.Forms.DialogResult] button
# (c) jun 2020 khristal
function Show-Message
{
    param (
        [string]$Msg                                    = "Hello!",                                           # message text
        [string]$Title                                  = "Info",                                             # dialog title
        [System.Windows.Forms.MessageBoxButtons]$Button = [System.Windows.Forms.MessageBoxButtons]::OK,       # dialog button
        [System.Windows.Forms.MessageBoxIcon]$Icon      = [System.Windows.Forms.MessageBoxIcon]::Information, # dialog icon
        [switch]$Top                                    = $true                                               # display topmost
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
    {
        if ($form = New-Object -TypeName 'System.Windows.Forms.Form')
        {
            $form.TopMost = $top
            $form.Activate()
            [System.Windows.Forms.MessageBox]::Show($form, $msg, $title, $button, $icon)
        }
    }
}

# Show Notification message <WFS> in taskbar and return [void]
# (c) dec 2020 khristal
function Notify-Message
{
    param (
        [string]$Msg                            = "Hello!",                                  # message text
        [string]$Title                          = "Info",                                    # balloon title
        [System.Windows.Forms.ToolTipIcon]$Icon = [System.Windows.Forms.ToolTipIcon]::Info , # balloon icon
        [int32]$Timeout                         = 60000,                                     # balloon timeout in milliseconds
        [switch]$Keep                           = $false                                     # keep tooltip after timeout
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
    {
        if ($tip = New-Object -TypeName 'System.Windows.Forms.NotifyIcon')
        {
            $tip.Text    = $title
            $tip.Icon    = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $pid).Path)
            $tip.Visible = $true
            $tip.ShowBalloonTip($timeout, $title, $msg, $icon)
            if (!$keep)
            {
                Start-Sleep -Milliseconds $timeout
                $tip.Dispose()
            }
        }
    }
}

# Show Open File dialog <WFS> and return selected file [string] or [string[]] path(s)
# (c) aug 2018 khristal
function Open-FileName
{
    param (
        [string]$Folder = [environment]::CurrentDirectory, # initial folder
        [string]$Filter = "All Files (*.*)|*.*",           # file filter
        [string]$Title  = "",                              # dialog title
        [switch]$Multi  = $false                           # allow multiselect
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
    {
        if ($dialog = New-Object -TypeName 'System.Windows.Forms.OpenFileDialog')
        {
            $dialog.InitialDirectory = $folder
            $dialog.Filter           = $filter
            $dialog.Title            = $title
            $dialog.Multiselect      = $multi
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                if ($multi) {$dialog.Filenames} else {$dialog.Filename}
            }
        }
    }
}

# Show Save File dialog <WFS> and return selected file [string] path
# (c) aug 2018 khristal
function Save-FileName
{
    param (
        [string]$Folder = [environment]::CurrentDirectory, # initial folder
        [string]$Filter = "All Files (*.*)|*.*",           # file filter
        [string]$Title  = "",                              # dialog title
        [switch]$Prompt = $true                            # overwrite prompt
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
    {
        if ($dialog = New-Object -TypeName 'System.Windows.Forms.SaveFileDialog')
        {
            $dialog.InitialDirectory = $folder
            $dialog.Filter           = $filter
            $dialog.Title            = $title
            $dialog.OverwritePrompt  = $prompt
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $dialog.Filename
            }
        }
    }
}

# Show Folder Browser dialog <WFS> and return selected folder [string] path
# (c) aug 2018 khristal
function Browse-Folder
{
    param (
        [string]$Folder = "MyComputer",    # initial special folder
        [string]$Desc   = "Select folder", # dialog description
        [switch]$New    = $false           # allow create button
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
    {
        if ($dialog = New-Object -TypeName 'System.Windows.Forms.FolderBrowserDialog')
        {
            $dialog.RootFolder          = $folder
            $dialog.Description         = $desc
            $dialog.ShowNewFolderButton = $new
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $dialog.SelectedPath
            }
        }
    }
}

# Show Folder Browser dialog <WS> and return selected folder [string] path
# (c) jun 2020 khristal
function BrowseFor-Folder
{
    param (
        [string]$Folder = "",              # initial root folder (also limits selection)
        [string]$Desc   = "Select folder", # dialog description
        [int32]$Type    = 1 + 16 + 512     # dialog type: 1 - return fso only, 16 - enable edit box, 512 - disable new button
    )

    if ($shell = $(New-Object -ComObject 'Shell.Application'))
    {
        if ($result = $shell.BrowseForFolder(0, $desc, $type, $folder))
        {
            $result.Self.Path
        }
    }
}

# Show Pick File/Folder dialog <MSO> and return selected file/folder [string[]] path(s)
# requires Microsoft Office installed
# (c) oct 2020 khristal
function Pick-Folder
{
    param (
        [string]$Folder    = [environment]::CurrentDirectory + "\", # initial folder
        [string[]]$Filters = @("All Files (*.*)|*.*" -split '\|'),  # file filter(s)
        [string]$Title     = "",                                    # dialog title
        [int32]$Type       = 4,                                     # dialog type: 1 - msoFileDialogOpen, 2 - msoFileDialogSaveAs, 3 - msoFileDialogFilePicker, 4 - msoFileDialogFolderPicker
        [switch]$Multi     = $false                                 # allow multiselect
    )

    if ($excel = $(New-Object -ComObject 'Excel.Application'))
    {
        if ($dialog = $excel.FileDialog($type))
        {
            $dialog.InitialFileName  = $folder
            $dialog.Title            = $title
            $dialog.AllowMultiselect = $multi
            if ($type -lt 4) {$dialog.Filters.Clear(); for ($i = 0; $i -lt $filters.Count; $i += 2) {
                $dialog.Filters.Add($filters[$i], $filters[$i+1]) | Out-Null}}
            if ($dialog.Show() -eq -1) # -1 = action, 0 = cancel
            {
                $dialog.SelectedItems
            }
        }
    }
    else
    {
        Write-Warning -Message "Failed to initialize MS Office application!"
    }
}

# Show Choose Color dialog <WFS> and return selected [System.Drawing.Color] color
# (c) oct 2020 khristal
function Pick-Color
{
    param (
        [psobject]$Color = 'Black', # initial color name or rgb value
        [int32[]]$Custom = @(),     # custom rgb palette
        [switch]$Full    = $false   # full open
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
    {
        if ($dialog = New-Object -TypeName 'System.Windows.Forms.ColorDialog')
        {
            $dialog.Color        = $color
            $dialog.CustomColors = $custom
            $dialog.FullOpen     = $full
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $dialog.Color
            }
        }
    }
}

# Show Choose Font dialog <WFS> and return selected [System.Drawing.Font] font + optional [System.Drawing.Color] color
# (c) oct 2020 khristal
function Pick-Font
{
    param (
        [string]$Font  = 'Tahoma', # initial font name
        [switch]$Color = $false,   # also selects font color
        [switch]$All   = $true     # show all fonts
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
    {
        if ($dialog = New-Object -TypeName 'System.Windows.Forms.FontDialog')
        {
            $dialog.Font        = $font
            $dialog.ShowColor   = $color
            $dialog.ScriptsOnly = !$all
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                if ($color) {Add-Member -InputObject $dialog.Font -MemberType 'NoteProperty' -Name 'Color' -Value $dialog.Color -Force -PassThru} else {$dialog.Font}
            }
        }
    }
}

# Show Month Calendar <WFS> and return selected [datetime] date, or [System.Windows.Forms.SelectionRange] dates
# (c) dec 2020 khristal
function Pick-Date
{
    param (
        [datetime[]]$Date = @(Get-Date), # initial date(s)
        [string]$Title    = "Calendar",  # dialog title
        [int32]$Multi     = 1,           # max selection count in days
        [switch]$Weeks    = $false       # show week numbers
    )

    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
    {
        if ($form = New-Object -TypeName 'System.Windows.Forms.Form')
        {
            $form.Text            = $title
            $form.Icon            = [System.Drawing.SystemIcons]::Application
            $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
            $form.ClientSize      = New-Object -TypeName 'System.Drawing.Size' -ArgumentList $(if ($weeks) {186} else {164}), 180
            $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
            $form.MinimizeBox     = $false
            $form.MaximizeBox     = $false
            $form.Add_Shown({$form.TopMost = $true; $form.Activate()})

            if ($cal = New-Object -TypeName 'System.Windows.Forms.MonthCalendar')
            {
                $cal.MaxSelectionCount = $multi
                $cal.SelectionStart    = if ($date[0] -le $date[-1]) {$date[0]} else {$date[-1]}
                $cal.SelectionEnd      = if ($date[0] -ge $date[-1]) {$date[0]} else {$date[-1]}
                $cal.ShowWeekNumbers   = $weeks
                $cal.ShowToday         = $true
                $cal.ShowTodayCircle   = $true
                $form.Controls.Add($cal)

                $buttons = @('OK', 'Cancel'); $btn = @()
                for ($i = 0; $i -lt $buttons.Count; $i++)
                {
                    if ($btn += New-Object -TypeName 'System.Windows.Forms.Button')
                    {
                        $btn[$i].Text         = $buttons[$i]
                        $btn[$i].DialogResult = [System.Windows.Forms.DialogResult]::$($buttons[$i])
                        $btn[$i].Location     = New-Object -TypeName 'System.Drawing.Point' -ArgumentList ($i * $form.ClientSize.Width / $buttons.Count), (9 * $form.ClientSize.Height / 10)
                        $btn[$i].Size         = New-Object -TypeName 'System.Drawing.Size' -ArgumentList ($form.ClientSize.Width / $buttons.Count), ($form.ClientSize.Height / 10)
                        $form.Controls.Add($btn[$i])
                    }
                }

                if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
                {
                    if ($multi -gt 1) {$cal.SelectionRange} else {$cal.SelectionStart}
                }
            }
        }
    }
}

# Show custom form dialog <WFS> and return [string] pressed button caption, confirmed textbox content, or 'Timeout' if reached
# (c) dec 2020 khristal
function Show-Dialog
{
    param (
        [string]$Msg       = "Hello!",      # message header
        [string]$Text      = "",            # message text
        [string]$Title     = "Info",        # dialog title
        [string[]]$Buttons = @('OK'),       # dialog button(s)
        [string]$Image     = "$env:WinDir\System32\SecurityAndMaintenance.png", # dialog picture
        [string]$Icon      = 'Information', # dialog icon
        [string]$Font      = 'Tahoma',      # dialog font name
        [single]$Scale     = 1,             # dialog scale factor
        [double]$Opacity   = 1,             # dialog opacity
        [int32]$Timeout    = 0,             # dialog timeout in milliseconds (set 0 for disable)
        [switch]$Prompt    = $false,        # return textbox content if confirmed by first button
        [switch]$Top       = $true          # display topmost
    )

    New-Variable -Name 'Result' -Value '' -Option 'AllScope' -Force # return value
    if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
    {
        if ($form = New-Object -TypeName 'System.Windows.Forms.Form')
        {
            $form.Text            = $title
            $form.Icon            = [System.Drawing.SystemIcons]::$icon
            $form.ClientSize      = New-Object -TypeName 'System.Drawing.Size' -ArgumentList (800 * $scale), (600 * $scale)
            $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
            $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $form.MinimizeBox     = $false
            $form.MaximizeBox     = $false
            $form.Opacity         = $opacity
            $form.Add_Shown({$form.TopMost = $top; $form.Activate()})
            $form.Add_FormClosing({if ($timer) {$timer.Stop(); $timer.Dispose()}})

            if ($image)
            {
                if ($pic = New-Object -TypeName 'System.Windows.Forms.PictureBox')
                {
                    if ($pic.Image = $([System.Drawing.Image]::FromFile($image)))
                    {
                        $pic.Location = New-Object -TypeName 'System.Drawing.Point' -ArgumentList (10 * $scale), (10 * $scale)
                        $pic.Size     = New-Object -TypeName 'System.Drawing.Size' -ArgumentList (1/3 * $form.ClientSize.Width - 20 * $scale), (8/9 * $form.ClientSize.Height - 20 * $scale)
                        $pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
                        $form.Controls.Add($pic)
                    }
                }
            }
            if ($msg)
            {
                if ($lbl = New-Object -TypeName 'System.Windows.Forms.Label')
                {
                    $lbl.Text      = $msg
                    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $lbl.Font      = New-Object -TypeName 'System.Drawing.Font' -ArgumentList $font, (18 + 10 * [math]::Log10($scale)), ([System.Drawing.FontStyle]::Bold)
                    $lbl.Location  = New-Object -TypeName 'System.Drawing.Point' -ArgumentList (1/3 * $form.ClientSize.Width), (10 * $scale)
                    $lbl.Size      = New-Object -TypeName 'System.Drawing.Size' -ArgumentList (2/3 * $form.ClientSize.Width - 20 * $scale), (4/9 * $form.ClientSize.Height - 20 * $scale)
                    $form.Controls.Add($lbl)
                }
            }
            if ($text -or $prompt)
            {
                if ($txt = New-Object -TypeName 'System.Windows.Forms.TextBox')
                {
                    $txt.Text        = $text
                    $txt.Multiline   = $true
                    $txt.ReadOnly    = !$prompt
                    $txt.BorderStyle = [System.Windows.Forms.BorderStyle]::None
                    $txt.Font        = New-Object -TypeName 'System.Drawing.Font' -ArgumentList $font, (14 + 10 * [math]::Log10($scale))
                    $txt.Location    = New-Object -TypeName 'System.Drawing.Point' -ArgumentList (1/3 * $form.ClientSize.Width), (4/9 * $form.ClientSize.Height)
                    $txt.Size        = New-Object -TypeName 'System.Drawing.Size' -ArgumentList (2/3 * $form.ClientSize.Width - 20 * $scale), (4/9 * $form.ClientSize.Height - 20 * $scale)
                    $form.Controls.Add($txt)
                }
            }
            if ($buttons)
            {
                $btn = @()
                for ($i = 0; $i -lt $buttons.Count; $i++)
                {
                    if ($btn += New-Object -TypeName 'System.Windows.Forms.Button')
                    {
                        $btn[$i].Text     = $buttons[$i]
                        $btn[$i].Tag      = $i
                        $btn[$i].Location = New-Object -TypeName 'System.Drawing.Point' -ArgumentList (($i * $form.ClientSize.Width + 10 * $scale) / $buttons.Count), (8/9 * $form.ClientSize.Height)
                        $btn[$i].Size     = New-Object -TypeName 'System.Drawing.Size' -ArgumentList (($form.ClientSize.Width - 20 * $scale) / $buttons.Count), (1/9 * $form.ClientSize.Height - 20 * $scale)
                        $btn[$i].Add_Click({$result = if ($prompt -and !$this.Tag) {$txt.Text} else {$this.Text}; $form.Dispose()})
                        $form.Controls.Add($btn[$i])
                    }
                }
                $btn[0].Select()
            }
            if ($timeout)
            {
                if ($timer = New-Object -TypeName 'System.Windows.Forms.Timer')
                {
                    $timer.Interval = $timeout
                    $timer.Add_Tick({$result = 'Timeout'; $form.Dispose()})
                    $timer.Start()
                }
            }

            $null = $form.ShowDialog()
            $result
        }
    }
}

# Show flexible PictureBox <WFS> and return [void]
# (c) dec 2020 khristal
function Show-Picture
{
    param (
        [psobject]$Image                               = (Open-FileName),                                 # picture path or stream
        [System.Windows.Forms.PictureBoxSizeMode]$Mode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom, # picture scale mode
        [switch]$Resize                                = $false                                           # allow resize
    )

    if ($image -and ($img = $(if ($image -is [string]) {[System.Drawing.Image]::FromFile($image)} else {[System.Drawing.Image]::FromStream($image)})))
    {
        if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
        {
            if ($form = New-Object -TypeName 'System.Windows.Forms.Form')
            {
                $form.Text            = "$($img.Width) x $($img.Height) @ $($img.VerticalResolution)DPI"
                $form.ClientSize      = New-Object -TypeName 'System.Drawing.Size' -ArgumentList $img.Width, $img.Height
                $form.WindowState     = [System.Windows.Forms.FormWindowState]::$(
                    if ($img.Width  -gt [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width -or
                        $img.Height -gt [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height) {'Maximized'} else {'Normal'})
                $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::$(if ($resize) {'Sizable'} else {'FixedSingle'})
                $form.MinimizeBox     = $true
                $form.MaximizeBox     = $resize
                $form.Add_Shown({$form.TopMost = $true; $form.Activate()})
                $form.Add_SizeChanged({$pic.Size = New-Object -TypeName 'System.Drawing.Size' -ArgumentList $form.ClientSize.Width, $form.ClientSize.Height})

                if ($pic = New-Object -TypeName 'System.Windows.Forms.PictureBox')
                {
                    $pic.Image    = $img
                    $pic.SizeMode = $mode
                    $pic.Location = New-Object -TypeName 'System.Drawing.Point' -ArgumentList 0, 0
                    $pic.Size     = New-Object -TypeName 'System.Drawing.Size' -ArgumentList $form.ClientSize.Width, $form.ClientSize.Height
                    $pic.Add_Click({$form.Dispose()})
                    $form.Controls.Add($pic)

                    $null = $form.ShowDialog()
                }
            }
        }
    }
}

# Play media file <WPF> and return [void]
# (c) dec 2020 khristal
function Play-Media
{
    param (
        [string[]]$List = @(Open-FileName -Multi), # playlist of media file(s) (override empty for stop)
        [switch]$Random = $false,                  # shuffle play
        [switch]$Loop   = $false,                  # loop play
        [switch]$Back   = $true                    # background playback
    )

    $job = 'Player' # background job name
    Get-Job | Where-Object -FilterScript {$_.Name -eq $job} | Remove-Job -Force
    if ($list)
    {
        if ($back)
        {
            Start-Job -Name $job -ScriptBlock ([scriptblock]::Create((${function:Play-Media} -replace "\[switch\]\$", "[bool]$"))) -ArgumentList $list, $random, $loop, $false | Out-Null
        }
        else
        {
            if ([System.Reflection.Assembly]::LoadWithPartialName('PresentationCore'))
            {
                if ($player = New-Object -TypeName 'System.Windows.Media.MediaPlayer')
                {
                    do
                    {
                        if ($random) {$list = $list | Sort-Object -Property {Get-Random}}
                        foreach ($file in $list)
                        {
                            $player.Open([uri]$file)
                            Start-Sleep -Milliseconds 600
                            $player.Play()
                            Start-Sleep -Seconds $player.NaturalDuration.TimeSpan.TotalSeconds
                            $player.Stop()
                            $player.Close()
                        }
                    }
                    while ($loop)
                }
            }
        }
    }
}

# Show Print Dialog <WFS> and print input content as plain text
# (c) oct 2020 khristal
function Print-Content
{
    param (
        [Parameter(ValueFromPipeline = $true)]
        [array]$Text     = @(),   # input content [may be piped]
        [string]$Printer = "",    # printer name; if not specified, system defaut printer will be used
        [int16]$Copies   = 1,     # number of copies
        [switch]$Quiet   = $false # silent mode
    )

    begin
    {
        $list = @()
    }
    process
    {
        $list += $text | % {$_ | Out-String}
    }
    end
    {
        if ([System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'))
        {
            if ($dialog = New-Object -TypeName 'System.Windows.Forms.PrintDialog')
            {
                $dialog.PrinterSettings.PrinterName = $printer
                $dialog.PrinterSettings.Copies      = $copies
                if ($quiet -or $dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
                {
                    foreach ($item in $list)
                    {
                        $copies = $dialog.PrinterSettings.Copies
                        while ($copies--)
                        {
                            $item | Out-Printer -Name $dialog.PrinterSettings.PrinterName
                        }
                    }
                }
            }
        }
    }
}

Export-ModuleMember -Function "*-*" -Variable "Pal*"
