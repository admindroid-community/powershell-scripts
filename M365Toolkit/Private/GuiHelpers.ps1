<#
    GUI helper functions.

    These live at module scope rather than nested inside Show-M365Console on purpose.
    WPF event handlers and DispatcherTimer ticks fire through scriptblocks created with
    GetNewClosure(), which captures *variables* but not function definitions. A helper
    defined inside Show-M365Console may therefore fail to resolve when a handler runs
    later. Module-scope functions always resolve.

    Shared UI state is held in $script:Gui* variables, set by Show-M365Console.
#>

$script:GuiUi       = $null
$script:GuiState    = $null
$script:GuiAllTasks = @()
$script:GuiRoot     = $null
$script:GuiWindow   = $null

function Invoke-InWorker {
    <#
        .SYNOPSIS
        Runs a scriptblock synchronously in the shared worker runspace.

        .DESCRIPTION
        All Microsoft 365 session state (Graph, Exchange) is per-runspace, so every
        toolkit call has to go through the same worker.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Script,
        [object[]]$Arguments = @()
    )

    $shell = [powershell]::Create()
    $shell.Runspace = $script:GuiState.Worker
    $null = $shell.AddScript($Script.ToString())
    foreach ($a in $Arguments) { $null = $shell.AddArgument($a) }

    try {
        return @($shell.Invoke())
    } catch {
        Write-M365Log -Level Error -Source 'GUI' -Message $_.Exception.Message
        return @()
    } finally {
        $shell.Dispose()
    }
}

function Set-GuiStatus {
    [CmdletBinding()]
    param([string]$Text, [switch]$Busy)

    $script:GuiUi.LblStatus.Text = $Text
    $script:GuiUi.Progress.Visibility = if ($Busy) { 'Visible' } else { 'Hidden' }
    $script:GuiUi.Progress.IsIndeterminate = [bool]$Busy
}

function Update-GuiLog {
    [CmdletBinding()]
    param()
    $entries = Invoke-InWorker -Script { Get-M365Log -Last 300 }
    $script:GuiUi.GridLog.ItemsSource = @($entries)
}

function Update-GuiConnection {
    [CmdletBinding()]
    param()

    $conns = Invoke-InWorker -Script { Get-M365Connection }
    $live  = @($conns | Where-Object { $_.Connected })
    $script:GuiState.Connected = ($live.Count -gt 0)

    if ($script:GuiState.Connected) {
        $script:GuiUi.ConnDot.Fill  = [Windows.Media.Brushes]::SeaGreen
        $script:GuiUi.ConnText.Text = 'Connected: ' + (($live | ForEach-Object { $_.Service }) -join ', ')
        $ident = ($live | Where-Object { $_.Identity } | Select-Object -First 1)
        $script:GuiUi.ConnDetail.Text = if ($ident) { "$($ident.Identity)  ($($ident.AuthType))" } else { '' }
    } else {
        $script:GuiUi.ConnDot.Fill  = [Windows.Media.Brushes]::Firebrick
        $script:GuiUi.ConnText.Text = 'Not connected'
        $script:GuiUi.ConnDetail.Text = ''
    }

    $script:GuiUi.BtnRun.IsEnabled = ($script:GuiState.Connected -and $null -ne $script:GuiState.SelectedTask)
}

function Build-GuiTree {
    [CmdletBinding()]
    param()

    $search   = $script:GuiUi.TxtSearch.Text
    $category = $script:GuiUi.CmbCategory.SelectedItem

    $filtered = $script:GuiAllTasks
    if ($category -and $category -ne 'All categories') {
        $filtered = $filtered | Where-Object { $_.Category -eq $category }
    }
    if (-not [string]::IsNullOrWhiteSpace($search)) {
        $filtered = $filtered | Where-Object {
            $_.Name -like "*$search*" -or $_.Synopsis -like "*$search*" -or
            $_.Id   -like "*$search*" -or ($_.Replaces -and $_.Replaces -like "*$search*")
        }
    }

    $script:GuiUi.TreeTasks.Items.Clear()

    foreach ($grp in ($filtered | Group-Object Category | Sort-Object Name)) {
        $node = New-Object Windows.Controls.TreeViewItem
        $node.Header     = "$($grp.Name)  ($($grp.Count))"
        $node.FontWeight = 'SemiBold'
        $node.IsExpanded = $true

        foreach ($t in ($grp.Group | Sort-Object Name)) {
            $child = New-Object Windows.Controls.TreeViewItem
            $child.Header  = $t.Name
            $child.Tag     = $t
            $child.ToolTip = $t.Synopsis
            if     ($t.Risk -eq 'Destructive') { $child.Foreground = [Windows.Media.Brushes]::Firebrick }
            elseif ($t.Risk -eq 'Write')       { $child.Foreground = [Windows.Media.Brushes]::DarkOrange }
            $null = $node.Items.Add($child)
        }
        $null = $script:GuiUi.TreeTasks.Items.Add($node)
    }

    Set-GuiStatus "$(@($filtered).Count) task(s) listed"
}

function Build-GuiParamForm {
    [CmdletBinding()]
    param($Task)

    $script:GuiUi.PanelParams.Children.Clear()
    $script:GuiState.ParamControls = @{}

    $params = @($Task.Parameters)
    if ($params.Count -eq 0) {
        $tb = New-Object Windows.Controls.TextBlock
        $tb.Text       = 'This task takes no parameters.'
        $tb.Foreground = [Windows.Media.Brushes]::Gray
        $tb.FontSize   = 12
        $null = $script:GuiUi.PanelParams.Children.Add($tb)
        return
    }

    foreach ($p in $params) {
        $ptype    = if ($p.ContainsKey('Type')) { $p.Type } else { 'String' }
        $required = [bool]($p.ContainsKey('Required') -and $p.Required)
        $default  = if ($p.ContainsKey('Default')) { $p.Default } else { $null }

        $row = New-Object Windows.Controls.StackPanel
        $row.Margin = '0,0,0,9'

        $label = New-Object Windows.Controls.TextBlock
        $label.Text       = if ($required) { "$($p.Name) *" } else { $p.Name }
        $label.FontSize   = 12
        $label.FontWeight = 'SemiBold'
        $label.Margin     = '0,0,0,2'
        $null = $row.Children.Add($label)

        if ($p.ContainsKey('Help') -and $p.Help) {
            $help = New-Object Windows.Controls.TextBlock
            $help.Text         = $p.Help
            $help.FontSize     = 11
            $help.Foreground   = [Windows.Media.Brushes]::Gray
            $help.TextWrapping = 'Wrap'
            $help.Margin       = '0,0,0,3'
            $null = $row.Children.Add($help)
        }

        $control = switch ($ptype) {
            'Bool' {
                $c = New-Object Windows.Controls.CheckBox
                $c.IsChecked = [bool]$default
                $c.Content   = 'Enabled'
                $c.FontSize  = 12
                $c
            }
            'Choice' {
                $c = New-Object Windows.Controls.ComboBox
                $c.ItemsSource = @($p.Options)
                if ($null -ne $default)            { $c.SelectedItem = $default }
                elseif (@($p.Options).Count -gt 0) { $c.SelectedIndex = 0 }
                $c.Padding = '5'
                $c
            }
            'MultiChoice' {
                $c = New-Object Windows.Controls.ListBox
                $c.ItemsSource   = @($p.Options)
                $c.SelectionMode = 'Multiple'
                $c.MaxHeight     = 84
                foreach ($d in @($default)) { if ($null -ne $d) { $null = $c.SelectedItems.Add($d) } }
                $c
            }
            'DateTime' {
                $c = New-Object Windows.Controls.DatePicker
                if ($default) { $c.SelectedDate = [datetime]$default }
                $c
            }
            default {
                $c = New-Object Windows.Controls.TextBox
                $c.Text    = if ($null -ne $default) { [string]$default } else { '' }
                $c.Padding = '5'
                $c
            }
        }

        $null = $row.Children.Add($control)
        $null = $script:GuiUi.PanelParams.Children.Add($row)

        $script:GuiState.ParamControls[$p.Name] = [pscustomobject]@{ Control = $control; Type = $ptype }
    }
}

function Get-GuiParamValue {
    [CmdletBinding()]
    param()

    $values = @{}
    foreach ($name in $script:GuiState.ParamControls.Keys) {
        $entry = $script:GuiState.ParamControls[$name]
        $c = $entry.Control

        switch ($entry.Type) {
            'Bool'        { $values[$name] = [bool]$c.IsChecked }
            'Choice'      { $values[$name] = $c.SelectedItem }
            'MultiChoice' { $values[$name] = @($c.SelectedItems) }
            'DateTime'    { if ($c.SelectedDate) { $values[$name] = $c.SelectedDate } }
            default       { if (-not [string]::IsNullOrWhiteSpace($c.Text)) { $values[$name] = $c.Text } }
        }
    }
    return $values
}

function Complete-GuiTaskRun {
    <#
        .SYNOPSIS
        Called from the DispatcherTimer tick once the background task finishes.
    #>
    [CmdletBinding()]
    param()

    $st = $script:GuiState
    $ui = $script:GuiUi

    $envelope = $null
    $failure  = $null
    try {
        $envelope = @($st.Shell.EndInvoke($st.Handle)) | Select-Object -First 1
    } catch {
        $failure = $_.Exception.Message
    }

    $errText = ($st.Shell.Streams.Error | ForEach-Object { $_.ToString() }) -join '; '
    try { $st.Shell.Dispose() } catch { }
    $st.Shell  = $null
    $st.Handle = $null

    $ui.BtnCancel.IsEnabled = $false
    $ui.BtnRun.IsEnabled    = $true

    if ($failure) {
        Set-GuiStatus "Failed: $failure"
        [void][Windows.MessageBox]::Show($failure, 'Task failed',
            [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Error)
    } elseif ($envelope -and -not $envelope.Succeeded) {
        Set-GuiStatus "Failed: $($envelope.Error)"
        [void][Windows.MessageBox]::Show($envelope.Error, 'Task failed',
            [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Error)
    } elseif ($envelope) {
        $st.Results = @($envelope.Results)
        $ui.GridResults.ItemsSource = $st.Results
        $ui.LblRowCount.Text = "$($envelope.RowCount) row(s) in $([math]::Round($envelope.Duration.TotalSeconds,1))s"

        $has = $st.Results.Count -gt 0
        $ui.BtnExportCsv.IsEnabled  = $has
        $ui.BtnExportHtml.IsEnabled = $has
        $ui.BtnCopy.IsEnabled       = $has

        Set-GuiStatus "Completed: $($envelope.RowCount) row(s)"
        $ui.TabsOutput.SelectedIndex = 0

        if ($envelope.MissingScopes -and @($envelope.MissingScopes).Count -gt 0) {
            [void][Windows.MessageBox]::Show(
                "The task completed, but the token appears to be missing these scopes:`n`n$(@($envelope.MissingScopes) -join "`n")`n`nResults may be incomplete.",
                'Permission warning',
                [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Warning)
        }
    } elseif ($errText) {
        Set-GuiStatus "Failed: $errText"
    } else {
        Set-GuiStatus 'Completed with no output'
    }

    Update-GuiLog
}

function Start-GuiTaskRun {
    [CmdletBinding()]
    param()

    $st = $script:GuiState
    $ui = $script:GuiUi
    if (-not $st.SelectedTask) { return }
    $task = $st.SelectedTask

    if ($task.Risk -ne 'ReadOnly') {
        $answer = [Windows.MessageBox]::Show(
            "'$($task.Name)' is a $($task.Risk) operation and will modify your tenant.`n`n$($task.Synopsis)`n`nContinue?",
            'Confirm write operation',
            [Windows.MessageBoxButton]::YesNo,
            [Windows.MessageBoxImage]::Warning)
        if ($answer -ne [Windows.MessageBoxResult]::Yes) { return }
    }

    $values = Get-GuiParamValue

    $ui.BtnRun.IsEnabled    = $false
    $ui.BtnCancel.IsEnabled = $true
    Set-GuiStatus "Running '$($task.Name)'..." -Busy

    $st.Shell = [powershell]::Create()
    $st.Shell.Runspace = $st.Worker
    $null = $st.Shell.AddScript({
        param($TaskId, $Values)
        Invoke-M365Task -Id $TaskId -Parameters $Values -Detailed -Force
    }).AddArgument($task.Id).AddArgument($values)

    $st.Handle = $st.Shell.BeginInvoke()

    # Poll on the UI thread instead of blocking it, so the window stays responsive
    # during a tenant-wide scan.
    $st.Timer = New-Object Windows.Threading.DispatcherTimer
    $st.Timer.Interval = [TimeSpan]::FromMilliseconds(300)
    $st.Timer.Add_Tick({
        if (-not $script:GuiState.Handle) { $script:GuiState.Timer.Stop(); return }
        if (-not $script:GuiState.Handle.IsCompleted) { return }
        $script:GuiState.Timer.Stop()
        Complete-GuiTaskRun
    })
    $st.Timer.Start()
}

function Export-GuiResult {
    [CmdletBinding()]
    param([ValidateSet('Csv','Html')][string]$Format)

    $st = $script:GuiState
    if ($st.Results.Count -eq 0) { return }

    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = if ($Format -eq 'Csv') { 'CSV file (*.csv)|*.csv' } else { 'HTML file (*.html)|*.html' }
    $stamp    = Get-Date -Format 'yyyyMMdd-HHmm'
    $safeName = ($st.SelectedTask.Id -replace '[^\w.-]', '_')
    $dialog.FileName = "$safeName-$stamp.$($Format.ToLower())"

    if (-not $dialog.ShowDialog()) { return }

    try {
        if ($Format -eq 'Csv') {
            $st.Results | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
        } else {
            $st.Results | ConvertTo-Html -Title $st.SelectedTask.Name |
                Set-Content -Path $dialog.FileName -Encoding UTF8
        }
        Set-GuiStatus "Exported $($st.Results.Count) row(s) to $($dialog.FileName)"
    } catch {
        [void][Windows.MessageBox]::Show($_.Exception.Message, 'Export failed',
            [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Error)
    }
}
