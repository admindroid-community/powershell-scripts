function Show-M365Console {
    <#
        .SYNOPSIS
        Opens the Microsoft 365 Engineering Console (WPF GUI).

        .DESCRIPTION
        A task-driven console over the toolkit's registry. Every registered task appears
        automatically, grouped by category, with a parameter form generated from its
        declared parameters - adding a task requires no GUI code changes.

        Architecture: all Microsoft 365 work happens in one long-lived background
        runspace. Session state (Connect-MgGraph, Connect-ExchangeOnline) is per-runspace,
        so the connection and every task run share that single worker. The UI thread only
        marshals results, which keeps the window responsive during a tenant-wide scan.

        Requires Windows - WPF is not available on Linux or macOS.

        .EXAMPLE
        Show-M365Console
    #>
    [CmdletBinding()]
    param()

    # -------------------------------------------------------------- guards
    if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
        throw "Show-M365Console requires Windows - WPF is not available on this platform. Use Invoke-M365Task from the command line instead."
    }

    try {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml -ErrorAction Stop
    } catch {
        throw "Failed to load WPF assemblies: $($_.Exception.Message)"
    }

    $script:GuiRoot = Join-Path $script:ModuleRoot 'GUI'
    $modulePath     = Join-Path $script:ModuleRoot 'M365Toolkit.psd1'

    $window = Import-M365Xaml -Path (Join-Path $script:GuiRoot 'MainWindow.xaml')
    $script:GuiWindow = $window

    # Bind every x:Name so handlers can reach controls by name.
    $ui = @{}
    foreach ($name in @(
        'ConnDot','ConnText','ConnDetail','BtnConnect','BtnDisconnect',
        'TxtSearch','CmbCategory','TreeTasks',
        'LblTaskName','RiskBadge','LblRisk','LblSynopsis','LblReplaces','LblScopes',
        'PanelParams','BtnRun','BtnCancel',
        'LblRowCount','BtnExportCsv','BtnExportHtml','BtnCopy','GridResults','GridLog',
        'LblStatus','Progress','TabsOutput')) {

        $control = $window.FindName($name)
        if ($null -eq $control) { throw "MainWindow.xaml is missing a control named '$name'." }
        $ui[$name] = $control
    }
    $script:GuiUi = $ui

    # --------------------------------------------------------------- state
    $script:GuiState = [pscustomobject]@{
        Worker        = $null
        Shell         = $null
        Handle        = $null
        Timer         = $null
        SelectedTask  = $null
        ParamControls = @{}
        Results       = @()
        Connected     = $false
    }

    # ----------------------------------------------------- worker runspace
    $worker = [runspacefactory]::CreateRunspace()
    $worker.ApartmentState = 'STA'
    $worker.ThreadOptions  = 'ReuseThread'
    $worker.Open()
    $worker.SessionStateProxy.SetVariable('ToolkitModulePath', $modulePath)
    $script:GuiState.Worker = $worker

    $init = [powershell]::Create()
    $init.Runspace = $worker
    $null = $init.AddScript({ Import-Module $ToolkitModulePath -Force -ErrorAction Stop })
    $null = $init.Invoke()
    if ($init.Streams.Error.Count -gt 0) {
        $msg = ($init.Streams.Error | ForEach-Object { $_.ToString() }) -join "`n"
        $init.Dispose(); $worker.Dispose()
        throw "Worker runspace failed to import M365Toolkit:`n$msg"
    }
    $init.Dispose()

    # -------------------------------------------------------- task catalog
    $script:GuiAllTasks = @(Invoke-InWorker -Script {
        Get-M365Task | Select-Object Id, Name, Category, Synopsis, Risk, Replaces, Notes, Scopes, Service, Parameters
    })

    if ($script:GuiAllTasks.Count -eq 0) {
        $worker.Dispose()
        throw "No tasks are registered - check that the Tasks folder loaded correctly."
    }

    $ui.CmbCategory.ItemsSource = @('All categories') +
        @($script:GuiAllTasks | Select-Object -ExpandProperty Category -Unique | Sort-Object)
    $ui.CmbCategory.SelectedIndex = 0

    # ------------------------------------------------------------ handlers
    $ui.TxtSearch.Add_TextChanged({ Build-GuiTree })
    $ui.CmbCategory.Add_SelectionChanged({ Build-GuiTree })

    $ui.TreeTasks.Add_SelectedItemChanged({
        $item = $script:GuiUi.TreeTasks.SelectedItem
        if (-not $item -or -not $item.Tag) { return }

        $task = $item.Tag
        $script:GuiState.SelectedTask = $task
        $u = $script:GuiUi

        $u.LblTaskName.Text = $task.Name
        $u.LblSynopsis.Text = $task.Synopsis
        $u.LblRisk.Text     = $task.Risk
        $u.RiskBadge.Visibility = 'Visible'
        $u.RiskBadge.Background = switch ($task.Risk) {
            'Destructive' { [Windows.Media.Brushes]::Firebrick }
            'Write'       { [Windows.Media.Brushes]::DarkOrange }
            default       { [Windows.Media.Brushes]::SeaGreen }
        }

        if ($task.Replaces) {
            $u.LblReplaces.Text = "Replaces: $($task.Replaces)"
            $u.LblReplaces.Visibility = 'Visible'
        } else {
            $u.LblReplaces.Visibility = 'Collapsed'
        }

        $svc    = (@($task.Service) -join ', ')
        $scopes = if (@($task.Scopes).Count -gt 0) { '  |  Scopes: ' + (@($task.Scopes) -join ', ') } else { '' }
        $u.LblScopes.Text = "Service: $svc$scopes"

        Build-GuiParamForm -Task $task
        $u.BtnRun.IsEnabled = $script:GuiState.Connected
    })

    $ui.BtnRun.Add_Click({ Start-GuiTaskRun })

    $ui.BtnCancel.Add_Click({
        $st = $script:GuiState
        if ($st.Shell) {
            try { $st.Shell.Stop() } catch { }
            if ($st.Timer) { $st.Timer.Stop() }
            Set-GuiStatus 'Cancelled'
            $script:GuiUi.BtnCancel.IsEnabled = $false
            $script:GuiUi.BtnRun.IsEnabled    = $true
        }
    })

    $ui.BtnExportCsv.Add_Click({  Export-GuiResult -Format Csv })
    $ui.BtnExportHtml.Add_Click({ Export-GuiResult -Format Html })

    $ui.BtnCopy.Add_Click({
        $st = $script:GuiState
        if ($st.Results.Count -eq 0) { return }
        try {
            [Windows.Clipboard]::SetText((($st.Results | ConvertTo-Csv -NoTypeInformation) -join "`r`n"))
            Set-GuiStatus "Copied $($st.Results.Count) row(s) to the clipboard"
        } catch {
            Set-GuiStatus "Clipboard copy failed: $($_.Exception.Message)"
        }
    })

    $ui.BtnConnect.Add_Click({
        $dialog = Import-M365Xaml -Path (Join-Path $script:GuiRoot 'ConnectDialog.xaml')
        $dialog.Owner = $script:GuiWindow

        $d = @{}
        foreach ($n in 'ChkGraph','ChkExo','ChkPnp','ChkTeams','RadInteractive','RadCert',
                       'TxtTenantId','TxtClientId','TxtThumbprint','TxtOrganization',
                       'TxtSharePointUrl','ChkInstall','BtnOk','BtnCancel') {
            $d[$n] = $dialog.FindName($n)
        }

        $d.BtnCancel.Add_Click({ $dialog.DialogResult = $false; $dialog.Close() }.GetNewClosure())
        $d.BtnOk.Add_Click({     $dialog.DialogResult = $true;  $dialog.Close() }.GetNewClosure())

        if (-not $dialog.ShowDialog()) { return }

        $services = @()
        if ($d.ChkGraph.IsChecked) { $services += 'Graph' }
        if ($d.ChkExo.IsChecked)   { $services += 'ExchangeOnline' }
        if ($d.ChkPnp.IsChecked)   { $services += 'PnP' }
        if ($d.ChkTeams.IsChecked) { $services += 'Teams' }

        if ($services.Count -eq 0) {
            [void][Windows.MessageBox]::Show('Select at least one service.', 'Nothing to connect',
                [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Information)
            return
        }

        $connectArgs = @{
            Service               = $services
            InstallMissingModules = [bool]$d.ChkInstall.IsChecked
        }
        if ($d.RadCert.IsChecked) {
            $connectArgs['TenantId']              = $d.TxtTenantId.Text.Trim()
            $connectArgs['ClientId']              = $d.TxtClientId.Text.Trim()
            $connectArgs['CertificateThumbprint'] = $d.TxtThumbprint.Text.Trim()
        }
        if ($d.TxtOrganization.Text.Trim())  { $connectArgs['Organization']  = $d.TxtOrganization.Text.Trim() }
        if ($d.TxtSharePointUrl.Text.Trim()) { $connectArgs['SharePointUrl'] = $d.TxtSharePointUrl.Text.Trim() }

        Set-GuiStatus "Connecting to $($services -join ', ')..." -Busy
        $script:GuiWindow.Cursor = [Windows.Input.Cursors]::Wait

        $err = Invoke-InWorker -Arguments @($connectArgs) -Script {
            param($ConnectArgs)
            try   { Connect-M365Toolkit @ConnectArgs | Out-Null; return $null }
            catch { return $_.Exception.Message }
        }
        $script:GuiWindow.Cursor = $null

        $errMsg = @($err | Where-Object { $_ }) | Select-Object -First 1
        if ($errMsg) {
            Set-GuiStatus 'Connection failed'
            [void][Windows.MessageBox]::Show($errMsg, 'Connection failed',
                [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Error)
        } else {
            Set-GuiStatus 'Connected'
        }

        Update-GuiConnection
        Update-GuiLog
    })

    $ui.BtnDisconnect.Add_Click({
        $null = Invoke-InWorker -Script { Disconnect-M365Toolkit -Confirm:$false }
        Update-GuiConnection
        Update-GuiLog
        Set-GuiStatus 'Disconnected'
    })

    $window.Add_Closed({
        $st = $script:GuiState
        if ($st.Timer)  { try { $st.Timer.Stop() } catch { } }
        if ($st.Shell)  { try { $st.Shell.Dispose() } catch { } }
        if ($st.Worker) { try { $st.Worker.Close(); $st.Worker.Dispose() } catch { } }
    })

    # -------------------------------------------------------------- launch
    Build-GuiTree
    Update-GuiConnection
    Set-GuiStatus "Ready - $($script:GuiAllTasks.Count) task(s) available. Connect to begin."

    $null = $window.ShowDialog()
}

function Import-M365Xaml {
    <#
        .SYNOPSIS
        Loads a XAML file into a WPF object tree.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "XAML not found: $Path" }

    try {
        $xml = [xml](Get-Content -LiteralPath $Path -Raw)
    } catch {
        throw "XAML file '$Path' is not well-formed XML: $($_.Exception.Message)"
    }

    $reader = New-Object System.Xml.XmlNodeReader $xml
    return [Windows.Markup.XamlReader]::Load($reader)
}
