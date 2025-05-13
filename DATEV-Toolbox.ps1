#region Administrator- und Sicherheits-Setup
# Prüft, ob das Skript mit Administratorrechten läuft und startet es ggf. neu mit erhöhten Rechten.
# Aktiviert TLS 1.2 für Webanfragen und blendet das PowerShell-Konsolenfenster aus.
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("Dieses Skript benötigt Administratorrechte. Es wird jetzt mit erhöhten Rechten neu gestartet.", "Administratorrechte erforderlich", 'OK', 'Warning')
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $psi.Verb = 'runas'
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName PresentationFramework
Add-Type -Name Win -Namespace Console -MemberDefinition '
[DllImport("kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'
$consolePtr = [Console.Win]::GetConsoleWindow()
[Console.Win]::ShowWindow($consolePtr, 0)
#endregion

#region Logging-Funktionen
# Funktionen für die Protokollierung von Aktionen und Fehlern im Log-Feld und in einer Datei.
function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format 'HH:mm:ss'
    try {
        if ($null -ne $global:Controls["txtLog"]) {
            $global:Controls["txtLog"].AppendText("[$timestamp] $message`n")
            $global:Controls["txtLog"].ScrollToEnd()
        }
    } catch {}
}
function Write-LogDirect {
    param([string]$message)
    $timestamp = Get-Date -Format 'HH:mm:ss'
    if ($null -ne $global:Controls["txtLog"]) {
        $global:Controls["txtLog"].AppendText("[$timestamp] $message`n")
        $global:Controls["txtLog"].ScrollToEnd()
    }
}
function Write-ErrorLog($message) {
    if ($PSCommandPath) {
        $logDir = Split-Path -Parent $PSCommandPath
    } elseif ($MyInvocation.MyCommand.Path) {
        $logDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        $logDir = $PSScriptRoot
    }
    $logPath = Join-Path $logDir 'DATEV-Toolbox-Fehler.log'
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $logPath -Value "[$timestamp] $message"
}
#endregion

#region UI-Initialisierung
# Lädt das XAML-Layout für das Hauptfenster und initialisiert das WPF-Fensterobjekt.
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="DATEV Toolbox" MinHeight="400" Width="400" ResizeMode="CanMinimize">
    <Window.Resources>
        <!-- Ressourcen können hier definiert werden -->
    </Window.Resources>
    <DockPanel>
        <Grid Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="*" />
                <RowDefinition Height="Auto" />
                <RowDefinition Height="100" />
                <RowDefinition Height="Auto" /> <!-- Fußzeile -->
            </Grid.RowDefinitions>
            <TabControl Grid.Row="0" Margin="0,0,0,0" VerticalAlignment="Stretch">
                <TabItem Header="DATEV Tools">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Orientation="Vertical" Margin="10">
                            <Label Content="Programme" FontWeight="Bold" Margin="5"/>
                            <Button Name="btnArbeitsplatz" Content="DATEV-Arbeitsplatz" Height="30" Margin="5"/>
                            <Button Name="btnInstallationsmanager" Content="Installationsmanager" Height="30" Margin="5"/>
                            <Button Name="btnServicetool" Content="Servicetool" Height="30" Margin="5"/>
                            <Label Content="Tools" FontWeight="Bold" Margin="5"/>
                            <Button Name="btnKonfigDBTools" Content="KonfigDB-Tools" Height="30" Margin="5"/>
                            <Button Name="btnEODBconfig" Content="EODBconfig" Height="30" Margin="5"/>
                            <Button Name="btnEOAufgabenplanung" Content="EO Aufgabenplanung" Height="30" Margin="5"/>
                            <Label Content="Performance" FontWeight="Bold" Margin="5"/>
                            <Button Name="btnNgenAll40" Content="Native Images erzwingen" Height="30" Margin="5" />
                            <Button Name="btnLeistungsindex" Content="Leistungsindex" Height="30" Margin="5" />
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
                <TabItem Header="Cloud Anwendungen">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Orientation="Vertical" Margin="10">
                            <Label Content="Hilfe und Support" FontWeight="Bold" Margin="5"/>
                            <Button Name="btnDATEVHilfeCenter"          Content="DATEV Hilfe Center"          Height="30" Margin="5"/>
                            <Button Name="btnServicekontaktuebersicht" Content="Servicekontakte"      Height="30" Margin="5"/>
                            <Label Content="Cloud" FontWeight="Bold" Margin="5"/>                    
                            <Button Name="btnMyDATEVPortal"             Content="MyDATEV Portal"              Height="30" Margin="5"/>
                            <Button Name="btnDATEVUnternehmenOnline"    Content="DATEV Unternehmen Online"   Height="30" Margin="5"/>
                            <Button Name="btnLogistikauftragOnline"      Content="Logistikauftrag Online"     Height="30" Margin="5"/>
                            <Button Name="btnLizenzverwaltungOnline"     Content="Lizenzverwaltung Online"    Height="30" Margin="5"/>
                            <Button Name="btnDATEVRechteraumOnline"      Content="DATEV Rechteraum online"    Height="30" Margin="5"/>
                            <Button Name="btnDATEVRechteverwaltungOnline" Content="DATEV Rechteverwaltung online" Height="30" Margin="5"/>
                            <Label Content="Verwaltung und Technik" FontWeight="Bold" Margin="5"/>
                            <Button Name="btnSmartLoginAdministration"   Content="SmartLogin Administration"  Height="30" Margin="5"/>
                            <Button Name="btnMyDATEVBestandsmanagement"   Content="MyDATEV Bestandsmanagement" Height="30" Margin="5"/>
                            <Button Name="btnWeitereCloudAnwendungen"    Content="Weitere Cloud Anwendungen"  Height="30" Margin="5"/>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
                <TabItem Header="Downloads">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Orientation="Vertical" Margin="10">
                            <Label Content="Downloads von externer Liste" FontWeight="Bold" Margin="5"/>
                            <ComboBox Name="cmbDynamicDownloads" Width="300" Margin="0,0,0,0" />
                            <Button Name="btnStartDynamicDownload" Content="Download starten" Height="30" Width="150" Margin="0,10,0,0" />
                            <Label Content="Downloads von datev.de" FontWeight="Bold" Margin="5"/>
                            <Button Name="btnDatevDownloadbereich" Content="DATEV Downloadbereich" Height="30" Margin="5"/>
                            <Button Name="btnDatevSmartDocs" Content="DATEV Smart Docs" Height="30" Margin="5"/>
                            <Button Name="btnDatentraegerDownloadPortal" Content="Datenträger Download Portal" Height="30" Margin="5"/>
                            <Label Content="Direkt Downloads" FontWeight="Bold" Margin="5"/>
                            <Button Name="btnDownloadSicherheitspaketCompact" Content="Sicherheitspaket compact" Height="30" Margin="5"/>
                            <Button Name="btnDownloadFernbetreuungOnline" Content="Fernbetreuung Online" Height="30" Margin="5"/>
                            <Button Name="btnDownloadBelegtransfer" Content="Belegtransfer V. 5.46" Height="30" Margin="5"/>
                            <Button Name="btnDownloadServerprep" Content="Serverprep" Height="30" Margin="5"/>
                            <Button Name="btnDownloadDeinstallationsnacharbeiten" Content="Deinstallationsnacharbeiten" Height="30" Margin="5"/>
                            <Button Name="btnOpenDownloadFolder" Height="32" Width="32" Margin="10,20,10,10" ToolTip="Download-Ordner öffnen">
                                <Button.Content>
                                    <Viewbox Width="24" Height="24">
                                        <Canvas Width="24" Height="24">
                                            <Path Data="M3,6 L21,6 L21,20 L3,20 Z M3,6 L12,2 L21,6" Stroke="Black" StrokeThickness="1.5" Fill="#FFC107"/>
                                        </Canvas>
                                    </Viewbox>
                                </Button.Content>
                            </Button>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
                <TabItem Header="Einstellungen">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Orientation="Vertical" Margin="10">
                            <TextBlock Text='Einstellungen (Platzhalter)' FontWeight='Bold' FontSize='14' Margin='0,0,0,10'/>
                            <TextBlock Text='Hier können Einstellungen ergänzt werden.' />
                            <Button Name="btnCheckUpdateSettings" Content="Script auf Update prüfen" Height="30" Margin="5" />
                            <Button Name="btnUpdateDownloadList" Content="Update Download-Liste von Github" Height="30" Margin="5" />
                            <TextBlock Name="txtDownloadListMeta" FontSize="10" Foreground="Gray" Margin="0,10,0,0" />
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
            </TabControl>
            <TextBox Name="txtLog" Grid.Row="2" IsReadOnly="True" VerticalScrollBarVisibility="Auto" Margin="0,5,0,0" TextWrapping="Wrap" FontSize="11" />
            <TextBlock Grid.Row="3" Text="Norman Zamponi | HEES GmbH | © 2025" HorizontalAlignment="Center" FontSize="11" Foreground="Gray" Margin="0,5,0,0" />
        </Grid>
    </DockPanel>
</Window>
"@
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
$localVersion = "1.0.8"
$window.Title = "DATEV Toolbox v$localVersion"
#endregion

#region Controls-Initialisierung
# Sammelt alle relevanten UI-Controls in einer Hashtable für den globalen Zugriff.
function Initialize-Controls {
    $global:Controls = @{
        "txtLog" = $window.FindName("txtLog")
    }
    $controlNames = @(
        "btnArbeitsplatz", "btnInstallationsmanager", "btnServicetool", "btnKonfigDBTools", "btnEOAufgabenplanung", "btnEODBconfig",
        "btnDATEVHilfeCenter", "btnServicekontaktuebersicht", "btnMyDATEVPortal", "btnDATEVUnternehmenOnline", "btnLogistikauftragOnline",
        "btnLizenzverwaltungOnline", "btnDATEVRechteraumOnline", "btnDATEVRechteverwaltungOnline", "btnSmartLoginAdministration",
        "btnMyDATEVBestandsmanagement", "btnWeitereCloudAnwendungen", "btnDatevDownloadbereich", "btnDatevSmartDocs", "btnDatentraegerDownloadPortal",
        "btnDownloadSicherheitspaketCompact", "btnDownloadFernbetreuungOnline", "btnDownloadBelegtransfer", "btnDownloadServerprep", "btnDownloadDeinstallationsnacharbeiten",
        "btnOpenDownloadFolder", "btnNgenAll40", "btnLeistungsindex", "menuSettings", "btnCheckUpdateSettings", "cmbDynamicDownloads", "btnStartDynamicDownload", "btnUpdateDownloadList", "txtDownloadListMeta"
    )
    foreach ($name in $controlNames) {
        $global:Controls[$name] = $window.FindName($name)
    }
}
Initialize-Controls
#endregion

#region Hilfsfunktionen
# Diverse Hilfsfunktionen, z.B. für Versionsvergleiche.
function Compare-Version {
    param([string]$v1, [string]$v2)
    try {
        $ver1 = [Version]$v1
        $ver2 = [Version]$v2
        return $ver1.CompareTo($ver2)
    } catch {
        return [string]::Compare($v1, $v2)
    }
}
#endregion

#region Download-Funktionen
# Funktionen zum Ermitteln des Download-Ordners und Herunterladen von Dateien mit Fortschritts- und Fehlerbehandlung.
function Get-DownloadFolder {
    $downloads = [Environment]::GetFolderPath('UserProfile')
    $targetDir = Join-Path $downloads "Downloads"
    $targetDir = Join-Path $targetDir "DATEV-Toolbox"
    if (-not (Test-Path $targetDir)) {
        New-Item -Path $targetDir -ItemType Directory | Out-Null
    }
    return $targetDir
}
function Get-DatevFile {
    param([string]$Url, [string]$FileName)
    $targetDir = Get-DownloadFolder
    $targetFile = Join-Path $targetDir $FileName
    if (Test-Path $targetFile) {
        $result = [System.Windows.MessageBox]::Show("Die Datei '$FileName' existiert bereits. Überschreiben?", "Datei existiert", 'YesNo', 'Warning')
        if ($result -ne 'Yes') {
            Write-Log "Download abgebrochen: $FileName existiert bereits."
            return
        }
    }
    try {
        Write-Log "Lade $FileName herunter ..."
        $window.Dispatcher.Invoke([action]{}, 'Background')
        $webRequest = [System.Net.WebRequest]::Create($Url)
        $webRequest.Timeout = 10000
        $response = $webRequest.GetResponse()
        $stream = $response.GetResponseStream()
        $fileStream = [System.IO.File]::Create($targetFile)
        $stream.CopyTo($fileStream)
        $fileStream.Close()
        $stream.Close()
        $response.Close()
        Write-Log "Download abgeschlossen: $targetFile"
        [System.Windows.MessageBox]::Show("Download abgeschlossen: $targetFile", "Download", 'OK', 'Information')
    } catch {
        Write-Log "Fehler beim Download: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Fehler beim Download: $($_.Exception.Message)", "Fehler", 'OK', 'Error')
    }
}
#endregion

#region Dynamische Download-Liste laden und Dropdown befüllen
# Liest die Datei 'downloads.json' ein und befüllt das Dropdown im Download-Tab
$dynamicDownloadsFile = Join-Path (Split-Path $PSCommandPath) 'downloads.json'
$global:DynamicDownloads = @()
if (Test-Path $dynamicDownloadsFile) {
    try {
        $downloadsMeta = Get-Content $dynamicDownloadsFile | ConvertFrom-Json
        $global:DynamicDownloads = $downloadsMeta.Downloads
        $Controls["cmbDynamicDownloads"].Items.Clear()
        foreach ($item in $global:DynamicDownloads) {
            $Controls["cmbDynamicDownloads"].Items.Add($item.Name)
        }
        if ($Controls["cmbDynamicDownloads"].Items.Count -gt 0) {
            $Controls["cmbDynamicDownloads"].SelectedIndex = 0
        }
        # Version und Datum im Tab Einstellungen anzeigen
        if ($downloadsMeta.Version -and $downloadsMeta.Datum) {
            $Controls["txtDownloadListMeta"].Text = "Download-Liste Version: $($downloadsMeta.Version), Datum: $($downloadsMeta.Datum)"
        } else {
            $Controls["txtDownloadListMeta"].Text = "Download-Liste: keine Metadaten gefunden."
        }
    } catch {
        Write-Log "Fehler beim Laden der Download-Liste: $($_.Exception.Message)"
        $Controls["txtDownloadListMeta"].Text = "Fehler beim Laden der Download-Liste."
    }
} else {
    Write-Log "downloads.json nicht gefunden. Das Dropdown bleibt leer."
    $Controls["txtDownloadListMeta"].Text = "downloads.json nicht gefunden."
}

# Event-Handler für den Button 'Download starten' im dynamischen Bereich
Register-ButtonAction -Button $Controls["btnStartDynamicDownload"] -Action {
    $selIndex = $Controls["cmbDynamicDownloads"].SelectedIndex
    if ($selIndex -lt 0 -or $selIndex -ge $global:DynamicDownloads.Count) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie einen Download aus der Liste.", "Hinweis", 'OK', 'Information')
        return
    }
    $entry = $global:DynamicDownloads[$selIndex]
    if ($entry.Url -and $entry.Name) {
        Get-DatevFile -Url $entry.Url -FileName ([System.IO.Path]::GetFileName($entry.Url))
    } else {
        [System.Windows.MessageBox]::Show("Ungültiger Eintrag in der Download-Liste.", "Fehler", 'OK', 'Error')
    }
}

# Event-Handler für den Button 'Download-Liste aktualisieren'
Register-ButtonAction -Button $Controls["btnUpdateDownloadList"] -Action {
    $onlineUrl = "https://raw.githubusercontent.com/Zdministrator/DATEV-Toolbox/refs/heads/main/downloads.json"
    try {
        Write-Log "Lade aktuelle Download-Liste von $onlineUrl ..."
        $window.Dispatcher.Invoke([action]{}, 'Background')
        Invoke-WebRequest -Uri $onlineUrl -OutFile $dynamicDownloadsFile -UseBasicParsing
        Write-Log "Download-Liste erfolgreich heruntergeladen und gespeichert."
        # Nach dem Download die Liste neu laden und Dropdown aktualisieren
        $downloadsMeta = Get-Content $dynamicDownloadsFile | ConvertFrom-Json
        $global:DynamicDownloads = $downloadsMeta.Downloads
        $Controls["cmbDynamicDownloads"].Items.Clear()
        foreach ($item in $global:DynamicDownloads) {
            $Controls["cmbDynamicDownloads"].Items.Add($item.Name)
        }
        if ($Controls["cmbDynamicDownloads"].Items.Count -gt 0) {
            $Controls["cmbDynamicDownloads"].SelectedIndex = 0
        }
        # Version und Datum im Tab Einstellungen anzeigen
        if ($downloadsMeta.Version -and $downloadsMeta.Datum) {
            $Controls["txtDownloadListMeta"].Text = "Download-Liste Version: $($downloadsMeta.Version), Datum: $($downloadsMeta.Datum)"
        } else {
            $Controls["txtDownloadListMeta"].Text = "Download-Liste: keine Metadaten gefunden."
        }
    } catch {
        Write-Log "Fehler beim Herunterladen der Download-Liste: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Fehler beim Herunterladen der Download-Liste: $($_.Exception.Message)", "Fehler", 'OK', 'Error')
        $Controls["txtDownloadListMeta"].Text = "Fehler beim Herunterladen der Download-Liste."
    }
}
#endregion

#region Update-Funktionen
# Funktionen zum Prüfen und Durchführen von Updates des Skripts über GitHub.
$versionUrl = "https://raw.githubusercontent.com/Zdministrator/DATEV-Toolbox/main/version.txt"
$scriptUrl = "https://raw.githubusercontent.com/Zdministrator/DATEV-Toolbox/main/DATEV-Toolbox.ps1"
function Test-ForUpdate {
    $testConnection = $false
    try {
        $ping = Test-Connection -ComputerName "www.google.com" -Count 1 -Quiet -ErrorAction Stop
        if ($ping) { $testConnection = $true }
    } catch { $testConnection = $false }
    if (-not $testConnection) {
        Write-Log "Keine Internetverbindung. Update-Check abgebrochen."
        [System.Windows.MessageBox]::Show("Es konnte keine Internetverbindung festgestellt werden. Der Update-Check wird abgebrochen.", "Update-Fehler", 'OK', 'Error')
        return
    }
    try {
        Write-Log "Prüfe auf Updates..."
        $window.Dispatcher.Invoke([action]{}, 'Background')
        $webRequest = [System.Net.WebRequest]::Create($versionUrl)
        $webRequest.Timeout = 5000
        $response = $webRequest.GetResponse()
        $stream = $response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        $remoteVersion = $reader.ReadToEnd().Trim()
        $reader.Close()
        $response.Close()
        Write-Log "Lokale Version: $localVersion, Online-Version: $remoteVersion"
        $cmp = Compare-Version $localVersion $remoteVersion
        if ($remoteVersion -and ($cmp -lt 0)) {
            Write-Log "Neue Version gefunden: $remoteVersion (aktuell: $localVersion)"
            $result = [System.Windows.MessageBox]::Show("Neue Version ($remoteVersion) verfügbar. Jetzt herunterladen?", "Update verfügbar", 'YesNo', 'Information')
            if ($result -eq 'Yes') {
                if ($PSCommandPath) {
                    $scriptDir = Split-Path -Parent $PSCommandPath
                    $scriptPath = $PSCommandPath
                } elseif ($MyInvocation.MyCommand.Path) {
                    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
                    $scriptPath = $MyInvocation.MyCommand.Path
                } else {
                    throw 'Konnte den Skriptpfad nicht ermitteln.'
                }
                $testFile = Join-Path $scriptDir "_write_test.tmp"
                try {
                    Set-Content -Path $testFile -Value "test" -ErrorAction Stop
                    Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
                } catch {
                    Write-Log "Keine Schreibrechte im Scriptverzeichnis: $scriptDir"
                    [System.Windows.MessageBox]::Show("Das Update kann nicht durchgeführt werden, da keine Schreibrechte im Scriptverzeichnis ($scriptDir) bestehen.", "Update-Fehler", 'OK', 'Error')
                    return
                }
                $newScriptPath = Join-Path $scriptDir "DATEV-Toolbox.ps1.new"
                $updateScriptPath = Join-Path $scriptDir "Update-DATEV-Toolbox.ps1"
                Write-Log "Lade neue Version herunter..."
                $window.Dispatcher.Invoke([action]{}, 'Background')
                Invoke-WebRequest -Uri $scriptUrl -OutFile $newScriptPath -UseBasicParsing
                Write-Log "Neue Version wurde als $newScriptPath gespeichert. Update-Vorgang wird vorbereitet."
                Write-Log "Update abgeschlossen. Changelog siehe: https://github.com/Zdministrator/DATEV-Toolbox/releases"
                [System.Windows.MessageBox]::Show("Das Update wurde abgeschlossen.`nEine Übersicht der Änderungen finden Sie unter:`nhttps://github.com/Zdministrator/DATEV-Toolbox/releases", "Update abgeschlossen", 'OK', 'Information')
                $updateScript = @"
Start-Sleep -Seconds 2
$timeout = 10
$elapsed = 0
while (Get-Process -Id $PID -ErrorAction SilentlyContinue) {
    Start-Sleep -Milliseconds 500
    $elapsed += 0.5
    if ($elapsed -ge $timeout) { break }
}
Move-Item -Path '$newScriptPath' -Destination '$scriptPath' -Force
Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File ""$scriptPath""'
Remove-Item -Path '$updateScriptPath' -Force
"@
                $updateScript = $updateScript -replace '\$PID', $PID
                try {
                    Set-Content -Path $updateScriptPath -Value $updateScript -Encoding UTF8
                    Write-Log "Update-Script wurde erstellt: $updateScriptPath"
                } catch {
                    Write-Log "Fehler beim Erstellen des Update-Scripts: $_"
                }
                [System.Windows.MessageBox]::Show("Das Programm wird für das Update beendet und automatisch neu gestartet.", "Update wird durchgeführt", 'OK', 'Information')
                Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$updateScriptPath`"" -WindowStyle Hidden
                exit
            } else {
                Write-Log "Update abgebrochen durch Benutzer."
            }
        } else {
            Write-Log "Keine neue Version verfügbar."
        }
    } catch {
        Write-Log "Fehler beim Update-Check: $($_.Exception.Message)"
        Write-ErrorLog "Fehler beim Update-Check: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Fehler beim Update-Check: $($_.Exception.Message)", "Update-Fehler", 'OK', 'Error')
    }
}
#endregion

#region Button- und Event-Handler-Registrierung
# Funktionen zur Registrierung von Event-Handlern für Buttons und Weblinks.
function Register-ToolButton {
    param ([string]$ButtonVar, [string]$ExePath, [string]$ToolName)
    $btn = $Controls[$ButtonVar]
    $exe = $ExePath
    $toolName = $ToolName
    if ([string]::IsNullOrWhiteSpace($exe) -or -not (Test-Path $exe)) {
        $btn.IsEnabled = $false
        $btn.ToolTip = "${toolName} nicht gefunden: $exe"
        Write-Log "$toolName nicht gefunden und Button deaktiviert: $exe"
        return
    } else {
        $btn.ToolTip = "$toolName starten"
    }
    $btn.Add_Click({
        try {
            Start-Process -FilePath $exe -ErrorAction Stop
            Write-LogDirect "$toolName gestartet: $exe"
        } catch {
            Write-LogDirect "Fehler beim Start von ${toolName}: $($_.Exception.Message)"
        }
    }.GetNewClosure())
}
function Register-WebLinkHandler {
    param ([System.Windows.Controls.Button]$Button, [string]$Name, [string]$Url)
    $Button.Add_Click({
        $btnContent = $Button.Content
        Write-LogDirect "Öffne $btnContent..."
        $reachable = $false
        try {
            $request = [System.Net.WebRequest]::Create($Url)
            $request.Method = "HEAD"
            $request.Timeout = 5000
            $response = $request.GetResponse()
            $response.Close()
            $reachable = $true
        } catch { $reachable = $false }
        if (-not $reachable) {
            Write-LogDirect "Keine Verbindung zu $Url möglich – $btnContent wird nicht geöffnet."
            return
        }
        try {
            Start-Process "explorer.exe" $Url
            Write-LogDirect "$btnContent geöffnet."
        } catch {
            Write-LogDirect ("Fehler beim Öffnen von {0}: {1}" -f $btnContent, $_)
        }
    }.GetNewClosure())
}
function Register-ButtonAction {
    param([Parameter(Mandatory)][System.Windows.Controls.Button]$Button, [Parameter(Mandatory)][scriptblock]$Action)
    if ($Button) {
        $Button.Add_Click($Action)
    }
}
#endregion

#region Button- und Event-Handler-Zuordnung
# Ordnet die definierten Funktionen den jeweiligen Buttons und Aktionen im UI zu.
$toolButtons = @(
    @{ Button = "btnInstallationsmanager"; Exe = "$env:DATEVPP\PROGRAMM\INSTALL\DvInesInstMan.exe"; Name = "Installationsmanager" },
    @{ Button = "btnServicetool"; Exe = "$env:DATEVPP\PROGRAMM\SRVTOOL\Srvtool.exe"; Name = "Servicetool" },
    @{ Button = "btnKonfigDBTools"; Exe = "$env:DATEVPP\PROGRAMM\B0001502\cdbtool.exe"; Name = "KonfigDB-Tools" },
    @{ Button = "btnEOAufgabenplanung"; Exe = "$env:DATEVPP\PROGRAMM\I0000085\EOControl.exe"; Name = "EO Aufgabenplanung" },
    @{ Button = "btnEODBconfig"; Exe = "$env:DATEVPP\PROGRAMM\EODB\EODBConfig.exe"; Name = "EODBconfig" },
    @{ Button = "btnArbeitsplatz"; Exe = "$env:DATEVPP\PROGRAMM\K0005000\Arbeitsplatz.exe"; Name = "DATEV-Arbeitsplatz" },
    @{ Button = "btnNgenAll40"; Exe = "$env:DATEVPP\Programm\B0001508\ngenall40.cmd"; Name = "NGENALL 4.0" }
)
foreach ($entry in $toolButtons) {
    Register-ToolButton -ButtonVar $entry.Button -ExePath $entry.Exe -ToolName $entry.Name
}

$cloudButtons = @(
    @{ Name = "btnDATEVHilfeCenter"; Url = "https://apps.datev.de/help-center/" },
    @{ Name = "btnServicekontaktuebersicht"; Url = "https://apps.datev.de/servicekontakt-online/contacts" },
    @{ Name = "btnMyDATEVPortal"; Url = "https://apps.datev.de/mydatev" },
    @{ Name = "btnDATEVUnternehmenOnline"; Url = "https://duo.datev.de/" },
    @{ Name = "btnLogistikauftragOnline"; Url = "https://apps.datev.de/lao" },
    @{ Name = "btnLizenzverwaltungOnline"; Url = "https://apps.datev.de/lizenzverwaltung" },
    @{ Name = "btnDATEVRechteraumOnline"; Url = "https://apps.datev.de/rechteraum" },
    @{ Name = "btnDATEVRechteverwaltungOnline"; Url = "https://apps.datev.de/rvo-administration" },
    @{ Name = "btnSmartLoginAdministration"; Url = "https://go.datev.de/smartlogin-administration" },
    @{ Name = "btnMyDATEVBestandsmanagement"; Url = "https://apps.datev.de/mydata/" },
    @{ Name = "btnWeitereCloudAnwendungen"; Url = "https://www.datev.de/web/de/mydatev/datev-cloud-anwendungen/" },
    @{ Name = "btnDatevDownloadbereich"; Url = "https://www.datev.de/download/" },
    @{ Name = "btnDatevSmartDocs"; Url = "https://www.datev.de/web/de/service-und-support/software-bereitstellung/download-bereich/it-loesungen-und-security/datev-smartdocs-skripte-zur-analyse-oder-reparatur/" },
    @{ Name = "btnDatentraegerDownloadPortal"; Url = "https://www.datev.de/web/de/service-und-support/software-bereitstellung/datentraeger-portal/" }
)
foreach ($entry in $cloudButtons) {
    $btn = $Controls[$entry.Name]
    if ($btn) {
        Register-WebLinkHandler -Button $btn -Name $entry.Name -Url $entry.Url
    }
}

Register-ButtonAction -Button $Controls["btnDownloadSicherheitspaketCompact"] -Action {
    Get-DatevFile -Url "https://download.datev.de/download/sipacompact/sipacompact.exe" -FileName "sipacompact.exe"
}
Register-ButtonAction -Button $Controls["btnDownloadFernbetreuungOnline"] -Action {
    Get-DatevFile -Url "https://download.datev.de/download/fbo-kp/datev_fernbetreuung_online.exe" -FileName "datev_fernbetreuung_online.exe"
}
Register-ButtonAction -Button $Controls["btnDownloadBelegtransfer"] -Action {
    Get-DatevFile -Url "https://download.datev.de/download/bedi/belegtransfer546.exe" -FileName "belegtransfer546.exe"
}
Register-ButtonAction -Button $Controls["btnDownloadServerprep"] -Action {
    Get-DatevFile -Url "https://download.datev.de/download/datevitfix/serverprep.exe" -FileName "serverprep.exe"
}
Register-ButtonAction -Button $Controls["btnDownloadDeinstallationsnacharbeiten"] -Action {
    Get-DatevFile -Url "https://download.datev.de/download/deinstallationsnacharbeiten_v311/deinstnacharbeitentool.exe" -FileName "deinstnacharbeitentool.exe"
}

Register-ButtonAction -Button $Controls["btnOpenDownloadFolder"] -Action {
    $targetDir = Get-DownloadFolder
    Write-Log "Öffne Download-Ordner..."
    Start-Process explorer.exe $targetDir
}

if ($global:Controls["btnLeistungsindex"]) {
    $exe = "$env:DATEVPP\PROGRAMM\RWAPPLIC\irw.exe"
    if (-not (Test-Path $exe)) {
        $global:Controls["btnLeistungsindex"].IsEnabled = $false
        $global:Controls["btnLeistungsindex"].ToolTip = "Leistungsindex-Programm nicht gefunden: $exe"
    } else {
        $global:Controls["btnLeistungsindex"].ToolTip = "Startet den Leistungsindex."
        $global:Controls["btnLeistungsindex"].Add_Click({
            Start-Process -FilePath $exe -ArgumentList "-ap:PerfIndex -d:IRW20011 -c" -Wait
            Start-Process -FilePath $exe -ArgumentList "-ap:PerfIndex -d:IRW20011" -Wait
        })
    }
}

if ($global:Controls["btnCheckUpdateSettings"]) {
    $global:Controls["btnCheckUpdateSettings"].Add_Click({
        Test-ForUpdate
    })
}

$Controls["btnArbeitsplatz"].ToolTip = "Startet den DATEV-Arbeitsplatz."
$Controls["btnInstallationsmanager"].ToolTip = "Startet den DATEV-Installationsmanager."
$Controls["btnServicetool"].ToolTip = "Startet das DATEV-Servicetool."
$Controls["btnKonfigDBTools"].ToolTip = "Startet die KonfigDB-Tools."
$Controls["btnEOAufgabenplanung"].ToolTip = "Startet die EO Aufgabenplanung."
$Controls["btnEODBconfig"].ToolTip = "Startet die EODBconfig."
$Controls["btnDownloadSicherheitspaketCompact"].ToolTip = "Lädt das Sicherheitspaket compact herunter."
$Controls["btnDownloadFernbetreuungOnline"].ToolTip = "Lädt die Fernbetreuung Online herunter."
$Controls["btnDownloadBelegtransfer"].ToolTip = "Lädt Belegtransfer V. 5.46 herunter."
$Controls["btnDownloadServerprep"].ToolTip = "Lädt Serverprep herunter."
$Controls["btnDownloadDeinstallationsnacharbeiten"].ToolTip = "Lädt das Deinstallationsnacharbeiten-Tool herunter."
$Controls["btnCheckUpdateSettings"].ToolTip = "Prüft das Script auf Updates."
#endregion

# Startet die automatische Update-Prüfung und zeigt das Hauptfenster an.
Test-ForUpdate
$window.ShowDialog() | Out-Null
