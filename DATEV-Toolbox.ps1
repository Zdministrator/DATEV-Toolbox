Add-Type -AssemblyName PresentationFramework

Add-Type -Name Win -Namespace Console -MemberDefinition '
[DllImport("kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'
$consolePtr = [Console.Win]::GetConsoleWindow()
# 0 = SW_HIDE, 5 = SW_SHOW
[Console.Win]::ShowWindow($consolePtr, 0)

# Lädt das XAML-Layout für das Hauptfenster der Toolbox
# und definiert die Benutzeroberfläche mit Tabs und Buttons
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
                <RowDefinition Height="Auto" /> <!-- Menüleiste (entfernt) -->
                <RowDefinition Height="100" /> <!-- Log-Ausgabe -->
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
                            <!-- Hier können weitere Download-Buttons eingefügt werden -->
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
                            <Button Name="btnCheckUpdateSettings" Content="Script auf Update prüfen" Height="30" Margin="0,20,0,0" />
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
            </TabControl>
            <TextBox Name="txtLog" Grid.Row="2" IsReadOnly="True" VerticalScrollBarVisibility="Auto" Margin="0,5,0,0" TextWrapping="Wrap" FontSize="11" />
        </Grid>
    </DockPanel>
</Window>
"@

# Initialisiert das Fensterobjekt aus dem XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Setzt die lokale Versionsnummer und ergänzt sie im Fenstertitel
$localVersion = "1.0.6"
$window.Title = "DATEV Toolbox v$localVersion"

# URLs für Online-Update-Prüfung und Script-Download
$versionUrl = "https://raw.githubusercontent.com/Zdministrator/DATEV-Toolbox/main/version.txt"
$scriptUrl = "https://raw.githubusercontent.com/Zdministrator/DATEV-Toolbox/main/DATEV-Toolbox.ps1"

# Initialisiert alle UI-Controls und speichert sie in einer Hashtable für globale Nutzung
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
        "btnOpenDownloadFolder", "btnNgenAll40", "btnLeistungsindex", "menuSettings", "btnCheckUpdateSettings"
    )
    foreach ($name in $controlNames) {
        $global:Controls[$name] = $window.FindName($name)
    }
}

# Nach dem Laden des Fensters Controls initialisieren
Initialize-Controls

# Utility für Logging im Event-Handler (direkt ins Log-Control schreiben)
function Write-LogDirect {
    param(
        [string]$message
    )
    $timestamp = Get-Date -Format 'HH:mm:ss'
    if ($null -ne $global:Controls["txtLog"]) {
        $global:Controls["txtLog"].AppendText("[$timestamp] $message`n")
        $global:Controls["txtLog"].ScrollToEnd()
    }
}

# Schreibt eine Logzeile in das Ausgabefeld (txtLog)
function Write-Log($message) {
    $timestamp = Get-Date -Format 'HH:mm:ss'
    try {
        if ($null -ne $global:Controls["txtLog"]) {
            $global:Controls["txtLog"].AppendText("[$timestamp] $message`n")
            $global:Controls["txtLog"].ScrollToEnd()
        }
    }
    catch {
        # Fehler beim Loggen ignorieren, damit das Script weiterläuft
    }
}

# Fehlerprotokollierung in Datei (kritische Fehler)
function Write-ErrorLog($message) {
    # Ermittle das Scriptverzeichnis
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

# Hilfsfunktion für numerischen Versionsvergleich
function Compare-Version {
    param(
        [string]$v1,
        [string]$v2
    )
    try {
        $ver1 = [Version]$v1
        $ver2 = [Version]$v2
        return $ver1.CompareTo($ver2)
    } catch {
        # Fallback auf Stringvergleich, falls Version-Parsing fehlschlägt
        return [string]::Compare($v1, $v2)
    }
}

# Prüft beim Start, ob eine neue Version des Scripts online verfügbar ist und bietet ggf. ein Update an
function Test-ForUpdate {
    # Prüfe Internetverbindung vor dem Update-Check
    $testConnection = $false
    try {
        $ping = Test-Connection -ComputerName "www.google.com" -Count 1 -Quiet -ErrorAction Stop
        if ($ping) { $testConnection = $true }
    } catch {
        $testConnection = $false
    }
    if (-not $testConnection) {
        Write-Log "Keine Internetverbindung. Update-Check abgebrochen."
        [System.Windows.MessageBox]::Show("Es konnte keine Internetverbindung festgestellt werden. Der Update-Check wird abgebrochen.", "Update-Fehler", 'OK', 'Error')
        return
    }
    try {
        Write-Log "Prüfe auf Updates..."
        $window.Dispatcher.Invoke([action]{}, 'Background')
        # Timeout für Webanfrage setzen
        $webRequest = [System.Net.WebRequest]::Create($versionUrl)
        $webRequest.Timeout = 5000 # 5 Sekunden Timeout
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
                # Ermittelt den Pfad des aktuell laufenden Scripts
                if ($PSCommandPath) {
                    $scriptDir = Split-Path -Parent $PSCommandPath
                    $scriptPath = $PSCommandPath
                }
                elseif ($MyInvocation.MyCommand.Path) {
                    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
                    $scriptPath = $MyInvocation.MyCommand.Path
                }
                else {
                    throw 'Konnte den Skriptpfad nicht ermitteln.'
                }
                # Prüfe Schreibrechte im Scriptverzeichnis
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

                # Zeige nach erfolgreichem Update einen Changelog-Link an (sofern vorhanden)
                Write-Log "Update abgeschlossen. Changelog siehe: https://github.com/Zdministrator/DATEV-Toolbox/releases"
                [System.Windows.MessageBox]::Show("Das Update wurde abgeschlossen. Eine Übersicht der Änderungen finden Sie unter:https://github.com/Zdministrator/DATEV-Toolbox/releases", "Update abgeschlossen", 'OK', 'Information')

                # Cleanup: Warte maximal 10 Sekunden auf das Beenden des Hauptscripts, dann fahre mit dem Update fort
                $updateScript = @"
Start-Sleep -Seconds 2
$timeout = 10
$elapsed = 0
while (Get-Process -Id $PID -ErrorAction SilentlyContinue) {
    Start-Sleep -Milliseconds 500
    $elapsed += 0.5
    if ($elapsed -ge $timeout) { break }
}
# Ersetzen
Move-Item -Path '$newScriptPath' -Destination '$scriptPath' -Force
# Neues Script starten
Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File ""$scriptPath""'
# Update-Script löschen
Remove-Item -Path '$updateScriptPath' -Force
"@
                $updateScript = $updateScript -replace '\$PID', $PID
                try {
                    Set-Content -Path $updateScriptPath -Value $updateScript -Encoding UTF8
                    Write-Log "Update-Script wurde erstellt: $updateScriptPath"
                }
                catch {
                    Write-Log "Fehler beim Erstellen des Update-Scripts: $_"
                }
                [System.Windows.MessageBox]::Show("Das Programm wird für das Update beendet und automatisch neu gestartet.", "Update wird durchgeführt", 'OK', 'Information')
                # Startet das Update-Script und beendet das Hauptscript
                Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$updateScriptPath`"" -WindowStyle Hidden
                exit
            }
            else {
                Write-Log "Update abgebrochen durch Benutzer."
            }
        }
        else {
            Write-Log "Keine neue Version verfügbar."
        }
    }
    catch {
        Write-Log "Fehler beim Update-Check: $($_.Exception.Message)"
        Write-ErrorLog "Fehler beim Update-Check: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Fehler beim Update-Check: $($_.Exception.Message)", "Update-Fehler", 'OK', 'Error')
    }
}

# Registriert einen Button für den Start eines lokalen DATEV-Programms
function Register-ToolButton {
    param (
        [string]$ButtonVar,
        [string]$ExePath,
        [string]$ToolName
    )
    $btn = $Controls[$ButtonVar]
    $exe = $ExePath
    $toolName = $ToolName

    if ([string]::IsNullOrWhiteSpace($exe) -or -not (Test-Path $exe)) {
        $btn.IsEnabled = $false
        $btn.ToolTip = "${toolName} nicht gefunden: $exe"
        Write-Log "$toolName nicht gefunden und Button deaktiviert: $exe"
        return
    }
    else {
        $btn.ToolTip = "$toolName starten"
    }

    $btn.Add_Click({
        try {
            Start-Process -FilePath $exe -ErrorAction Stop
            Write-LogDirect "$toolName gestartet: $exe"
        }
        catch {
            Write-LogDirect "Fehler beim Start von ${toolName}: $($_.Exception.Message)"
        }
    }.GetNewClosure())
}

# Definiert die Zuordnung von Buttons zu lokalen Programmen
$toolButtons = @(
    @{ Button = "btnInstallationsmanager"; Exe = "$env:DATEVPP\PROGRAMM\INSTALL\DvInesInstMan.exe"; Name = "Installationsmanager" },
    @{ Button = "btnServicetool"; Exe = "$env:DATEVPP\PROGRAMM\SRVTOOL\Srvtool.exe"; Name = "Servicetool" },
    @{ Button = "btnKonfigDBTools"; Exe = "$env:DATEVPP\PROGRAMM\B0001502\cdbtool.exe"; Name = "KonfigDB-Tools" },
    @{ Button = "btnEOAufgabenplanung"; Exe = "$env:DATEVPP\PROGRAMM\I0000085\EOControl.exe"; Name = "EO Aufgabenplanung" },
    @{ Button = "btnEODBconfig"; Exe = "$env:DATEVPP\PROGRAMM\EODB\EODBConfig.exe"; Name = "EODBconfig" },
    @{ Button = "btnArbeitsplatz"; Exe = "$env:DATEVPP\PROGRAMM\K0005000\Arbeitsplatz.exe"; Name = "DATEV-Arbeitsplatz" },
    @{ Button = "btnNgenAll40"; Exe = "$env:DATEVPP\Programm\B0001508\ngenall40.cmd"; Name = "NGENALL 4.0" }
)

# Registriert Event-Handler für alle Tool-Buttons
foreach ($entry in $toolButtons) {
    Register-ToolButton -ButtonVar $entry.Button -ExePath $entry.Exe -ToolName $entry.Name
}

# Registriert einen Button für das Öffnen eines Weblinks
function Register-WebLinkHandler {
    param (
        [System.Windows.Controls.Button]$Button,
        [string]$Name,
        [string]$Url
    )
    $Button.Add_Click({
        $btnContent = $Button.Content
        Write-LogDirect "Öffne $btnContent..."
        # Webverbindung direkt im Handler prüfen (HEAD-Request)
        $reachable = $false
        try {
            $request = [System.Net.WebRequest]::Create($Url)
            $request.Method = "HEAD"
            $request.Timeout = 5000
            $response = $request.GetResponse()
            $response.Close()
            $reachable = $true
        }
        catch {
            $reachable = $false
        }
        if (-not $reachable) {
            Write-LogDirect "Keine Verbindung zu $Url möglich – $btnContent wird nicht geöffnet."
            return
        }
        try {
            Start-Process "explorer.exe" $Url
            Write-LogDirect "$btnContent geöffnet."
        }
        catch {
            Write-LogDirect ("Fehler beim Öffnen von {0}: {1}" -f $btnContent, $_)
        }
    }.GetNewClosure())
}

# Definiert die Zuordnung von Cloud-Buttons zu Weblinks
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

# Registriert Event-Handler für alle Cloud/Weblink-Buttons
foreach ($entry in $cloudButtons) {
    $btn = $Controls[$entry.Name]
    if ($btn) {
        Register-WebLinkHandler -Button $btn -Name $entry.Name -Url $entry.Url
    }
}

# Gibt den Download-Ordner für die Toolbox zurück
function Get-DownloadFolder {
    $downloads = [Environment]::GetFolderPath('UserProfile')
    $targetDir = Join-Path $downloads "Downloads"
    $targetDir = Join-Path $targetDir "DATEV-Toolbox"
    if (-not (Test-Path $targetDir)) {
        New-Item -Path $targetDir -ItemType Directory | Out-Null
    }
    return $targetDir
}

# Lädt eine Datei von einer angegebenen URL in den Download-Ordner
function Get-DatevFile {
    param(
        [string]$Url,
        [string]$FileName
    )
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
        $window.Dispatcher.Invoke([action]{}, 'Background') # UI-Update erzwingen
        # Timeout für Webanfrage setzen
        $webRequest = [System.Net.WebRequest]::Create($Url)
        $webRequest.Timeout = 10000 # 10 Sekunden Timeout
        $response = $webRequest.GetResponse()
        $stream = $response.GetResponseStream()
        $fileStream = [System.IO.File]::Create($targetFile)
        $stream.CopyTo($fileStream)
        $fileStream.Close()
        $stream.Close()
        $response.Close()
        Write-Log "Download abgeschlossen: $targetFile"
        [System.Windows.MessageBox]::Show("Download abgeschlossen: $targetFile", "Download", 'OK', 'Information')
    }
    catch {
        Write-Log "Fehler beim Download: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Fehler beim Download: $($_.Exception.Message)", "Fehler", 'OK', 'Error')
    }
}

# Registriert einen Button für eine beliebige Aktion (z. B. Download, Ordner öffnen)
function Register-ButtonAction {
    param(
        [Parameter(Mandatory)][System.Windows.Controls.Button]$Button,
        [Parameter(Mandatory)][scriptblock]$Action
    )
    if ($Button) {
        $Button.Add_Click($Action)
    }
}

# Registriert Event-Handler für alle Download-Buttons
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

# Registriert Event-Handler für das Öffnen des Download-Ordners
Register-ButtonAction -Button $Controls["btnOpenDownloadFolder"] -Action {
    $targetDir = Get-DownloadFolder
    Write-Log "Öffne Download-Ordner..."
    Start-Process explorer.exe $targetDir
}

# Registriert Event-Handler für den Leistungsindex-Button
if ($global:Controls["btnLeistungsindex"]) {
    $global:Controls["btnLeistungsindex"].Add_Click({
        $exe = "$env:DATEVPP\PROGRAMM\RWAPPLIC\irw.exe"
        Start-Process -FilePath $exe -ArgumentList "-ap:PerfIndex -d:IRW20011 -c" -Wait
        Start-Process -FilePath $exe -ArgumentList "-ap:PerfIndex -d:IRW20011" -Wait
    })
}

# Registriert Event-Handler für den Button 'btnCheckUpdateSettings'
if ($global:Controls["btnCheckUpdateSettings"]) {
    $global:Controls["btnCheckUpdateSettings"].Add_Click({
        Test-ForUpdate
    })
}

# Tooltips für alle Buttons ergänzen
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

# Prüft nach dem Laden des Fensters auf Updates
Test-ForUpdate

# Zeigt das Fenster an und startet die Event-Loop
$window.ShowDialog() | Out-Null
