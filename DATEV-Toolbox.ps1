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
    param(
        [string]$message,
        [switch]$IsError
    )
    $timestamp = Get-Date -Format 'HH:mm:ss'
    try {
        if ($null -ne $global:Controls["txtLog"]) {
            $global:Controls["txtLog"].AppendText("[$timestamp] $message`n")
            $global:Controls["txtLog"].ScrollToEnd()
        }
        if ($IsError) { Write-ErrorLog $message }
    }
    catch {}
}

function Write-LogDirect {
    param([string]$message)
    $timestamp = Get-Date -Format 'HH:mm:ss'
    if ($null -ne $global:Controls["txtLog"]) {
        $global:Controls["txtLog"].AppendText("[$timestamp] $message`n")
        $global:Controls["txtLog"].ScrollToEnd()
    }
}

function Write-ErrorLog {
    param([string]$message)
    $settingsFile = Get-SettingsFilePath
    $logDir = Split-Path -Parent $settingsFile
    $logPath = Join-Path $logDir 'DATEV-Toolbox-Fehler.log'
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $logPath -Value "[$timestamp] $message"
}

function Register-ButtonAction {
    param([Parameter(Mandatory)][System.Windows.Controls.Button]$Button, [Parameter(Mandatory)][scriptblock]$Action)
    if ($Button) {
        $Button.Add_Click($Action)
    }
}
#endregion

#region Settings-Funktionen
# Standardwerte für Einstellungen zentral definieren
$defaultSettings = @{
    # Beispiel-Keys, nach Bedarf anpassen/erweitern
    "Sprache" = "de"
    # ...weitere Standardwerte...
}

function Get-SettingsFilePath {
    $settingsDir = Join-Path $env:APPDATA 'DATEV-Toolbox'
    if (-not (Test-Path $settingsDir)) {
        New-Item -Path $settingsDir -ItemType Directory | Out-Null
        Write-Log "Settings-Verzeichnis neu angelegt: $settingsDir"
    }
    return (Join-Path $settingsDir 'settings.json')
}

function Import-Settings {
    $settingsFile = Get-SettingsFilePath
    Write-Log "Lade Einstellungen ..."
    if (Test-Path $settingsFile) {
        try {
            $raw = Get-Content $settingsFile -Raw | ConvertFrom-Json
            $global:Settings = @{}
            foreach ($p in $raw.PSObject.Properties) {
                $global:Settings[$p.Name] = $p.Value
            }
            # Fehlende Keys ergänzen
            foreach ($key in $defaultSettings.Keys) {
                if (-not $global:Settings.ContainsKey($key)) {
                    $global:Settings[$key] = $defaultSettings[$key]
                }
            }
            Write-Log "Einstellungen geladen aus $settingsFile."
        }
        catch {
            $global:Settings = $defaultSettings.Clone()
            Write-Log "Fehler beim Laden der Einstellungen. Verwende Standardwerte. Fehlerdetails: $($_.Exception.Message) | Datei: $settingsFile | Inhalt: $(Get-Content $settingsFile -Raw)" -IsError
        }
    }
    else {
        $global:Settings = $defaultSettings.Clone()
        Write-Log "Keine Einstellungen gefunden. Verwende Standardwerte."
    }
}

function Save-Settings {
    $settingsFile = Get-SettingsFilePath
    try {
        if (-not $global:Settings) { $global:Settings = $defaultSettings.Clone() }
        $global:Settings | ConvertTo-Json -Depth 5 | Set-Content $settingsFile -Encoding UTF8
        Write-Log "Einstellungen gespeichert nach $settingsFile."
    }
    catch {
        Write-Log "Fehler beim Speichern der Einstellungen: $($_.Exception.Message)" -IsError
    }
}
#endregion

Import-Settings

#region UI-Initialisierung
Write-Log "Initialisiere Benutzeroberfläche ..."
# Lädt das XAML-Layout für das Hauptfenster und initialisiert das WPF-Fensterobjekt.
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="DATEV Toolbox" MinHeight="400" Width="480" ResizeMode="CanMinimize">
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
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Anstehende Update Termine" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <StackPanel Name="spUpdateDates" />
                                    <Button Name="btnUpdateDates" Content="Termine aktualisieren" Height="22" Margin="5" ToolTip="Lädt die aktuellen Update-Termine aus der ICS-Datei."/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Programme" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <Button Name="btnArbeitsplatz" Content="DATEV-Arbeitsplatz" Height="30" Margin="5" ToolTip="Startet den DATEV-Arbeitsplatz."/>
                                    <Button Name="btnInstallationsmanager" Content="Installationsmanager" Height="30" Margin="5" ToolTip="Startet den DATEV-Installationsmanager."/>
                                    <Button Name="btnServicetool" Content="Servicetool" Height="30" Margin="5" ToolTip="Startet das DATEV-Servicetool."/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Tools" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <Button Name="btnKonfigDBTools" Content="KonfigDB-Tools" Height="30" Margin="5" ToolTip="Startet die KonfigDB-Tools."/>
                                    <Button Name="btnEODBconfig" Content="EODBconfig" Height="30" Margin="5" ToolTip="Startet die EODBconfig."/>
                                    <Button Name="btnEOAufgabenplanung" Content="EO Aufgabenplanung" Height="30" Margin="5" ToolTip="Startet die EO Aufgabenplanung."/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Performance" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <Button Name="btnNgenAll40" Content="Native Images erzwingen" Height="30" Margin="5" ToolTip="Erzwingt die Erstellung von .NET Native Images (ngenall40)."/>
                                    <Button Name="btnLeistungsindex" Content="Leistungsindex" Height="30" Margin="5" ToolTip="Startet den Leistungsindex."/>
                                </StackPanel>
                            </GroupBox>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <TabItem Header="Cloud Anwendungen">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Orientation="Vertical" Margin="10">
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Hilfe und Support" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <Button Name="btnDATEVHilfeCenter"          Content="DATEV Hilfe Center"          Height="30" Margin="5" ToolTip="Öffnet das DATEV Hilfe Center (Web)."/>
                                    <Button Name="btnServicekontaktuebersicht" Content="Servicekontakte"      Height="30" Margin="5" ToolTip="Öffnet die Servicekontaktübersicht (Web)."/>
                                    <Button Name="btnMyUpdateLink" Content="DATEV myUpdate" Height="30" Margin="5" ToolTip="Öffnet DATEV myUpdate (Web)."/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Cloud" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <Button Name="btnMyDATEVPortal"             Content="MyDATEV Portal"              Height="30" Margin="5" ToolTip="Öffnet das MyDATEV Portal (Web)."/>
                                    <Button Name="btnDATEVUnternehmenOnline"    Content="DATEV Unternehmen Online"   Height="30" Margin="5" ToolTip="Öffnet DATEV Unternehmen Online (Web)."/>
                                    <Button Name="btnLogistikauftragOnline"      Content="Logistikauftrag Online"     Height="30" Margin="5" ToolTip="Öffnet Logistikauftrag Online (Web)."/>
                                    <Button Name="btnLizenzverwaltungOnline"     Content="Lizenzverwaltung Online"    Height="30" Margin="5" ToolTip="Öffnet die Lizenzverwaltung Online (Web)."/>
                                    <Button Name="btnDATEVRechteraumOnline"      Content="DATEV Rechteraum online"    Height="30" Margin="5" ToolTip="Öffnet den DATEV Rechteraum online (Web)."/>
                                    <Button Name="btnDATEVRechteverwaltungOnline" Content="DATEV Rechteverwaltung online" Height="30" Margin="5" ToolTip="Öffnet die DATEV Rechteverwaltung online (Web)."/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Verwaltung und Technik" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <Button Name="btnSmartLoginAdministration"   Content="SmartLogin Administration"  Height="30" Margin="5" ToolTip="Öffnet die SmartLogin Administration (Web)."/>
                                    <Button Name="btnMyDATEVBestandsmanagement"   Content="MyDATEV Bestandsmanagement" Height="30" Margin="5" ToolTip="Öffnet das MyDATEV Bestandsmanagement (Web)."/>
                                    <Button Name="btnWeitereCloudAnwendungen"    Content="Weitere Cloud Anwendungen"  Height="30" Margin="5" ToolTip="Zeigt weitere DATEV Cloud-Anwendungen (Web)."/>
                                </StackPanel>
                            </GroupBox>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
                <TabItem Header="Downloads">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Orientation="Vertical" Margin="10">
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Direkt Downloads von Liste" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <ComboBox Name="cmbDynamicDownloads" Width="300" Margin="5" ToolTip="Wählen Sie einen Eintrag für den Direkt-Download."/>
                                    <TextBlock Name="txtDownloadListMeta" FontSize="10" Foreground="Gray" Margin="5" HorizontalAlignment="Center" />
                                    <Button Name="btnStartDynamicDownload" Content="Download starten" Height="30" Width="150" Margin="5" ToolTip="Startet den Download des gewählten Eintrags."/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Downloads von datev.de" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <Button Name="btnDatevDownloadbereich" Content="DATEV Downloadbereich" Height="30" Margin="5" ToolTip="Öffnet den offiziellen DATEV Downloadbereich (Web)."/>
                                    <Button Name="btnDatevSmartDocs" Content="DATEV Smart Docs" Height="30" Margin="5" ToolTip="Öffnet die DATEV Smart Docs Downloadseite (Web)."/>
                                    <Button Name="btnDatentraegerDownloadPortal" Content="Datenträger Download Portal" Height="30" Margin="5" ToolTip="Öffnet das Datenträger Download Portal (Web)."/>
                                </StackPanel>
                            </GroupBox>
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
                <TabItem Header="Checklisten">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Orientation="Vertical" Margin="10">
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Checklistenverwaltung" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <!-- Buttons über dem Dropdown -->
                                    <StackPanel Orientation="Horizontal" Margin="0,5,0,5">
                                        <Button Name="btnChecklistNew" Content="Neu" Height="20" Width="120" Margin="5"/>
                                        <Button Name="btnChecklistDelete" Content="Löschen" Height="20" Width="120" Margin="5"/>
                                        <Button Name="btnChecklistRename" Content="Umbenennen" Height="20" Width="120"/>
                                    </StackPanel>
                                    <!-- Dropdown -->
                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                        <TextBlock Text="Checkliste:" VerticalAlignment="Center" Margin="0,0,5,0"/>
                                        <ComboBox Name="cmbChecklists" Width="220" Margin="0,0,10,0"/>
                                        <Button Name="btnAddChecklistItem" Content="+" Width="50" ToolTip="Neuen Checkpunkt hinzufügen"/>
                                    </StackPanel>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Name="gbChecklistContent" Margin="5,10,5,5">
                                <GroupBox.Header>
                                    <TextBlock Name="txtChecklistName" Text="Keine Checkliste ausgewählt" FontWeight="Bold" FontSize="12" Margin="10,10,5,5"/>
                                </GroupBox.Header>
                                <StackPanel Name="spChecklistDynamic" />
                            </GroupBox>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
                <TabItem Header="Einstellungen">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel Orientation="Vertical" Margin="10">
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Update Funktionen" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <Button Name="btnCheckUpdateSettings" Content="Script auf Update prüfen" Height="30" Margin="5" ToolTip="Prüft das Script auf Updates."/>
                                    <Button Name="btnUpdateDownloadList" Content="Download-Liste aktualisieren" Height="30" Margin="5" ToolTip="Lädt die aktuelle Download-Liste von GitHub."/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Einstellungen" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <!-- Hier könnten weitere Einstellungen ergänzt werden -->
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Margin="5,5,5,5">
                                <GroupBox.Header>
                                    <TextBlock Text="Systeminformationen" FontWeight="Bold" FontSize="12"/>
                                </GroupBox.Header>
                                <StackPanel>
                                    <TextBlock Name="txtSysInfoOS" FontSize="12" Margin="2"/>
                                    <TextBlock Name="txtSysInfoUser" FontSize="12" Margin="2"/>
                                    <TextBlock Name="txtSysInfoComputer" FontSize="12" Margin="2"/>
                                    <TextBlock Name="txtSysInfoPS" FontSize="12" Margin="2"/>
                                    <TextBlock Name="txtSysInfoDotNet" FontSize="12" Margin="2"/>
                                    <TextBlock Name="txtSysInfoDisk" FontSize="12" Margin="2"/>
                                    <TextBlock Name="txtSysInfoDATEVPP" FontSize="12" Margin="2"/>
                                    <Button Name="btnRefreshSysInfo" Content="Aktualisieren" Width="100" Margin="5,10,0,0" ToolTip="Systeminformationen neu einlesen."/>
                                </StackPanel>
                            </GroupBox>
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
Write-Log "UI geladen."
#endregion

#region Version und Titel
# Definiert die lokale Version des Skripts und setzt den Fenstertitel.
$localVersion = "1.0.13"
$window.Title = "DATEV Toolbox v$localVersion"
Write-Log "Script-Version: $localVersion"
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
        "btnOpenDownloadFolder", "btnNgenAll40", "btnLeistungsindex", "btnCheckUpdateSettings", "cmbDynamicDownloads", "btnStartDynamicDownload", "btnUpdateDownloadList", "txtDownloadListMeta",        # Systeminfo Controls
        "txtSysInfoOS", "txtSysInfoUser", "txtSysInfoComputer", "txtSysInfoPS", "txtSysInfoDotNet", "txtSysInfoDisk", "txtSysInfoDATEVPP", "btnRefreshSysInfo",
        # Update Termine Controls
        "spUpdateDates",
        "btnUpdateDates",
        # MyUpdate Link
        "btnMyUpdateLink",
        # Dynamische Checklistenverwaltung
        "cmbChecklists", "btnChecklistNew", "btnChecklistDelete", "btnChecklistRename", "spChecklistDynamic", "btnAddChecklistItem", "txtChecklistName"
    )
    foreach ($name in $controlNames) {
        $ctrl = $window.FindName($name)
        $global:Controls[$name] = $ctrl
        if (-not $ctrl) {
            Write-Log "Warnung: Control '$name' wurde im XAML nicht gefunden." -IsError
        }
        else {
            Write-Log "Control initialisiert: $name"
        }
    }
}
Initialize-Controls

# === Dynamische Checklistenverwaltung ===

# Hilfsfunktion: Objekt (rekursiv) in Hashtable umwandeln
function ConvertTo-Hashtable {
    param($obj)
    if ($obj -is [hashtable]) { return $obj }
    if ($obj -is [pscustomobject]) {
        $ht = @{}
        foreach ($prop in $obj.PSObject.Properties) {
            $ht[$prop.Name] = ConvertTo-Hashtable $prop.Value
        }
        return $ht
    }
    # Korrekte Typprüfung für String!
    if ($obj -is [System.Collections.IEnumerable] -and $obj -isnot [string]) {
        $arr = @()
        foreach ($item in $obj) { $arr += ConvertTo-Hashtable $item }
        return $arr
    }
    return $obj
}

$ChecklistDir = Join-Path $env:APPDATA 'DATEV-Toolbox/Checklisten'
if (-not (Test-Path $ChecklistDir)) {
    New-Item -Path $ChecklistDir -ItemType Directory | Out-Null
}

# Hinzufügen globaler Variablen
$global:Checklists = @()
$global:CurrentChecklist = $null

# Funktion zum Aktualisieren der Dropdown-Liste
function Update-ChecklistDropdown {
    $cmbChecklists = $global:Controls["cmbChecklists"]
    if (-not $cmbChecklists) { 
        Write-Log "Fehler: ComboBox für Checklisten nicht gefunden" -IsError
        return 
    }

    $cmbChecklists.Items.Clear()
    foreach ($checklist in $global:Checklists) {
        [void]$cmbChecklists.Items.Add($checklist.Name)
    }
    
    if ($cmbChecklists.Items.Count -gt 0) {
        $cmbChecklists.SelectedIndex = 0
    }
}

# Funktion zum Laden aller Checklisten
function Import-Checklists {
    $global:Checklists = @()
    
    Get-ChildItem -Path $ChecklistDir -Filter "*.json" | ForEach-Object {
        try {
            $content = Get-Content $_.FullName -Raw | ConvertFrom-Json
            $newChecklist = @{
                Name = $content.Name
                Items = @()
                FilePath = $_.FullName
            }
            
            if ($content.Items) {
                foreach ($item in $content.Items) {
                    $newItem = ConvertTo-Hashtable $item
                    if (-not $newItem.ContainsKey('Id') -or [string]::IsNullOrWhiteSpace($newItem.Id)) {
                        $newItem.Id = [guid]::NewGuid().ToString()
                    }
                    $newChecklist.Items += $newItem
                }
            }
            
            $global:Checklists += $newChecklist
        }
        catch {
            Write-Log "Fehler beim Laden der Checkliste $($_.Name): $($_.Exception.Message)" -IsError
        }
    }
    
    Update-ChecklistDropdown
    Write-Log "Checklisten geladen: $($global:Checklists.Count) gefunden"
}

function Save-CurrentChecklist {
    if (-not $global:CurrentChecklist) { return }
    
    try {
        $content = @{
            Name = $global:CurrentChecklist.Name
            Items = $global:CurrentChecklist.Items
        }
        $content | ConvertTo-Json -Depth 10 | Set-Content $global:CurrentChecklist.FilePath -Encoding UTF8
        Write-Log "Checkliste gespeichert: $($global:CurrentChecklist.Name)"
    }
    catch {
        Write-Log "Fehler beim Speichern der Checkliste: $($_.Exception.Message)" -IsError
    }
}

function Add-ChecklistItem {
    param($checklist, $text)
    if (-not $checklist) { return }
    
    $newItem = @{
        Id = [guid]::NewGuid().ToString()
        Text = $text
        IsChecked = $false
    }
    
    $checklist.Items += $newItem
    Save-CurrentChecklist
    Show-Checklist $checklist
}

# Event-Handler für Checkbox-Änderungen
function Register-CheckboxEvent {
    param($checkbox, $itemId)
    
    $checkbox.Add_Checked({
        if (-not $global:CurrentChecklist) { return }
        $item = $global:CurrentChecklist.Items | Where-Object { $_.Id -eq $itemId }
        if ($item) {
            $item.IsChecked = $true
            Save-CurrentChecklist
        }
    })
    
    $checkbox.Add_Unchecked({
        if (-not $global:CurrentChecklist) { return }
        $item = $global:CurrentChecklist.Items | Where-Object { $_.Id -eq $itemId }
        if ($item) {
            $item.IsChecked = $false
            Save-CurrentChecklist
        }
    })
}

# Aktualisierte Show-Checklist Funktion
function Show-Checklist {
    param($checklist)
    if (-not $checklist) { 
        $global:Controls["txtChecklistName"].Text = "Keine Checkliste ausgewählt"
        return 
    }
    $global:CurrentChecklist = $checklist
    $global:Controls["txtChecklistName"].Text = $checklist.Name
    
    $spChecklistDynamic = $global:Controls["spChecklistDynamic"]
    $spChecklistDynamic.Children.Clear()
    
    foreach ($item in $checklist.Items) {
        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Orientation = "Horizontal"
        $panel.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
        $panel.MaxWidth = 400

        $textBlock = New-Object System.Windows.Controls.TextBlock
        $textBlock.Text = $item.Text
        $textBlock.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $textBlock.MaxWidth = 360

        $checkbox = New-Object System.Windows.Controls.CheckBox
        $checkbox.Content = $textBlock
        $checkbox.IsChecked = [System.Convert]::ToBoolean($item.IsChecked)
        $checkbox.Tag = $item.Id
        $checkbox.MaxWidth = 380
        $checkbox.ToolTip = "Klicken Sie hier, um den Status zu ändern. Aktueller Status: $($item.IsChecked ? 'Erledigt' : 'Offen')"        # Erstelle Kontextmenü
        $contextMenu = New-Object System.Windows.Controls.ContextMenu
        
        # Bearbeiten-Option
        $editMenuItem = New-Object System.Windows.Controls.MenuItem
        $editMenuItem.Header = "Bearbeiten"
        $editMenuItem.Add_Click({
            param($menuSource, $menuEventArgs)
            $itemId = $menuSource.DataContext
            $item = $global:CurrentChecklist.Items | Where-Object { $_.Id -eq $itemId }
            if ($item) {
                Add-Type -AssemblyName 'Microsoft.VisualBasic'
                $newText = [Microsoft.VisualBasic.Interaction]::InputBox(
                    "Text bearbeiten:",
                    "Eintrag bearbeiten",
                    $item.Text
                )
                if (-not [string]::IsNullOrWhiteSpace($newText) -and $newText -ne $item.Text) {
                    $item.Text = $newText
                    Save-CurrentChecklist
                    Show-Checklist $global:CurrentChecklist
                    Write-Log "Checklisteneintrag wurde bearbeitet"
                }
            }
        })
        $editMenuItem.DataContext = $item.Id
        $contextMenu.Items.Add($editMenuItem)        # Nach oben-Option
        $moveUpMenuItem = New-Object System.Windows.Controls.MenuItem
        $moveUpMenuItem.Header = "Nach oben"
        $currentIndex = $global:CurrentChecklist.Items.IndexOf(($global:CurrentChecklist.Items | Where-Object { $_.Id -eq $item.Id }))
        $moveUpMenuItem.IsEnabled = $currentIndex -gt 0
        $moveUpMenuItem.Add_Click({
            param($menuSource, $menuEventArgs)
            $itemId = $menuSource.DataContext
            $currentIndex = $global:CurrentChecklist.Items.IndexOf(($global:CurrentChecklist.Items | Where-Object { $_.Id -eq $itemId }))
            if ($currentIndex -gt 0) {
                $temp = $global:CurrentChecklist.Items[$currentIndex]
                $global:CurrentChecklist.Items[$currentIndex] = $global:CurrentChecklist.Items[$currentIndex - 1]
                $global:CurrentChecklist.Items[$currentIndex - 1] = $temp
                Save-CurrentChecklist
                Show-Checklist $global:CurrentChecklist
                Write-Log "Checklisteneintrag wurde nach oben verschoben"
            }
        })
        $moveUpMenuItem.DataContext = $item.Id
        $contextMenu.Items.Add($moveUpMenuItem)

        # Nach unten-Option
        $moveDownMenuItem = New-Object System.Windows.Controls.MenuItem
        $moveDownMenuItem.Header = "Nach unten"
        $moveDownMenuItem.IsEnabled = $currentIndex -lt ($global:CurrentChecklist.Items.Count - 1)
        $moveDownMenuItem.Add_Click({
            param($menuSource, $menuEventArgs)
            $itemId = $menuSource.DataContext
            $currentIndex = $global:CurrentChecklist.Items.IndexOf(($global:CurrentChecklist.Items | Where-Object { $_.Id -eq $itemId }))
            if ($currentIndex -lt $global:CurrentChecklist.Items.Count - 1) {
                $temp = $global:CurrentChecklist.Items[$currentIndex]
                $global:CurrentChecklist.Items[$currentIndex] = $global:CurrentChecklist.Items[$currentIndex + 1]
                $global:CurrentChecklist.Items[$currentIndex + 1] = $temp
                Save-CurrentChecklist
                Show-Checklist $global:CurrentChecklist
                Write-Log "Checklisteneintrag wurde nach unten verschoben"
            }
        })
        $moveDownMenuItem.DataContext = $item.Id
        $contextMenu.Items.Add($moveDownMenuItem)

        # Löschen-Option
        $deleteMenuItem = New-Object System.Windows.Controls.MenuItem
        $deleteMenuItem.Header = "Löschen"
        $deleteMenuItem.Add_Click({
            param($menuSource, $menuEventArgs)
            $itemId = $menuSource.DataContext
            $result = [System.Windows.MessageBox]::Show(
                "Möchten Sie diesen Eintrag wirklich löschen?",
                "Eintrag löschen",
                'YesNo',
                'Warning'
            )
            if ($result -eq 'Yes') {                $global:CurrentChecklist.Items = @($global:CurrentChecklist.Items | Where-Object { $_.Id -ne $itemId })
                Save-CurrentChecklist
                Show-Checklist $global:CurrentChecklist
                Write-Log "Checklisteneintrag wurde gelöscht"
            }
        })
        $deleteMenuItem.DataContext = $item.Id
        $contextMenu.Items.Add($deleteMenuItem)
        $checkbox.ContextMenu = $contextMenu
        
        $checkbox.Add_Checked({
            param($s, $e)
            $itemId = $s.Tag
            $item = $global:CurrentChecklist.Items | Where-Object { $_.Id -eq $itemId }
            if ($item) {
                $item.IsChecked = $true
                Save-CurrentChecklist
                Write-Log "Checkbox '$($item.Text)' wurde als erledigt markiert"
            }
        })
        
        $checkbox.Add_Unchecked({
            param($s, $e)
            $itemId = $s.Tag
            $item = $global:CurrentChecklist.Items | Where-Object { $_.Id -eq $itemId }
            if ($item) {
                $item.IsChecked = $false
                Save-CurrentChecklist
                Write-Log "Checkbox '$($item.Text)' wurde als unerledigt markiert"
            }
        })
        
        $panel.Children.Add($checkbox)
        $spChecklistDynamic.Children.Add($panel)
    }
}

# Event-Handler für den "Neu" Button
Register-ButtonAction -Button $Controls["btnChecklistNew"] -Action {
    $name = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Name der neuen Checkliste:",
        "Neue Checkliste"
    )
    
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        $filePath = Join-Path $ChecklistDir "$([guid]::NewGuid()).json"
        $newChecklist = @{
            Name = $name
            Items = @()
            FilePath = $filePath
        }
        
        @{ Name = $name; Items = @() } | 
            ConvertTo-Json | 
            Set-Content $filePath -Encoding UTF8
            
        $global:Checklists += $newChecklist
        Update-ChecklistDropdown
        $Controls["cmbChecklists"].SelectedItem = $name
        Write-Log "Neue Checkliste erstellt: $name"
    }
}

# Event-Handler für den "Löschen" Button
Register-ButtonAction -Button $Controls["btnChecklistDelete"] -Action {
    $selectedName = $Controls["cmbChecklists"].SelectedItem
    if ($selectedName) {
        $result = [System.Windows.MessageBox]::Show(
            "Möchten Sie die Checkliste '$selectedName' wirklich löschen?",
            "Checkliste löschen",
            'YesNo',
            'Warning'
        )
        if ($result -eq 'Yes') {
            $checklist = $global:Checklists | Where-Object { $_.Name -eq $selectedName }
            if ($checklist -and (Test-Path $checklist.FilePath)) {
                Remove-Item $checklist.FilePath
            }
            $global:Checklists = @($global:Checklists | Where-Object { $_.Name -ne $selectedName })
            $global:CurrentChecklist = $null
            Update-ChecklistDropdown
            $Controls["spChecklistDynamic"].Children.Clear()
            Write-Log "Checkliste gelöscht: $selectedName"
        }
    }
}

# Event-Handler für den "Umbenennen" Button
Register-ButtonAction -Button $Controls["btnChecklistRename"] -Action {
    $selectedName = $Controls["cmbChecklists"].SelectedItem
    if ($selectedName) {
        $newName = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Neuer Name für die Checkliste:",
            "Checkliste umbenennen",
            $selectedName
        )
        if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne $selectedName) {
            $checklist = $global:Checklists | Where-Object { $_.Name -eq $selectedName }
            if ($checklist) {
                $checklist.Name = $newName
                Save-CurrentChecklist
                Update-ChecklistDropdown
                $Controls["cmbChecklists"].SelectedItem = $newName
                Write-Log "Checkliste umbenannt: $selectedName -> $newName"
            }
        }
    }
    else {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie eine Checkliste aus.", "Hinweis", "OK", "Information")
    }
}

# Event-Handler für den "+" Button
Register-ButtonAction -Button $Controls["btnAddChecklistItem"] -Action {
    if (-not $global:CurrentChecklist) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie zuerst eine Checkliste aus.", "Hinweis", "OK", "Information")
        return
    }
    
    $text = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Text für den neuen Checkpunkt:",
        "Checkpunkt hinzufügen"
    )
    
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        Add-ChecklistItem $global:CurrentChecklist $text
        Write-Log "Neuer Checkpunkt hinzugefügt: $text"
    }
}

# Event-Handler für das Dropdown
$Controls["cmbChecklists"].Add_SelectionChanged({
    param($senderObj, $e)
    $selectedName = $senderObj.SelectedItem
    if ($selectedName) {
        $checklist = $global:Checklists | Where-Object { $_.Name -eq $selectedName }
        Show-Checklist $checklist
    }
})

# Lade Checklisten beim Start
Import-Checklists
#endregion

# Systeminfo-Funktionen müssen vor dem ersten Aufruf definiert sein!
function Get-SystemInfo {
    $os = Get-CimInstance Win32_OperatingSystem
    $osVersion = "{0} ({1})" -f $os.Caption, $os.Version
    $user = $env:USERNAME
    $computer = $env:COMPUTERNAME
    $psVersion = $PSVersionTable.PSVersion.ToString()
    # .NET-Versionen
    $dotNet = (Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse -ErrorAction SilentlyContinue |
        Get-ItemProperty -Name Version -ErrorAction SilentlyContinue |
        Where-Object { $_.Version -match '^\d' } |
        Select-Object -ExpandProperty Version -Unique) -join ', '
    if (-not $dotNet) { $dotNet = "Nicht gefunden" }
    # Freier Speicher auf C:
    $sysDrive = (Get-PSDrive -Name C -ErrorAction SilentlyContinue)
    if ($sysDrive) {
        $freeGB = [math]::Round($sysDrive.Free / 1GB, 1)
        $disk = "Freier Speicher (C:): $freeGB GB"
    }
    else {
        $disk = "Freier Speicher (C:): Nicht ermittelbar"
    }
    $datevpp = $env:DATEVPP
    if (-not $datevpp) { $datevpp = "Nicht gesetzt" }
    return @{
        OS       = "Betriebssystem: $osVersion"
        User     = "Benutzer: $user"
        Computer = "Computername: $computer"
        PS       = "PowerShell-Version: $psVersion"
        DotNet   = ".NET-Version(en): $dotNet"
        Disk     = $disk
        DATEVPP  = "DATEVPP: $datevpp"
    }
}

function Show-SystemInfo {
    Write-Log "Lese Systeminformationen ..."
    $info = Get-SystemInfo
    $Controls["txtSysInfoOS"].Text = $info.OS
    $Controls["txtSysInfoUser"].Text = $info.User
    $Controls["txtSysInfoComputer"].Text = $info.Computer
    $Controls["txtSysInfoPS"].Text = $info.PS
    $Controls["txtSysInfoDotNet"].Text = $info.DotNet
    $Controls["txtSysInfoDisk"].Text = $info.Disk
    $Controls["txtSysInfoDATEVPP"].Text = $info.DATEVPP
    Write-Log "Systeminformationen aktualisiert."
}

Show-SystemInfo
#endregion

#region Hilfsfunktionen
# Diverse Hilfsfunktionen, z.B. für Versionsvergleiche.
function Compare-Version {
    param([string]$v1, [string]$v2)
    try {
        $ver1 = [Version]$v1
        $ver2 = [Version]$v2
        return $ver1.CompareTo($ver2)
    }
    catch {
        return [string]::Compare($v1, $v2)
    }
}

function Update-DownloadListMetaDisplay {
    param($downloadsMeta)
    if ($downloadsMeta.Version -and $downloadsMeta.Datum) {
        $Controls["txtDownloadListMeta"].Text = "Download-Liste Version: $($downloadsMeta.Version), Datum: $($downloadsMeta.Datum)"
    }
    else {
        $Controls["txtDownloadListMeta"].Text = "Download-Liste: keine Metadaten gefunden."
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
        Write-Log "Download-Ordner neu angelegt: $targetDir"
    }
    return $targetDir
}

function Get-DatevFile {
    param([string]$Url, [string]$FileName)
    $targetDir = Get-DownloadFolder
    $targetFile = Join-Path $targetDir $FileName
    Write-Log "Starte Download: $Url -> $targetFile"
    if (Test-Path $targetFile) {
        $result = [System.Windows.MessageBox]::Show("Die Datei '$FileName' existiert bereits. Überschreiben?", "Datei existiert", 'YesNo', 'Warning')
        if ($result -ne 'Yes') {
            Write-Log "Download abgebrochen: $FileName existiert bereits."
            return
        }
    }
    $response = $null
    $stream = $null
    $fileStream = $null
    try {
        Write-Log "Lade $FileName herunter ..."
        $window.Dispatcher.Invoke([action] {}, 'Background')
        $webRequest = [System.Net.WebRequest]::Create($Url)
        $webRequest.Timeout = 60000  # 60 Sekunden
        $response = $webRequest.GetResponse()
        $stream = $response.GetResponseStream()
        $fileStream = [System.IO.File]::Create($targetFile)
        $stream.CopyTo($fileStream)
        Write-Log "Download abgeschlossen: $targetFile"
        [System.Windows.MessageBox]::Show("Download abgeschlossen: $targetFile", "Download", 'OK', 'Information')
    }
    catch {
        Write-Log "Fehler beim Download: $($_.Exception.Message)" -IsError
        [System.Windows.MessageBox]::Show("Fehler beim Download: $($_.Exception.Message)", "Fehler", 'OK', 'Error')
    }
    finally {
        if ($fileStream) { $fileStream.Close() }
        if ($stream) { $stream.Close() }
        if ($response) { $response.Close() }
    }
}
#endregion

#region Dynamische Download-Liste laden und Dropdown befüllen
Write-Log "Lade dynamische Download-Liste ..."
# Liest die Datei 'downloads.json' ein und befüllt das Dropdown im Download-Tab
$dynamicDownloadsFile = Join-Path (Join-Path $env:APPDATA 'DATEV-Toolbox') 'downloads.json'
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
        Update-DownloadListMetaDisplay $downloadsMeta
        Write-Log "Download-Liste geladen."
    }
    catch {
        Write-Log "Fehler beim Laden der Download-Liste: $($_.Exception.Message)" -IsError
        $Controls["txtDownloadListMeta"].Text = "Fehler beim Laden der Download-Liste."
    }
}
else {
    Write-Log "downloads.json nicht gefunden. Das Dropdown bleibt leer."
    $Controls["txtDownloadListMeta"].Text = "downloads.json nicht gefunden."
}

# Event-Handler für den Button 'Download starten' im dynamischen Bereich
Register-ButtonAction -Button $Controls["btnStartDynamicDownload"] -Action {
    Write-Log "Benutzeraktion: Download starten geklickt."
    $selIndex = $Controls["cmbDynamicDownloads"].SelectedIndex
    if ($selIndex -lt 0 -or $selIndex -ge $global:DynamicDownloads.Count) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie einen Download aus der Liste.", "Hinweis", 'OK', 'Information')
        Write-Log "Kein Download-Eintrag ausgewählt."
        return
    }
    $entry = $global:DynamicDownloads[$selIndex]
    if ($entry.Url -and $entry.Name) {
        Write-Log "Starte Download für: $($entry.Name) ($($entry.Url))"
        Get-DatevFile -Url $entry.Url -FileName ([System.IO.Path]::GetFileName($entry.Url))
    }
    else {
        [System.Windows.MessageBox]::Show("Ungültiger Eintrag in der Download-Liste.", "Fehler", 'OK', 'Error')
        Write-Log "Ungültiger Eintrag in der Download-Liste."
    }
}

# Event-Handler für den Button 'Download-Liste aktualisieren'
Register-ButtonAction -Button $Controls["btnUpdateDownloadList"] -Action {
    $onlineUrl = "https://raw.githubusercontent.com/Zdministrator/DATEV-Toolbox/refs/heads/main/downloads.json"
    Write-Log "Benutzeraktion: Download-Liste aktualisieren geklickt."
    try {
        Write-Log "Lade aktuelle Download-Liste von $onlineUrl ..."
        $window.Dispatcher.Invoke([action] {}, 'Background')
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
        Update-DownloadListMetaDisplay $downloadsMeta
        Write-Log "Download-Liste aktualisiert."
    }
    catch {
        Write-Log "Fehler beim Herunterladen der Download-Liste: $($_.Exception.Message)" -IsError
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
    Write-Log "Starte Update-Check ..."
    $testConnection = $false
    try {
        $request = [System.Net.WebRequest]::Create($versionUrl)
        $request.Method = "HEAD"
        $request.Timeout = 3000
        $response = $request.GetResponse()
        if ($response.StatusCode -eq 200 -or $response.StatusCode -eq 0) { $testConnection = $true }
        $response.Close()
    }
    catch { $testConnection = $false }
    if (-not $testConnection) {
        Write-Log "Keine Internetverbindung. Update-Check abgebrochen."
        [System.Windows.MessageBox]::Show("Es konnte keine Internetverbindung festgestellt werden. Der Update-Check wird abgebrochen.", "Update-Fehler", 'OK', 'Error')
        return
    }
    try {
        Write-Log "Prüfe auf Updates ..."
        $window.Dispatcher.Invoke([action] {}, 'Background')
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
                Write-Log "Update wird durchgeführt ..."
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
                $testFile = Join-Path $scriptDir "_write_test.tmp"
                try {
                    Set-Content -Path $testFile -Value "test" -ErrorAction Stop
                    Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Log "Keine Schreibrechte im Scriptverzeichnis: $scriptDir" -IsError
                    [System.Windows.MessageBox]::Show("Das Update kann nicht durchgeführt werden, da keine Schreibrechte im Scriptverzeichnis ($scriptDir) bestehen.", "Update-Fehler", 'OK', 'Error')
                    return
                }
                $newScriptPath = Join-Path $scriptDir "DATEV-Toolbox.ps1.new"
                $updateScriptPath = Join-Path $scriptDir "Update-DATEV-Toolbox.ps1"
                Write-Log "Lade neue Version herunter..."
                $window.Dispatcher.Invoke([action] {}, 'Background')
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
                }
                catch {
                    Write-Log "Fehler beim Erstellen des Update-Scripts: $_" -IsError
                }
                [System.Windows.MessageBox]::Show("Das Programm wird für das Update beendet und automatisch neu gestartet.", "Update wird durchgeführt", 'OK', 'Information')
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
        Write-Log "Fehler beim Update-Check: $($_.Exception.Message)" -IsError
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
    }
    else {
        $btn.ToolTip = "$toolName starten"
    }
    $btn.Add_Click({
            Write-Log "Starte $toolName ($exe) ..."
            try {
                Start-Process -FilePath $exe -ErrorAction Stop
                Write-LogDirect "$toolName gestartet: $exe"
            }
            catch {
                Write-LogDirect "Fehler beim Start von ${toolName}: $($_.Exception.Message)"
                Write-Log "Fehler beim Start von ${toolName}: $($_.Exception.Message)" -IsError
            }
        }.GetNewClosure())
}

function Register-WebLinkHandler {
    param ([System.Windows.Controls.Button]$Button, [string]$Name, [string]$Url)
    $Button.Add_Click({
            $btnContent = $Button.Content
            Write-LogDirect "Öffne $btnContent..."
            Write-Log "Öffne Weblink: $btnContent ($Url)"
            $reachable = $false
            try {
                $request = [System.Net.WebRequest]::Create($Url)
                $request.Method = "HEAD"
                $request.Timeout = 5000
                $response = $request.GetResponse()
                $response.Close()
                $reachable = $true
            }
            catch { $reachable = $false }
            if (-not $reachable) {
                Write-LogDirect "Keine Verbindung zu $Url möglich – $btnContent wird nicht geöffnet."
                Write-Log "Weblink nicht erreichbar: $Url" -IsError
                return
            }
            try {
                Start-Process "explorer.exe" $Url
                Write-LogDirect "$btnContent geöffnet."
                Write-Log "Weblink geöffnet: $Url"
            }
            catch {
                Write-LogDirect ("Fehler beim Öffnen von {0}: {1}" -f $btnContent, $_)
                Write-Log "Fehler beim Öffnen von ${Url}: $($_.Exception.Message)" -IsError
            }
        }.GetNewClosure())
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

# Register the DATEV myUpdate link as a web link button
if ($Controls["btnMyUpdateLink"]) {
    Register-WebLinkHandler -Button $Controls["btnMyUpdateLink"] -Name "btnMyUpdateLink" -Url "https://apps.datev.de/myupdates/home"
}
# Event-Handler für den Button 'Script auf Update prüfen' im Tab Einstellungen
if ($Controls["btnCheckUpdateSettings"]) {
    Register-ButtonAction -Button $Controls["btnCheckUpdateSettings"] -Action {
        Write-Log "Benutzeraktion: Script auf Update prüfen geklickt."
        Test-ForUpdate
    }
}
# Event-Handler für den Button 'Download Ordner öffnen'
if ($Controls["btnOpenDownloadFolder"]) {
    Register-ButtonAction -Button $Controls["btnOpenDownloadFolder"] -Action {
        $folder = Get-DownloadFolder
        Write-Log "Öffne Download-Ordner: $folder"
        try {
            Start-Process explorer.exe $folder
        }
        catch {
            Write-Log "Fehler beim Öffnen des Download-Ordners: $($_.Exception.Message)" -IsError
            [System.Windows.MessageBox]::Show("Fehler beim Öffnen des Download-Ordners: $($_.Exception.Message)", "Fehler", 'OK', 'Error')
        }
    }
}
#endregion

#region System- und Umgebungsinformationen
function Get-SystemInfo {
    $os = Get-CimInstance Win32_OperatingSystem
    $osVersion = "{0} ({1})" -f $os.Caption, $os.Version
    $user = $env:USERNAME
    $computer = $env:COMPUTERNAME
    $psVersion = $PSVersionTable.PSVersion.ToString()
    # .NET-Versionen
    $dotNet = (Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse -ErrorAction SilentlyContinue |
        Get-ItemProperty -Name Version -ErrorAction SilentlyContinue |
        Where-Object { $_.Version -match '^\d' } |
        Select-Object -ExpandProperty Version -Unique) -join ', '
    if (-not $dotNet) { $dotNet = "Nicht gefunden" }
    # Freier Speicher auf C:
    $sysDrive = (Get-PSDrive -Name C -ErrorAction SilentlyContinue)
    if ($sysDrive) {
        $freeGB = [math]::Round($sysDrive.Free / 1GB, 1)
        $disk = "Freier Speicher (C:): $freeGB GB"
    }
    else {
        $disk = "Freier Speicher (C:): Nicht ermittelbar"
    }
    $datevpp = $env:DATEVPP
    if (-not $datevpp) { $datevpp = "Nicht gesetzt" }
    return @{
        OS       = "Betriebssystem: $osVersion"
        User     = "Benutzer: $user"
        Computer = "Computername: $computer"
        PS       = "PowerShell-Version: $psVersion"
        DotNet   = ".NET-Version(en): $dotNet"
        Disk     = $disk
        DATEVPP  = "DATEVPP: $datevpp"
    }
}

function Show-SystemInfo {
    $info = Get-SystemInfo
    $Controls["txtSysInfoOS"].Text = $info.OS
    $Controls["txtSysInfoUser"].Text = $info.User
    $Controls["txtSysInfoComputer"].Text = $info.Computer
    $Controls["txtSysInfoPS"].Text = $info.PS
    $Controls["txtSysInfoDotNet"].Text = $info.DotNet
    $Controls["txtSysInfoDisk"].Text = $info.Disk
    $Controls["txtSysInfoDATEVPP"].Text = $info.DATEVPP
}
#endregion

# --- Update Termine aus ICS anzeigen ---
function Show-NextUpdateDates {
    Write-Log "Lese Update-Termine aus ICS-Datei ..."
    $icsFile = Join-Path (Split-Path $dynamicDownloadsFile) 'Jahresplanung_2025.ics'
    $sp = $Controls["spUpdateDates"]
    $sp.Children.Clear()
    if (-not (Test-Path $icsFile)) {
        Write-Log "Keine lokale ICS-Datei gefunden: $icsFile" -IsError
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = "Keine lokale ICS-Datei gefunden. Bitte erst aktualisieren."
        $sp.Children.Add($tb) | Out-Null
        return
    }
    try {
        $icsContent = Get-Content $icsFile -Raw
        $lines = $icsContent -split "`n"
        $events = @()
        $currentEvent = @{}
        $lastKey = $null
        foreach ($lineRaw in $lines) {
            $line = $lineRaw.TrimEnd("`r", "`n")
            if ($line -eq "BEGIN:VEVENT") { $currentEvent = @{}; $lastKey = $null }
            elseif ($line -eq "END:VEVENT") {
                if ($currentEvent.DTSTART -and $currentEvent.SUMMARY) {
                    $events += [PSCustomObject]@{
                        DTSTART     = $currentEvent.DTSTART
                        SUMMARY     = $currentEvent.SUMMARY
                        DESCRIPTION = $currentEvent.DESCRIPTION
                    }
                }
                $currentEvent = @{}; $lastKey = $null
            }
            elseif ($line -match "^([A-Z]+).*:(.*)$") {
                $key = $matches[1]
                $val = $matches[2]
                $lastKey = $key
                if ($key -eq "DTSTART") { $currentEvent.DTSTART = $val }
                elseif ($key -eq "SUMMARY") { $currentEvent.SUMMARY = $val }
                elseif ($key -eq "DESCRIPTION") { $currentEvent.DESCRIPTION = $val }
            }
            elseif ($line -match "^[ \t](.*)$" -and $lastKey) {
                # Fortgesetzte Zeile (folded line)
                $continued = $matches[1]
                if ($lastKey -eq "DESCRIPTION") {
                    $currentEvent.DESCRIPTION += [Environment]::NewLine + $continued
                }
                elseif ($lastKey -eq "SUMMARY") {
                    $currentEvent.SUMMARY += $continued
                }
            }
        }
        Write-Log ("ICS: {0} VEVENTs gefunden." -f $events.Count)
        $now = Get-Date
        $upcoming = $events | Where-Object {
            $date = $_.DTSTART
            if ($date.Length -eq 8) { $date = [datetime]::ParseExact($date, 'yyyyMMdd', $null) }
            elseif ($date.Length -ge 15) { $date = [datetime]::ParseExact($date.Substring(0, 8), 'yyyyMMdd', $null) }
            else { $date = $null }
            $date -and $date -ge $now.Date
        } | Sort-Object {
            if ($_.DTSTART.Length -eq 8) { [datetime]::ParseExact($_.DTSTART, 'yyyyMMdd', $null) }
            elseif ($_.DTSTART.Length -ge 15) { [datetime]::ParseExact($_.DTSTART.Substring(0, 8), 'yyyyMMdd', $null) }
            else { $null }
        } | Select-Object -First 3
        Write-Log ("{0} anstehende Termine werden angezeigt." -f $upcoming.Count)
        if ($upcoming.Count -eq 0) {
            Write-Log "Keine anstehenden Termine gefunden."
            $tb = New-Object System.Windows.Controls.TextBlock
            $tb.Text = "Keine anstehenden Termine gefunden."
            $sp.Children.Add($tb) | Out-Null
        }
        else {
            foreach ($ev in $upcoming) {
                $date = $ev.DTSTART
                if ($date.Length -eq 8) { $date = [datetime]::ParseExact($date, 'yyyyMMdd', $null) }
                elseif ($date.Length -ge 15) { $date = [datetime]::ParseExact($date.Substring(0, 8), 'yyyyMMdd', $null) }
                else { $date = $null }
                if ($date) {
                    $tb = New-Object System.Windows.Controls.TextBlock
                    $tb.Text = "{0:dd.MM.yyyy} - {1}" -f $date, $ev.SUMMARY
                    if ($ev.DESCRIPTION) { $tb.ToolTip = $ev.DESCRIPTION }
                    $tb.FontSize = 12
                    $tb.Margin = '2'
                    $sp.Children.Add($tb) | Out-Null
                    Write-Log ("Termin: {0:dd.MM.yyyy} - {1}" -f $date, $ev.SUMMARY)
                }
            }
        }
    }
    catch {
        Write-Log "Fehler beim Laden oder Parsen der ICS-Datei: $($_.Exception.Message)" -IsError
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = "Fehler beim Laden der Termine."
        $sp.Children.Add($tb) | Out-Null
    }
}

Register-ButtonAction -Button $Controls["btnUpdateDates"] -Action {
    $icsUrl = "https://apps.datev.de/myupdates/assets/files/Jahresplanung_2025.ics"
    $icsFile = Join-Path (Split-Path $dynamicDownloadsFile) 'Jahresplanung_2025.ics'
    $sp = $Controls["spUpdateDates"]
    $sp.Children.Clear()
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = "Lade ICS-Datei ..."
    $sp.Children.Add($tb) | Out-Null
    Write-Log "Benutzeraktion: Update-Termine aktualisieren geklickt. Lade ICS von $icsUrl ..."
    try {
        Invoke-WebRequest -Uri $icsUrl -OutFile $icsFile -UseBasicParsing -TimeoutSec 15
        $sp.Children.Clear()
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = "ICS-Datei geladen. Lese Termine ..."
        $sp.Children.Add($tb) | Out-Null
        Write-Log "ICS-Datei erfolgreich geladen: $icsFile"
        Show-NextUpdateDates
    }
    catch {
        $sp.Children.Clear()
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = "Fehler beim Laden der ICS-Datei."
        $sp.Children.Add($tb) | Out-Null
        Write-Log "Fehler beim Laden der ICS-Datei: $($_.Exception.Message)" -IsError
    }
}
# Startet die automatische Update-Prüfung und zeigt das Hauptfenster an.
Write-Log "Scriptstart abgeschlossen. UI wird angezeigt."
Test-ForUpdate

# Zeige das Fenster nur, wenn es noch nicht geschlossen wurde
if ($window.Visibility -ne 'Closed') {
    $window.ShowDialog() | Out-Null
}

Write-Log "UI geschlossen. Speichere Einstellungen ..."
Save-CurrentChecklist
Save-Settings
Write-Log "Script beendet."
