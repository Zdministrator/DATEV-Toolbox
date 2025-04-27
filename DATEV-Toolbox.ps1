Add-Type -AssemblyName PresentationFramework

# XAML-Definition für das Fenster
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="DATEV Toolbox" MinHeight="400" Width="400" ResizeMode="CanMinimize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="100" /> <!-- Log-Ausgabe -->
        </Grid.RowDefinitions>
        <TabControl Grid.Row="0" Margin="0,0,0,0" VerticalAlignment="Stretch">
            <TabItem Header="DATEV Tools">
                <StackPanel Orientation="Vertical" Margin="10">
                    <Label Content="Programme" FontWeight="Bold" Margin="5"/>
                    <Button Name="btnArbeitsplatz" Content="DATEV-Arbeitsplatz" Height="30" Margin="5"/>
                    <Button Name="btnInstallationsmanager" Content="Installationsmanager" Height="30" Margin="5"/>
                    <Button Name="btnServicetool" Content="Servicetool" Height="30" Margin="5"/>
                    <Label Content="Tools" FontWeight="Bold" Margin="5"/>
                    <Button Name="btnKonfigDBTools" Content="KonfigDB-Tools" Height="30" Margin="5"/>
                    <Button Name="btnEODBconfig" Content="EODBconfig" Height="30" Margin="5"/>
                    <Button Name="btnEOAufgabenplanung" Content="EO Aufgabenplanung" Height="30" Margin="5"/>
                </StackPanel>
            </TabItem>
            <TabItem Header="Cloud Anwendungen">
                <StackPanel Orientation="Vertical" Margin="10">
                    <Button Name="btnDATEVHilfeCenter"          Content="DATEV Hilfe Center"          Height="30" Margin="5"/>
                    <Button Name="btnMyDATEVPortal"             Content="MyDATEV Portal"              Height="30" Margin="5"/>
                    <Button Name="btnDATEVUnternehmenOnline"    Content="DATEV Unternehmen Online"   Height="30" Margin="5"/>
                    <Button Name="btnLogistikauftragOnline"      Content="Logistikauftrag Online"     Height="30" Margin="5"/>
                    <Button Name="btnLizenzverwaltungOnline"     Content="Lizenzverwaltung Online"    Height="30" Margin="5"/>
                    <Button Name="btnDATEVRechteraumOnline"      Content="DATEV Rechteraum online"    Height="30" Margin="5"/>
                    <Button Name="btnDATEVRechteverwaltungOnline" Content="DATEV Rechteverwaltung online" Height="30" Margin="5"/>
                    <Button Name="btnSmartLoginAdministration"   Content="SmartLogin Administration"  Height="30" Margin="5"/>
                    <Button Name="btnMyDATEVBestandsmanagement"   Content="MyDATEV Bestandsmanagement" Height="30" Margin="5"/>
                    <Button Name="btnWeitereCloudAnwendungen"    Content="Weitere Cloud Anwendungen"  Height="30" Margin="5"/>
                </StackPanel>
            </TabItem>
            <TabItem Header="Downloads">
                <StackPanel Orientation="Vertical" Margin="10">
                    <!-- Hier können Download-Buttons eingefügt werden -->
                </StackPanel>
            </TabItem>
            <TabItem Header="Performance">
                <StackPanel Orientation="Vertical" Margin="10">
                    <!-- Hier können Performance-Buttons eingefügt werden -->
                </StackPanel>
            </TabItem>
        </TabControl>
        <TextBox Name="txtLog" Grid.Row="1" IsReadOnly="True" VerticalScrollBarVisibility="Auto" Margin="0,5,0,0" TextWrapping="Wrap" />
    </Grid>
</Window>
"@

# Versionsnummer des lokalen Scripts
$localVersion = "1.0.1"

# URL zur Online-Versionsdatei
$versionUrl = "https://raw.githubusercontent.com/Zdministrator/DATEV-Toolbox/main/version.txt"
$scriptUrl = "https://raw.githubusercontent.com/Zdministrator/DATEV-Toolbox/main/DATEV-Toolbox.ps1"

function Write-Log($message) {
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $txtLog.AppendText("[$timestamp] $message`n")
}

function Test-ForUpdate {
    try {
        Write-Log "Prüfe auf Updates..."
        $remoteVersion = Invoke-RestMethod -Uri $versionUrl -UseBasicParsing
        if ($remoteVersion -and ($remoteVersion -ne $localVersion)) {
            Write-Log "Neue Version gefunden: $remoteVersion (aktuell: $localVersion)"
            $result = [System.Windows.MessageBox]::Show("Neue Version ($remoteVersion) verfügbar. Jetzt herunterladen?", "Update verfügbar", 'YesNo', 'Information')
            if ($result -eq 'Yes') {
                # Sicherstellen, dass der Script-Pfad immer verfügbar ist
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
                $newScriptPath = Join-Path -Path $scriptDir -ChildPath "DATEV-Toolbox.ps1.new"
                $updateScriptPath = Join-Path -Path $scriptDir -ChildPath "Update-DATEV-Toolbox.ps1"
                Write-Log "Lade neue Version herunter..."
                Invoke-WebRequest -Uri $scriptUrl -OutFile $newScriptPath -UseBasicParsing
                Write-Log "Neue Version wurde als $newScriptPath gespeichert. Update-Vorgang wird vorbereitet."

                # Update-Script erzeugen
                $updateScript = @"
Start-Sleep -Seconds 2
# Warten, bis das Hauptscript nicht mehr läuft
while (Get-Process -Id $PID -ErrorAction SilentlyContinue) { Start-Sleep -Milliseconds 500 }
# Ersetzen
Move-Item -Path '$newScriptPath' -Destination '$scriptPath' -Force
# Neues Script starten
Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File ""$scriptPath""'
# Update-Script löschen
Remove-Item -Path '$updateScriptPath' -Force
"@
                $updateScript = $updateScript -replace '\$PID', $PID
                Set-Content -Path $updateScriptPath -Value $updateScript -Encoding UTF8
                Write-Log "Update-Script wurde erstellt: $updateScriptPath"
                [System.Windows.MessageBox]::Show("Das Programm wird für das Update beendet und automatisch neu gestartet.", "Update wird durchgeführt", 'OK', 'Information')
                # Update-Script starten und beenden
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
        Write-Log "Fehler beim Update-Check: $_"
    }
}

# XAML laden
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Fenstertitel um Versionsnummer ergänzen
$window.Title = "DATEV Toolbox v$localVersion"

# Controls referenzieren
$controls = @("txtLog", "btnInstallationsmanager", "btnServicetool", "btnKonfigDBTools", "btnEOAufgabenplanung", "btnEODBconfig")
foreach ($controlName in $controls) {
    Set-Variable -Name $controlName -Value $window.FindName($controlName) -Scope Local
}

# Hilfsfunktion für Programm-Start
function Register-ToolButton {
    param (
        [string]$ButtonVar,
        [string]$ExePath,
        [string]$ToolName
    )
    $btn = Get-Variable -Name $ButtonVar -ValueOnly
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
                Start-Process -FilePath $exe
                Write-Log "$toolName gestartet: $exe"
            }
            catch {
                Write-Log "Fehler beim Start von ${toolName}: $_"
            }
        })
}

# Zuordnungstabelle für Buttons und Programme
$toolButtons = @(
    @{ Button = "btnInstallationsmanager"; Exe = "$env:DATEVPP\PROGRAMM\INSTALL\DvInesInstMan.exe"; Name = "Installationsmanager" },
    @{ Button = "btnServicetool"; Exe = "$env:DATEVPP\PROGRAMM\SRVTOOL\Srvtool.exe"; Name = "Servicetool" },
    @{ Button = "btnKonfigDBTools"; Exe = "$env:DATEVPP\PROGRAMM\B0001502\cdbtool.exe"; Name = "KonfigDB-Tools" },
    @{ Button = "btnEOAufgabenplanung"; Exe = "$env:DATEVPP\PROGRAMM\I0000085\EOControl.exe"; Name = "EO Aufgabenplanung" },
    @{ Button = "btnEODBconfig"; Exe = "$env:DATEVPP\PROGRAMM\EODB\EODBConfig.exe"; Name = "EODBconfig" }
)

# Event-Handler für alle Tool-Buttons registrieren
foreach ($entry in $toolButtons) {
    Register-ToolButton -ButtonVar $entry.Button -ExePath $entry.Exe -ToolName $entry.Name
}

# Überarbeitung der Weblink-Buttons im "Cloud Anwendungen" Tab
function Test-WebConnection {
    param (
        [Parameter(Mandatory)][string]$Url
    )
    try {
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Method = "HEAD"
        $request.Timeout = 5000
        $response = $request.GetResponse()
        $response.Close()
        return $true
    }
    catch {
        return $false
    }
}

function Start-WebLink {
    param (
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Url
    )

    Write-Log "Öffne $Name..."

    if (-not (Test-WebConnection -Url $Url)) {
        Write-Log "Keine Verbindung zu $Url möglich – $Name wird nicht geöffnet."
        return
    }

    try {
        Start-Process $Url
        Write-Log "$Name geöffnet."
    }
    catch {
        Write-Log ("Fehler beim Öffnen von {0}: {1}" -f $Name, $_)
    }
}

function Register-WebLinkHandler {
    param (
        [System.Windows.Controls.Button]$Button,
        [string]$Name,
        [string]$Url
    )

    $Button.Add_Click({
            Start-WebLink -Name $Name -Url $Url
        }.GetNewClosure())
}

# WebLinks Definition
$cloudButtons = @(
    @{ Name = "btnDATEVHilfeCenter"; Url = "https://apps.datev.de/help-center/" },
    @{ Name = "btnMyDATEVPortal"; Url = "https://apps.datev.de/mydatev" },
    @{ Name = "btnDATEVUnternehmenOnline"; Url = "https://duo.datev.de/" },
    @{ Name = "btnLogistikauftragOnline"; Url = "https://apps.datev.de/lao" },
    @{ Name = "btnLizenzverwaltungOnline"; Url = "https://apps.datev.de/lizenzverwaltung" },
    @{ Name = "btnDATEVRechteraumOnline"; Url = "https://apps.datev.de/rechteraum" },
    @{ Name = "btnDATEVRechteverwaltungOnline"; Url = "https://apps.datev.de/rvo-administration" },
    @{ Name = "btnSmartLoginAdministration"; Url = "https://go.datev.de/smartlogin-administration" },
    @{ Name = "btnMyDATEVBestandsmanagement"; Url = "https://apps.datev.de/mydata/" },
    @{ Name = "btnWeitereCloudAnwendungen"; Url = "https://www.datev.de/web/de/mydatev/datev-cloud-anwendungen/" }
)

foreach ($entry in $cloudButtons) {
    $btn = $window.FindName($entry.Name)
    if ($btn) {
        Register-WebLinkHandler -Button $btn -Name $entry.Name -Url $entry.Url
    }
}

# Nach dem Laden des Fensters auf Updates prüfen
Test-ForUpdate

# Fenster anzeigen
$window.ShowDialog() | Out-Null
