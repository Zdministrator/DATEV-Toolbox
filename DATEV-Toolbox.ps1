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
                </StackPanel>
            </TabItem>
            <TabItem Header="Cloud Anwendungen">
                <StackPanel Orientation="Vertical" Margin="10">
                    <!-- Hier können Cloud-Buttons eingefügt werden -->
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
        <TextBox Name="txtLog" Grid.Row="1" IsReadOnly="True" VerticalScrollBarVisibility="Auto" Margin="0,5,0,0" />
    </Grid>
</Window>
"@

# Versionsnummer des lokalen Scripts
$localVersion = "1.0.0"

# URL zur Online-Versionsdatei (auf RAW-Content umgestellt)
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
                $tempFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "DATEV-Toolbox.ps1.new"
                Write-Log "Lade neue Version herunter..."
                Invoke-WebRequest -Uri $scriptUrl -OutFile $tempFile -UseBasicParsing
                Write-Log "Neue Version wurde als $tempFile gespeichert. Bitte das laufende Skript schließen und die Datei manuell ersetzen."
                [System.Windows.MessageBox]::Show("Neue Version wurde als $tempFile gespeichert.\nBitte das laufende Skript schließen und die Datei manuell ersetzen.", "Update heruntergeladen", 'OK', 'Information')
            } else {
                Write-Log "Update abgebrochen durch Benutzer."
            }
        } else {
            Write-Log "Keine neue Version verfügbar."
        }
    } catch {
        Write-Log "Fehler beim Update-Check: $_"
    }
}

# XAML laden
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Fenstertitel um Versionsnummer ergänzen
$window.Title = "DATEV Toolbox v$localVersion"

# Controls referenzieren
$txtLog = $window.FindName("txtLog")

# Nach dem Laden des Fensters auf Updates prüfen
Test-ForUpdate

# Fenster anzeigen
$window.ShowDialog() | Out-Null
