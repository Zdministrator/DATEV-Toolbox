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

# URL zur Online-Versionsdatei
$versionUrl = "https://github.com/Zdministrator/DATEV-Toolbox/blob/main/version.txt"
$scriptUrl = "https://github.com/Zdministrator/DATEV-Toolbox/blob/main/DATEV-Toolbox.ps1"

function Check-ForUpdate {
    try {
        $window.txtLog.AppendText("Prüfe auf Updates...`n")
        $remoteVersion = Invoke-RestMethod -Uri $versionUrl -UseBasicParsing
        if ($remoteVersion -and ($remoteVersion -ne $localVersion)) {
            $window.txtLog.AppendText("Neue Version gefunden: $remoteVersion (aktuell: $localVersion)`n")
            $result = [System.Windows.MessageBox]::Show("Neue Version ($remoteVersion) verfügbar. Jetzt herunterladen und Script ersetzen?", "Update verfügbar", 'YesNo', 'Information')
            if ($result -eq 'Yes') {
                $window.txtLog.AppendText("Lade neue Version herunter...`n")
                Invoke-WebRequest -Uri $scriptUrl -OutFile $MyInvocation.MyCommand.Definition -UseBasicParsing
                $window.txtLog.AppendText("Update abgeschlossen. Bitte Script neu starten.`n")
                [System.Windows.MessageBox]::Show("Update abgeschlossen. Bitte Script neu starten.", "Update", 'OK', 'Information')
                exit
            } else {
                $window.txtLog.AppendText("Update abgebrochen durch Benutzer.`n")
            }
        } else {
            $window.txtLog.AppendText("Keine neue Version verfügbar.`n")
        }
    } catch {
        $window.txtLog.AppendText("Fehler beim Update-Check: $_`n")
    }
}
Check-ForUpdate

# XAML laden
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Fenster anzeigen
$window.ShowDialog() | Out-Null

# Nach dem Laden des Fensters kannst du Log-Ausgaben so hinzufügen:
# $window.txtLog.AppendText("Logeintrag`n")
