# Versionsnummer des lokalen Scripts
define $localVersion = "1.0.0"

# URL zur Online-Versionsdatei (z.B. auf GitHub als raw .txt)
$versionUrl = "https://example.com/DATEV-Toolbox-version.txt"  # Hier eigene URL eintragen
$scriptUrl = "https://example.com/DATEV-Toolbox.ps1"           # Hier eigene URL eintragen

function Check-ForUpdate {
    try {
        $remoteVersion = Invoke-RestMethod -Uri $versionUrl -UseBasicParsing
        if ($remoteVersion -and ($remoteVersion -ne $localVersion)) {
            $result = [System.Windows.MessageBox]::Show("Neue Version ($remoteVersion) verfügbar. Jetzt herunterladen und Script ersetzen?", "Update verfügbar", 'YesNo', 'Information')
            if ($result -eq 'Yes') {
                Invoke-WebRequest -Uri $scriptUrl -OutFile $MyInvocation.MyCommand.Definition -UseBasicParsing
                [System.Windows.MessageBox]::Show("Update abgeschlossen. Bitte Script neu starten.", "Update", 'OK', 'Information')
                exit
            }
        }
    } catch {
        # Fehler beim Update-Check ignorieren
    }
}

Check-ForUpdate

Add-Type -AssemblyName PresentationFramework

# XAML-Definition für das Fenster
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="DATEV Toolbox" MinHeight="400" Width="400" ResizeMode="CanMinimize">
    <Grid Margin="10">
        <TabControl Margin="0,0,0,0" VerticalAlignment="Stretch">
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
    </Grid>
</Window>
"@

# XAML laden
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Fenster anzeigen
$window.ShowDialog() | Out-Null
