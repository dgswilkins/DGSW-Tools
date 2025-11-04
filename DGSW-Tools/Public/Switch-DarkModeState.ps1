function Switch-DarkModeState {
    <#
    .SYNOPSIS
    Toggles between Windows dark and light theme modes.
    .DESCRIPTION
    Checks the current Windows theme setting and switches between dark and light modes
    by applying the corresponding theme file. Also attempts to close the System Settings
    window if it appears.
    #>

    [cmdletbinding()]
    param()

    $xpath = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
    $xname = 'AppsUseLightTheme'
    if ((Get-ItemProperty $xpath -Name AppsUseLightTheme).$xname -eq 1) {
        Start-Process 'C:\Windows\Resources\Themes\dark.theme' -Wait # set the dark theme of choice
    } else {
        Start-Process 'C:\Windows\Resources\Themes\aero.theme' -Wait # set the light theme of choice
    }

    for ( $i = 0; $i -lt 100; $i++ ) {
        $systemSetting = Get-Process | Where-Object { $_.ProcessName -eq 'SystemSettings' }
        if ( $systemSetting ) {
            $systemSetting | ForEach-Object { $_.Kill() } # The SystemSetting window may be displayed for seconds before being killed.
            $i = 100
        }
        Start-Sleep -Milliseconds 50
    } # Does anyone know a better way to wait for the Process to start?
}
