function Show-TerminalColors {
    <#
    .SYNOPSIS
    Displays all possible foreground and background color combinations in the terminal.
    .DESCRIPTION
    Iterates through all available console colors and prints sample text for each
    foreground/background combination, allowing you to preview how colors appear in your terminal.
    #>

    [cmdletbinding()]
    param()

    $colors = [enum]::GetValues([System.ConsoleColor]) |
        Select-Object @{N = 'ColorObject'; E = { $_ } }, @{N = 'ColorName'; E = { if ($_.ToString().substring(0, 3) -eq 'Dar' ) {
                    $_.ToString().Substring(4) + 'DARK' 
                } else {
                    $_.ToString() 
                } } 
        } | Sort-Object Colorname
    foreach ($bgcolor in $colors.ColorObject) { 
        foreach ($fgcolor in $colors.ColorObject) { 
            Write-Host "$fgcolor|" -ForegroundColor $fgcolor -BackgroundColor $bgcolor -NoNewline 
        } 
        Write-Host " on $bgcolor" 
    }
}
