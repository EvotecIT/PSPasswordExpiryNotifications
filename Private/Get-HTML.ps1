function Get-HTML {
    [CmdletBinding()]
    param (
        [string] $text
    )
    $text = $text.Split("`r")
    foreach ($t in $text) {
        Write-Host $t
    }
}