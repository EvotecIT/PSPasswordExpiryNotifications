function Test-Prerequisits {
    [CmdletBinding()]
    param()
    try {
        $null = Get-ADForest
    } catch {
        if ($_.Exception -match "Unable to find a default server with Active Directory Web Services running.") {
            Write-Color @script:WriteParameters "[-] ", "Active Directory", " not found. Please run this script with access to ", "Domain Controllers." -Color White, Red, White, Red
        }
        Write-Color @script:WriteParameters "[-] ", "Error: $($_.Exception.Message)" -Color White, Red
        Exit
    }
}