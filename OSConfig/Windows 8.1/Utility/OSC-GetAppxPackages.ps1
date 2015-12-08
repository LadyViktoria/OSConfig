#https://winpeguy.wordpress.com/my-scripts/tool-osconfig/

Function Use-RunAs 
{    
    # Check if script is running as Adminstrator and if not use RunAs
    # Use Check Switch to check if admin

    param([Switch]$Check) 

    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()` 
        ).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if ($Check) { return $IsAdmin }

    if ($MyInvocation.ScriptName -ne "")
    {
        if (-not $IsAdmin)
        {
            try
            {
                $arg = "-file `"$($MyInvocation.ScriptName)`""
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'
            }
            catch
            {
                Write-Warning "Error - Failed to restart script with runas"
                break
            }
            exit # Quit this session of powershell
        }
    }
    else
    {
        Write-Warning "Error - Script must be saved as a .ps1 file first"  
        break  
    }
}

Use-RunAs 
"Script Running Elevated" 

Get-AppxPackage | Sort Name | Select Name | Out-File -FilePath $PSScriptRoot\OSC-AppxPackage.txt -Verbose
Get-AppxPackage | Sort Name | Out-File -FilePath $PSScriptRoot\OSC-AppxPackage.txt -Append -Verbose

Get-ProvisionedAppxPackage -Online | Sort DisplayName | Select DisplayName | Out-File -FilePath $PSScriptRoot\OSC-AppxPackageProvisioned.txt -Verbose
Get-ProvisionedAppxPackage -Online | Sort DisplayName | Out-File -FilePath $PSScriptRoot\OSC-AppxPackageProvisioned.txt -Append -Verbose

