$ErrorActionPreference = "Stop"

$errorHandling = "$((Get-Item $PSScriptRoot).Parent.FullName)\Write-ErrorLog.ps1"

. $errorHandling


$path = @($out_path)

function Azure-ConditionalAccessTrustedLocations {
Try {
    $tenantLicense = (Get-MgSubscribedSku).ServicePlans
    
    If ($tenantLicense.ServicePlanName -match "AAD_PREMIUM*") {
        
        $secureDefault = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -Property IsEnabled | Select-Object IsEnabled
        $conditionalAccess = Get-MgIdentityConditionalAccessPolicy

        If ($secureDefault.IsEnabled -eq $true) {
            Return $null
        }
        ElseIf (($secureDefault.IsEnabled -eq $false) -and ($conditionalAccess.count -eq 0)) {
            return $false
        }
        else {
            $locations = Get-MgIdentityConditionalAccessNamedLocation

            if ($locations) {
                return $locations.DisplayName
            }
            
       
        }
    }
    Else {
        Return "Tenant is not licensed for Conditional Access."
    }

}
Catch {
Write-Warning "Error message: $_"
$message = $_.ToString()
$exception = $_.Exception
$strace = $_.ScriptStackTrace
$failingline = $_.InvocationInfo.Line
$positionmsg = $_.InvocationInfo.PositionMessage
$pscommandpath = $_.InvocationInfo.PSCommandPath
$failinglinenumber = $_.InvocationInfo.ScriptLineNumber
$scriptname = $_.InvocationInfo.ScriptName
Write-Verbose "Write to log"
Write-ErrorLog -message $message -exception $exception -scriptname $scriptname
Write-Verbose "Errors written to log"
}

}

return Azure-ConditionalAccessTrustedLocations


