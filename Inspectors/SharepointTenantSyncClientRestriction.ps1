$ErrorActionPreference = "Stop"

$errorHandling = "$((Get-Item $PSScriptRoot).Parent.FullName)\Write-ErrorLog.ps1"

. $errorHandling


function Inspect-SharepointTenantSyncClientRestriction {
Try {

	If ((Get-SPOTenantSyncClientRestriction).TenantRestrictionEnabled -eq $false){
		return "Not restrictions configured for synching with unmanged devices"
	} else {
		If ((Get-SPOTenant).AllowedDomainList){
			return "syncing is restricted to a domain white list: $((Get-SPOTenant).AllowedDomainList)"
		}
		
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

return Inspect-SharepointTenantSyncClientRestriction


