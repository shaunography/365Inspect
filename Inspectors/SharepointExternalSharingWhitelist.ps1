$ErrorActionPreference = "Stop"

$errorHandling = "$((Get-Item $PSScriptRoot).Parent.FullName)\Write-ErrorLog.ps1"

. $errorHandling


function Inspect-SharepointExternalSharingWhitelist {
Try {

	If ((Get-SPOTenant).SharingDomainRestrictionMode -eq "None"){
		return "External sharing is not restricted by domain"
	} else {
		If ((Get-SPOTenant).SharingAllowedDomainList){
			return "External sharing is restricted to a domain white list: $((Get-SPOTenant).SharingAllowedDomainList)"
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

return Inspect-SharepointExternalSharingWhitelist


