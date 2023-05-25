$ErrorActionPreference = "Stop"

$errorHandling = "$((Get-Item $PSScriptRoot).Parent.FullName)\Write-ErrorLog.ps1"

. $errorHandling


function Inspect-MailTipsEnabled {
Try {

	$results = @()

	$configuration = Get-OrganizationConfig

	If ($configuration.MailTipsAllTipsEnabled -eq $false) {
		$results += "MailTipsAllTipsEnabled = False)"
	}

	If ($configuration.MailTipsExternalRecipientsTipsEnabled -eq $false) {
		$results += "MailTipsExternalRecipientsTipsEnabled : False)"
	}

	If ($configuration.MailTipsGroupMetricsEnabled -eq $false) {
		$results += "MailTipsGroupMetricsEnabled : False)"
	}

	If ($configuration.MailTipsLargeAudienceThreshold -gt 24) {
		$results += "MailTipsLargeAudienceThreshold : $($configuration.MailTipsLargeAudienceThreshold)"
	}
	
If ($results){
	Return $results
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

return Inspect-MailTipsEnabled


