$ErrorActionPreference = "Stop"

$errorHandling = "$((Get-Item $PSScriptRoot).Parent.FullName)\Write-ErrorLog.ps1"

. $errorHandling


function Inspect-SharepointPublicSites {
Try {

	$sites = Get-SPOSite -limit all

	$affected = @()
	
	foreach ($site in $sites) { 
		try { 
		
			$res = Get-SPOSiteGroup -Site $site.url | Where-Object {$site.Users -Match “spo-grid-all-users”}
		
			if ($res -ne $null) { 
		
				$affected += $site.url
		
			}
		} Catch {

		}

	}
	
	if ( $affected ) {
	
		return $affected
	
	} else {

		return $null

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

return Inspect-SharepointPublicSites


