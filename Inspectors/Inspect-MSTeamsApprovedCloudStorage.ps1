$ErrorActionPreference = "Stop"

$errorHandling = "$((Get-Item $PSScriptRoot).Parent.FullName)\Write-ErrorLog.ps1"

. $errorHandling


Function Inspect-MSTeamsAllowedDomains {
Try {

	Try {

         $results = @()

		 $configuration = Get-CsTeamsClientConfiguration

         If (($configuration.AllowDropBox -eq $true) {
             $results += "DropBox Allowed"
         }

         If (($configuration.AllowBox -eq $true) {
             $results += "Box Allowed"
         }

         If (($configuration.AllowGoogleDrive -eq $true) {
             $results += "Google Drive Allowed"
         }

         If (($configuration.AllowShareFile -eq $true) {
             $results += "Share File Allowed"
         }

         If (($configuration.AllowEgnyte -eq $true) {
             $results += "Egnyte Allowed"
         }
         
        If ($results){
            Return $results
        }

	}
	Catch {
        Write-Warning -Message "Error processing request. Manual verification required."
        Return "Error processing request."
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
return Inspect-MSTeamsAllowedDomains


