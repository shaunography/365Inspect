$ErrorActionPreference = "Stop"

$errorHandling = "$((Get-Item $PSScriptRoot).Parent.FullName)\Write-ErrorLog.ps1"

. $errorHandling


$path = @($out_path)

function Azure-StorageAccountDefaultNetworkAccessRule {
Try {
    
    $storageaccounts = Get-AzStorageAccount | Get-AzStorageAccountNetworkRuleSet | where { $_.DefaultAction -eq "Allow" }

    if ($storageaccounts) {
        return $storageaccounts.StorageAccountName
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

return Azure-StorageAccountDefaultNetworkAccessRule


