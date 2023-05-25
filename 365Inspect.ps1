<#
  .SYNOPSIS
  Performs Office 365 security assessment.

  .DESCRIPTION
  Automate the security assessment of Microsoft Office 365 environments.

  .PARAMETER OrgName
  The name of the core organization or "company" of your O365 instance, which will be inspected.

  .PARAMETER OutPath
  The path to a folder where the report generated by 365Inspect will be placed.

  .PARAMETER Auth
  Switch that should be one of the literal values "MFA", "CMDLINE", or "ALREADY_AUTHED".

  .PARAMETER UserPrincipalName
  UserPrincipalName of O365 account.

  .PARAMETER Password
  Password of O365 account.

  .INPUTS
  None. You cannot pipe objects to 365Inspect.ps1.

  .OUTPUTS
  None. 365Inspect.ps1 does not generate any output.

  .EXAMPLE
  PS> .\365Inspect.ps1
#>


param (
    [Parameter(Mandatory = $true,
        HelpMessage = 'Output path for report')]
    [string] $OutPath,
    [Parameter(Mandatory = $true,
        HelpMessage = 'UserPrincipalName required for Exchange Connection')]
    [string] $UserPrincipalName,
    [Parameter(Mandatory = $false,
        HelpMessage = "Report Output Format")]
    [ValidateSet("All", "HTML", "CSV", "XML", "JSON",
        IgnoreCase = $true)]
    [string] $reportType = "All",
    [Parameter(Mandatory = $true,
        HelpMessage = 'Auth type')]
    [ValidateSet('ALREADY_AUTHED', 'MFA',
        IgnoreCase = $false)]
    [string] $Auth = "MFA",
    [string[]] $SelectedInspectors = @(),
    [string[]] $ExcludedInspectors = @()
)

$global:orgInfo = $null
$out_path = $OutPath
$selected_inspectors = $SelectedInspectors
$excluded_inspectors = $ExcludedInspectors

. .\Write-ErrorLog.ps1

$MaximumFunctionCount = 32768

Function Connect-Services {
    # Log into every service prior to the analysis.
    If ($auth -EQ "MFA") {
        Try {
            Write-Output "Connecting to Microsoft Graph"
            Connect-MgGraph -ContextScope Process -Scopes "AuditLog.Read.All", "Policy.Read.All", "Directory.Read.All", "IdentityProvider.Read.All", "Organization.Read.All", "Securityevents.Read.All", "ThreatIndicators.Read.All", "SecurityActions.Read.All", "User.Read.All", "UserAuthenticationMethod.Read.All", "MailboxSettings.Read", "DeviceManagementManagedDevices.Read.All", "DeviceManagementApps.Read.All", "UserAuthenticationMethod.ReadWrite.All", "DeviceManagementServiceConfig.Read.All", "DeviceManagementConfiguration.Read.All"
            #Select-MgProfile -Name beta
            $global:orgInfo = ((Get-MgOrganization).VerifiedDomains | Where-Object { $_.Name -match 'onmicrosoft.com' })[0].Name
            Write-Output "Connected via Graph to $((Get-MgOrganization).DisplayName)"
        }
        Catch {
            Write-Output "Connecting to Microsoft Graph Failed."
            Write-Error $_.Exception.Message
            Break
        }
        Try {
            Write-Output "Connecting to Exchange Online"
            Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false
        }
        Catch {
            Write-Output "Connecting to Exchange Online Failed."
            Write-Error $_.Exception.Message
            Break
        }
        Try {
            Write-Output "Connecting to SharePoint Service"
            #$org_name = ($global:orgInfo -split '.onmicrosoft.com')[0]
            $org_name = ($global:orgInfo -split '\.')[0]
            Connect-SPOService -Url "https://$org_name-admin.sharepoint.com"
        }
        Catch {
            Write-Output "Connecting to SharePoint Service Failed."
            Write-Error $_.Exception.Message
            Break
        }
        Try {
            Write-Output "Connecting to Microsoft Teams"
            Connect-MicrosoftTeams
        }
        Catch {
            Write-Output "Connecting to Microsoft Teams Failed."
            Write-Error $_.Exception.Message
            Break
        }
        Try {
            Write-Output "Connecting to Security and Compliance Center"
            Connect-IPPSSession -UserPrincipalName $UserPrincipalName
        }
        Catch {
            Write-Output "Connecting to Security and Compliance Center Failed."
            Write-Error $_.Exception.Message
            Break
        }
    }
    Else {
        $global:orgInfo = ((Get-MgOrganization).VerifiedDomains | Where-Object { $_.Name -match 'onmicrosoft.com' })[0].Name
    }
}

#Function to change color of text on errors for specific messages
Function Colorize($ForeGroundColor) {
    $color = $Host.UI.RawUI.ForegroundColor
    $Host.UI.RawUI.ForegroundColor = $ForeGroundColor
  
    if ($args) {
        Write-Output $args
    }
  
    $Host.UI.RawUI.ForegroundColor = $color
}


Function Confirm-Close {
    Read-Host "Press Enter to Exit"
    Exit
}

Function Confirm-InstalledModules {
    #Check for required Modules and versions; Prompt for install if missing and import.
    $ExchangeOnlineManagement = @{ Name = "ExchangeOnlineManagement"; MinimumVersion = "2.0.5" }
    $SharePoint = @{ Name = "Microsoft.Online.SharePoint.PowerShell"; MinimumVersion = "16.0.22601.12000" }
    $Graph = @{ Name = "Microsoft.Graph"; MinimumVersion = "1.9.6" }
    $MSTeams = @{ Name = "MicrosoftTeams"; MinimumVersion = "4.4.1" }
    $psGet = @{ Name = "PowerShellGet"; RequiredVersion = "2.2.5" }

    #Try {
    #    $psGetVersion = Get-InstalledModule -Name PowerShellGet -ErrorAction Stop
#
    #    If ($psGetVersion.Version -lt '2.2.5') {
    #        Write-Host "[-] " -ForegroundColor Red -NoNewline
    #        Write-Warning "PowerShellGet is not the correct version. Please install using the following command:"
    #        Write-Host "Update-Module " -ForegroundColor Yellow -NoNewline
    #        Write-Host "-Name " -ForegroundColor Gray -NoNewline
    #        Write-Host "PowerShellGet " -ForegroundColor White -NoNewline
    #        Write-Host '-Force' -ForegroundColor Gray
    #        $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    #        if (-not $IsAdmin) {
    #            Write-Warning "PowerShellGet is not the correct version. Please install using the following command:"
    #            Write-Host "Update-Module " -ForegroundColor Yellow -NoNewline
    #            Write-Host "-Name " -ForegroundColor Gray -NoNewline
    #            Write-Host "PowerShellGet " -ForegroundColor White -NoNewline
    #            Write-Host '-Force' -ForegroundColor Gray
    #        }
    #        Else {
    #            Write-Host "Installing PowerShellGet`n" -ForegroundColor Magenta
    #            Install-Module -Name 'PowerShellGet' -AllowPrerelease -AllowClobber -Force -MinimumVersion '2.2.5'
    #        }
    #    }
    #}
    #Catch {
    #    $exc = $_
    #    if ($exc -like "*No match was found for the specified search criteria and module names 'powershellget'*") {
    #        Write-Host "[-] " -ForegroundColor Red -NoNewline
    #        Write-Warning "PowerShellGet was not installed via PowerShell Gallery. Please install using the following command:"
    #        Write-Host "Install-Module " -ForegroundColor Yellow -NoNewline
    #        Write-Host "-Name " -ForegroundColor Gray -NoNewline
    #        Write-Host "PowerShellGet " -ForegroundColor White -NoNewline
    #        Write-Host "-RequiredVersion " -ForegroundColor Gray -NoNewline
    #        Write-Host '2.2.5 ' -ForegroundColor White -NoNewline
    #        Write-Host '-Force' -ForegroundColor Gray
    #    }
    #}

    $modules = @($psGet, $ExchangeOnlineManagement, $Graph, $SharePoint, $MSTeams)
    $count = 0

    Write-Output "Verifying environment. `n"

    foreach ($module in $modules) {
        $installedVersion = [Version](((Get-InstalledModule -Name $module.Name).Version -split "-")[0])

        If (($module.Name -eq (Get-InstalledModule -Name $module.Name).Name) -and (([Version]$module.MinimumVersion -le $installedVersion))) {
            If ($PSVersionTable.PSVersion.Major -eq 5) {
                Write-Host "Environment is $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
                Write-Host "`t[+] " -NoNewLine -ForeGroundColor Green
                Write-Output "$($module.Name) is installed."
                
                If ($module.Name -ne 'Microsoft.Graph') {
                    Write-Host "`tImporting $($module.Name)" -ForeGroundColor Green
                    Import-Module -Name $module.Name | Out-Null
                }
                Else {
                    Write-Host "`tImporting Microsoft.Graph" -ForeGroundColor Green
                    Import-Module -Name Microsoft.Graph.Identity.DirectoryManagement | Out-Null
                    Import-Module -Name Microsoft.Graph.Identity.SignIns | Out-Null
                    Import-Module -Name Microsoft.Graph.Users | Out-Null
                    Import-Module -Name Microsoft.Graph.Applications | Out-Null
                }
            }
            Elseif ($PSVersionTable.PSVersion.Major -ge 6) {
                If ($IsWindows) {
                    Write-Host "Environment is $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
                    Write-Host "`t[+] " -NoNewLine -ForeGroundColor Green
                    Write-Output "$($module.Name) is installed."

                    If (($module.Name -ne 'Microsoft.Graph') -and ($module.Name -ne 'ExchangeOnlineManagement')) {
                        Try {
                            Write-Host "`tImporting $($module.Name)" -ForeGroundColor Green
                            Import-Module -Name $module.Name -UseWindowsPowerShell -WarningAction SilentlyContinue | Out-Null
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
                            Write-ErrorLog -message $message -exception $exception -scriptname $scriptname -failinglinenumber $failinglinenumber -failingline $failingline -pscommandpath $pscommandpath -positionmsg $pscommandpath -stacktrace $strace
                            Write-Verbose "Errors written to log"
                        }
                    }
                    Else {
                        Try {
                            Write-Host "`tInporting ExchangeOnlineManagement"
                            Import-Module -Name ExchangeOnlineManagement | Out-Null
                            Write-Host "`tImporting Microsoft.Graph" -ForeGroundColor Green
                            Import-Module -Name Microsoft.Graph.Identity.DirectoryManagement | Out-Null
                            Import-Module -Name Microsoft.Graph.Identity.SignIns | Out-Null
                            Import-Module -Name Microsoft.Graph.Users | Out-Null
                            Import-Module -Name Microsoft.Graph.Applications | Out-Null
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
                            Write-ErrorLog -message $message -exception $exception -scriptname $scriptname -failinglinenumber $failinglinenumber -failingline $failingline -pscommandpath $pscommandpath -positionmsg $pscommandpath -stacktrace $strace
                            Write-Verbose "Errors written to log"
                        }
                    }
                }
                Else {
                    Write-Host "We're sorry, due to various module dependency requirements, this tool will not work on a non-Windows operating system." -ForegroundColor Yellow
                    Exit
                }
            }
            $count ++
        }
        Else {
            $message = Write-Output "`n$($module.Name) is not installed."
            $message1 = Write-Output "The module may be installed by running `"Install-Module -Name $($module.Name) -AllowPrerelease -AllowClobber -Force -MinimumVersion $($module)`" in an elevated PowerShell window."
            Colorize Red ($message)
            Colorize Yellow ($message1)
            $install = Read-Host -Prompt "Would you like to attempt installation now? (Y|N)"
            If ($install -eq 'y') {
                Install-Module -Name $module.Name -AllowPrerelease -AllowClobber -Scope CurrentUser -Force -MinimumVersion $module
                $count ++
            }
        }
    }

    If ($count -lt 5) {
        Write-Output ""
        Write-Output ""
        $message = Write-Output "Dependency checks failed. Please install all missing modules before running this script."
        Colorize Red ($message)
        Confirm-Close
    }
    Else {
        Connect-Services
    }
}


If ($Auth -eq 'ALREADY_AUTHED') {
    Connect-Services
}
Else {
    #Start Script
    Confirm-InstalledModules
}

# Obtain tenant info
$org_name = ($global:orgInfo -split '.onmicrosoft')
$tenantDisplayName = ($global:orgInfo).DisplayName

# Get a list of every available detection module by parsing the PowerShell
# scripts present in the .\inspectors folder. 
#Exclude specified Inspectors
If ($excluded_inspectors -and $excluded_inspectors.Count) {
    $excluded_inspectors = foreach ($inspector in $excluded_inspectors) { "$inspector.ps1" }
    $inspectors = (Get-ChildItem .\inspectors\*.ps1 -exclude $excluded_inspectors).Name | ForEach-Object { ($_ -split ".ps1")[0] }
}
else {
    $inspectors = (Get-ChildItem .\inspectors\*.ps1).Name | ForEach-Object { ($_ -split ".ps1")[0] }
}

#Use Selected Inspectors
If ($selected_inspectors -AND $selected_inspectors.Count) {
    "The following inspectors were selected for use: "
    Foreach ($inspector in $selected_inspectors) {
        Write-Output $inspector
    }
}
elseif ($excluded_Inspectors -and $excluded_inspectors.Count) {
    $selected_inspectors = $inspectors
    Write-Output "Using inspectors:`n"
    Foreach ($inspector in $inspectors) {
        Write-Output $inspector
    }
}
Else {
    "Using all inspectors."
    $selected_inspectors = $inspectors
}

#Create Output Directory if required
Try {
    New-Item -ItemType Directory -Force -Path $out_path | Out-Null
    If ((Test-Path $out_path) -eq $true) {
        $path = Resolve-Path $out_path
        Write-Output "$($path.Path) created successfully."
    }
}
Catch {
    Write-Error "Directory not created. Please check permissions."
    Confirm-Close
}

# Maintain a list of all findings, beginning with an empty list.
$findings = @()

# For every inspector the user wanted to run...
ForEach ($selected_inspector in $selected_inspectors) {
    # ...if the user selected a valid inspector...
    If ($inspectors.Contains($selected_inspector)) {
        Write-Output "Invoking Inspector: $selected_inspector"
		
        # Get the static data (finding description, remediation etc.) associated with that inspector module.
        $finding = Get-Content .\inspectors\$selected_inspector.json | Out-String | ConvertFrom-Json
		
        # Invoke the actual inspector module and store the resulting list of insecure objects.
        $finding.AffectedObjects = Invoke-Expression ".\inspectors\$selected_inspector.ps1"
		
        # Add the finding to the list of all findings.
        $findings += $finding
    }
}

# Function that retrieves templating information from 
Function HTML-Report {
    # Function that retrieves templating information from 
    function Parse-Template {
        $template = (Get-Content ".\365InspectDefaultTemplate.html") -join "`n"
        $template -match '\<!--BEGIN_FINDING_LONG_REPEATER-->([\s\S]*)\<!--END_FINDING_LONG_REPEATER-->'
        $findings_long_template = $matches[1]
        
        $template -match '\<!--BEGIN_FINDING_SHORT_REPEATER-->([\s\S]*)\<!--END_FINDING_SHORT_REPEATER-->'
        $findings_short_template = $matches[1]
        
        $template -match '\<!--BEGIN_AFFECTED_OBJECTS_REPEATER-->([\s\S]*)\<!--END_AFFECTED_OBJECTS_REPEATER-->'
        $affected_objects_template = $matches[1]
        
        $template -match '\<!--BEGIN_REFERENCES_REPEATER-->([\s\S]*)\<!--END_REFERENCES_REPEATER-->'
        $references_template = $matches[1]
        
        $template -match '\<!--BEGIN_EXECSUM_TEMPLATE-->([\s\S]*)\<!--END_EXECSUM_TEMPLATE-->'
        $execsum_template = $matches[1]
        
        return @{
            FindingShortTemplate    = $findings_short_template;
            FindingLongTemplate     = $findings_long_template;
            AffectedObjectsTemplate = $affected_objects_template;
            ReportTemplate          = $template;
            ReferencesTemplate      = $references_template;
            ExecsumTemplate         = $execsum_template
        }
    }
    
    $templates = Parse-Template
    
    # Maintain a running list of each finding, represented as HTML
    $short_findings_html = "" 
    $long_findings_html = ""
    
    $findings_count = 0
    
    #$sortedFindings1 = $findings | Sort-Object {$_.FindingName}
    $sortedFindings = $findings | Sort-Object { Switch -Regex ($_.Impact) { 'Critical' { 1 }	'High' { 2 }	'Medium' { 3 }	'Low' { 4 }	'Informational' { 5 } }; $_.FindingName } 
    ForEach ($finding in $sortedFindings) {
        # If the result from the inspector was not $null,
        # it identified a real finding that we must process.
        If ($null -NE $finding.AffectedObjects) {
            # Increment total count of findings
            $findings_count += 1
            
            # Keep an HTML variable representing the current finding as HTML
            $short_finding_html = $templates.FindingShortTemplate
            $long_finding_html = $templates.FindingLongTemplate
            
            # Insert finding name and number into template HTML
            $short_finding_html = $short_finding_html.Replace("{{FINDING_NAME}}", $finding.FindingName)
            $short_finding_html = $short_finding_html.Replace("{{FINDING_NUMBER}}", $findings_count.ToString())
            $short_finding_html = $short_finding_html.Replace("{{CIS}}", $findings.CIS)
            $long_finding_html = $long_finding_html.Replace("{{FINDING_NAME}}", $finding.FindingName)
            $long_finding_html = $long_finding_html.Replace("{{FINDING_NUMBER}}", $findings_count.ToString())
            
            # Finding Impact
            If ($finding.Impact -eq 'Critical') {
                $htmlImpact = '<span style="color:Crimson;"><strong>Critical</strong></span>'
                $short_finding_html = $short_finding_html.Replace("{{IMPACT}}", $htmlImpact)
                $long_finding_html = $long_finding_html.Replace("{{IMPACT}}", $htmlImpact)
            }
            ElseIf ($finding.Impact -eq 'High') {
                $htmlImpact = '<span style="color:DarkOrange;"><strong>High</strong></span>'
                $short_finding_html = $short_finding_html.Replace("{{IMPACT}}", $htmlImpact)
                $long_finding_html = $long_finding_html.Replace("{{IMPACT}}", $htmlImpact)
            }
            Else {
                $short_finding_html = $short_finding_html.Replace("{{IMPACT}}", $finding.Impact)
                $long_finding_html = $long_finding_html.Replace("{{IMPACT}}", $finding.Impact)
            }
            $short_finding_html = $short_finding_html.Replace("{{RISKRATING}}", $finding.RiskRating)
            $long_finding_html = $long_finding_html.Replace("{{RISKRATING}}", $finding.RiskRating)
            
            # Finding description
            $long_finding_html = $long_finding_html.Replace("{{DESCRIPTION}}", $finding.Description)
    
            # Finding default value
            $long_finding_html = $long_finding_html.Replace("{{DEFAULTVALUE}}", $finding.DefaultValue)
    
            # Finding expected value
            $long_finding_html = $long_finding_html.Replace("{{EXPECTEDVALUE}}", $finding.ExpectedValue)
                    
            # Finding Remediation
            If ($finding.Remediation.length -GT 300) {
                $short_finding_text = "Complete remediation advice is provided in the body of the report. Clicking the link to the left will take you there."
            }
            Else {
                $short_finding_text = $finding.Remediation
            }
            
            $short_finding_html = $short_finding_html.Replace("{{REMEDIATION}}", $short_finding_text)
            $long_finding_html = $long_finding_html.Replace("{{REMEDIATION}}", $finding.Remediation)
            
            # Affected Objects
            If ($finding.AffectedObjects.Count -GT 25) {
                $condensed = "<a href='{name}'>{count} Affected Objects Identified<a/>."
                $condensed = $condensed.Replace("{count}", $finding.AffectedObjects.Count.ToString())
                $condensed = $condensed.Replace("{name}", $finding.FindingName)
                $affected_object_html = $templates.AffectedObjectsTemplate.Replace("{{AFFECTED_OBJECT}}", $condensed)
                $fname = $finding.FindingName
                $finding.AffectedObjects | Out-File -FilePath "$out_path\$fname.txt"
            }
            Else {
                $affected_object_html = ''
                ForEach ($affected_object in $finding.AffectedObjects) {
                    $affected_object_html += $templates.AffectedObjectsTemplate.Replace("{{AFFECTED_OBJECT}}", $affected_object)
                }
            }
            
            $long_finding_html = $long_finding_html.Replace($templates.AffectedObjectsTemplate, $affected_object_html)
            
            # References
            $reference_html = ''
            ForEach ($reference in $finding.References) {
                $this_reference = $templates.ReferencesTemplate.Replace("{{REFERENCE_URL}}", $reference.Url)
                $this_reference = $this_reference.Replace("{{REFERENCE_TEXT}}", $reference.Text)
                $reference_html += $this_reference
            }
            
            $long_finding_html = $long_finding_html.Replace($templates.ReferencesTemplate, $reference_html)
            
            # Add the completed short and long findings to the running list of findings (in HTML)
            $short_findings_html += $short_finding_html
            $long_findings_html += $long_finding_html
        }
    }
    
    # Insert command line execution information. This is coupled kinda badly, as is the Affected Objects html.
    $flags = "<b>Prepared for organization:</b><br/>" + $org_name + "<br/><br/>"
    $flags = $flags + "<b>Stats</b>:<br/> <b>" + $findings_count + "</b> out of <b>" + $inspectors.Count + "</b> executed inspector modules identified possible opportunities for improvement.<br/><br/>"  
    $flags = $flags + "<b>Inspector Modules Executed</b>:<br/>" + [String]::Join("<br/>", $selected_inspectors)
    
    $output = $templates.ReportTemplate.Replace($templates.FindingShortTemplate, $short_findings_html)
    $output = $output.Replace($templates.FindingLongTemplate, $long_findings_html)
    $output = $output.Replace($templates.ExecsumTemplate, $templates.ExecsumTemplate.Replace("{{CMDLINEFLAGS}}", $flags))
    
    $output | Out-File -FilePath $out_path\Report_$(Get-Date -Format "yyyy-MM-dd_hh-mm-ss").html
}

Function CSV-Report {
    $sortedFindings = $findings | Sort-Object { Switch -Regex ($_.Impact) { 'Critical' { 1 }	'High' { 2 }	'Medium' { 3 }	'Low' { 4 }	'Informational' { 5 } }; $_.FindingName }

    $results = @()

    $findings_count = 0

    foreach ($finding in $sortedFindings) {
        If ($null -NE $finding.AffectedObjects) {
            $findings_count += 1

            $refs = @()

            foreach ($ref in $finding.References) {
                $refs += "$($ref.Text) : $($ref.Url)"
            }

            $result = New-Object psobject
            $result | Add-Member -MemberType NoteProperty -name ID -Value $findings_count.ToString() -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name FindingName -Value $finding.FindingName -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name AffectedObjects -Value $("$($finding.AffectedObjects)" | Out-String).Trim() -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name Finding -Value $(($finding.Description) -join " ") -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name DefaultValue -Value $(($finding.DefaultValue) -join " ") -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name ExpectedValue -Value $(($finding.ExpectedValue) -join " ") -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name InherentRisk -Value $finding.Impact -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name 'Residual Risk' -Value " "
            $result | Add-Member -MemberType NoteProperty -name Remediation -Value $(($finding.Remediation) -join " ") -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name References -Value $(($refs) -join ';')  -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name 'Remediation Status' -Value " "
            $result | Add-Member -MemberType NoteProperty -name 'Required Resources' -Value " "
            $result | Add-Member -MemberType NoteProperty -name 'Start Date' -Value " "
            $result | Add-Member -MemberType NoteProperty -name 'Completion Date' -Value " "
            $result | Add-Member -MemberType NoteProperty -name 'Notes' -Value " "
            
            $results += $result
        }
    }

    $results | Export-Csv "$out_path\Report_$(Get-Date -Format "yyyy-MM-dd_hh-mm-ss").csv" -Delimiter '^' -NoTypeInformation -Append -Force

}

Function XML-Report {
    $sortedFindings = $findings | Sort-Object { Switch -Regex ($_.Impact) { 'Critical' { 1 }	'High' { 2 }	'Medium' { 3 }	'Low' { 4 }	'Informational' { 5 } }; $_.FindingName }

    $results = @()

    $findings_count = 0

    foreach ($finding in $sortedFindings) {
        If ($null -NE $finding.AffectedObjects) {
            $findings_count += 1

            $refs = @()

            foreach ($ref in $finding.References) {
                $refs += "$($ref.Text) : $($ref.Url)"
            }

            $result = New-Object psobject
            $result | Add-Member -MemberType NoteProperty -name ID -Value $findings_count.ToString() -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name FindingName -Value $finding.FindingName -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name AffectedObjects -Value $("$($finding.AffectedObjects)" | Out-String).Trim() -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name Finding -Value $finding.Description -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name DefaultValue -Value $finding.DefaultValue -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name ExpectedValue -Value $finding.ExpectedValue -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name InherentRisk -Value $finding.Impact -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name 'Residual Risk' -Value " "
            $result | Add-Member -MemberType NoteProperty -name Remediation -Value $finding.Remediation -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name References -Value $($refs | Out-String)  -ErrorAction SilentlyContinue
            
            $results += $result
        }
    }

    $results | Export-Clixml -Depth 3 -Path "$out_path\Report_$(Get-Date -Format "yyyy-MM-dd_hh-mm-ss").xml"
}

Function JSON-Report {
    $sortedFindings = $findings | Sort-Object { Switch -Regex ($_.Impact) { 'Critical' { 1 }	'High' { 2 }	'Medium' { 3 }	'Low' { 4 }	'Informational' { 5 } }; $_.FindingName }

    $results = @()

    $findings_count = 0

    foreach ($finding in $sortedFindings) {
        If ($null -NE $finding.AffectedObjects) {
            $findings_count += 1

            $refs = @()

            foreach ($ref in $finding.References) {
                $refs += "$($ref.Text) : $($ref.Url)"
            }

            $result = New-Object psobject
            $result | Add-Member -MemberType NoteProperty -name ID -Value $findings_count.ToString() -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name FindingName -Value $finding.FindingName -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name AffectedObjects -Value $("$($finding.AffectedObjects)" | Out-String).Trim() -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name Finding -Value $finding.Description -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name DefaultValue -Value $finding.DefaultValue -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name ExpectedValue -Value $finding.ExpectedValue -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name InherentRisk -Value $finding.Impact -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name 'Residual Risk' -Value " "
            $result | Add-Member -MemberType NoteProperty -name Remediation -Value $finding.Remediation -ErrorAction SilentlyContinue
            $result | Add-Member -MemberType NoteProperty -name References -Value $($refs | Out-String)  -ErrorAction SilentlyContinue
            
            $results += $result
        }
    }

    $results | ConvertTo-Json | Out-File -FilePath $out_path\Report_$(Get-Date -Format "yyyy-MM-dd_hh-mm-ss").json
}

Function All-Report {
    CSV-Report
    XML-Report
    JSON-Report
    HTML-Report
}

If ($reportType -eq "HTML") {
    HTML-Report
}
Elseif ($reportType -eq "CSV") {
    CSV-Report
}
Elseif ($reportType -eq "XML") {
    XML-Report
}
Elseif ($reportType -eq "JSON") {
    JSON-Report
}
Else {
    All-Report
}

$compress = @{
    Path             = $out_path
    CompressionLevel = "Fastest"
    DestinationPath  = "$out_path\$($org_name)_Report.zip"
}
Compress-Archive @compress

function Disconnect {
    <#Write-Output "Disconnect from MSOnline Service"
	[Microsoft.Online.Administration.Automation.ConnectMsolService]::ClearUserSessionState()#>
    Write-Output "Disconnect from Azure Active Directory"
    Disconnect-AzureAD
    Write-Output "Disconnect from Exchange Online"
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Output "Disconnect from SharePoint Service"
    Disconnect-SPOService
    Write-Output "Disconnect from Microsoft Teams"
    Disconnect-MicrosoftTeams
    Write-Output "Disconnect from Microsoft Intune"
    Write-Output "Disconnect from Microsoft Graph"
    Disconnect-MgGraph
}

$removeSession = Read-Host -Prompt "Do you wish to disconnect your session? (Y|N)"

If ($removeSession -ne 'n') {
    Disconnect
}


return