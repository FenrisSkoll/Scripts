<#PSScriptInfo
.VERSION 2.0
.GUID 2b6b6c5a-0b1e-4f8c-9f41-7bf3b14b5c33
.AUTHOR You+ChatGPT
.COPYRIGHT (c) You
#>

<# 
.SYNOPSIS
SharePoint Online reporting tool (PnP.PowerShell + SPO Mgmt Shell), refactored for modularity & automation.

.DESCRIPTION
Generates:
- Inactive Sites
- Sharing Settings
- Duplicate Files
- Large Files (>= threshold)
- User Access Review
- External Users by Site

Improvements:
- Non-interactive mode with -OutDir (CI/Intune friendly)
- Centralised connection & elevation with tracked cleanup
- Single site or whole-tenant selectors
- Consistent CSV exports & error capture
- WhatIf/Confirm, progress, cancellation
- Optional App-Only Cert auth (keeps Interactive as default)

.REQUIRES
PowerShell 7.3+
Modules: PnP.PowerShell, Microsoft.Online.SharePoint.PowerShell

#>

#Requires -Version 7.3
[CmdletBinding()]
param()

#region ========= Settings / Globals =========
$script:AdminUrl                = $null
$script:ClientId                = $null
$script:Tenant                  = $null
$script:AuthMode                = 'Interactive' # 'Interactive' | 'AppCert'
$script:CertPath                = $null
$script:CertPassword            = $null
$script:AddedOwnerSites         = [System.Collections.Generic.HashSet[string]]::new()
$script:CancellationRequested   = $false
$ErrorActionPreference          = 'Stop'
#endregion

#region ========= Utility =========
function Write-Info    { param($m) Write-Host $m -ForegroundColor Cyan   }
function Write-Good    { param($m) Write-Host $m -ForegroundColor Green  }
function Write-Bad     { param($m) Write-Host $m -ForegroundColor Red    }
function Write-Warn2   { param($m) Write-Warning $m }

function Stop-IfCancelled {
    if ($script:CancellationRequested) { throw 'UserCancelled' }
    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq 'Q') { throw 'UserCancelled' }
    }
}

function Ensure-Modules {
    $mods = @('PnP.PowerShell','Microsoft.Online.SharePoint.PowerShell')
    foreach ($m in $mods) {
        if (-not (Get-Module -ListAvailable -Name $m)) {
            Write-Host "Missing module: $m" -ForegroundColor Yellow
            Install-Module -Name $m -Scope CurrentUser -Force -ErrorAction Stop
            Write-Good "Installed $m"
        }
    }
}

function Get-Timestamp { (Get-Date).ToString('yyyy-MM-dd_HH-mm-ss') }

function Resolve-PathForExport {
    param(
        [string]$OutDir,
        [string]$DefaultFileName
    )
    if ($OutDir) {
        if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
        return (Join-Path $OutDir $DefaultFileName)
    }

    # Windows-only SaveFileDialog when interactive
    if ($IsWindows) {
        try {
            Add-Type -AssemblyName System.Windows.Forms
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Title = "Save CSV"
            $dlg.Filter = "CSV files (*.csv)|*.csv"
            $dlg.FileName = $DefaultFileName
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                return $dlg.FileName
            } else {
                Write-Warn2 "Save cancelled."
                return $null
            }
        } catch {
            Write-Warn2 "Save dialog failed, falling back to current directory."
        }
    }
    return (Join-Path (Get-Location) $DefaultFileName)
}

function Export-ReportCsv {
    param(
        [Parameter(Mandatory)] $Data,
        [Parameter(Mandatory)][string] $Path
    )
    if ($null -eq $Data) { return }
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $Data | Export-Csv -Path $Path -NoTypeInformation -UseQuotes AsNeeded
    Write-Good "Saved: $Path"
}

#endregion

#region ========= Auth / Connections =========
function Set-SPOConnectionSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [string]$Tenant,
        [ValidateSet('Interactive','AppCert')] [string]$AuthMode = 'Interactive',
        [string]$CertificatePath,
        [securestring]$CertificatePassword
    )
    $script:AdminUrl     = $AdminUrl
    $script:ClientId     = $ClientId
    $script:Tenant       = $Tenant
    $script:AuthMode     = $AuthMode
    $script:CertPath     = $CertificatePath
    $script:CertPassword = $CertificatePassword
    Write-Good "Connection settings updated."
}

function Connect-SPO {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url
    )
    Stop-IfCancelled

    $params = @{
        Url = $Url
    }

    switch ($script:AuthMode) {
        'Interactive' {
            $params.ClientId    = $script:ClientId
            $params.Interactive = $true
        }
        'AppCert' {
            if (-not $script:Tenant -or -not $script:CertPath) {
                throw "AppCert auth requires Tenant and CertificatePath."
            }
            $params.ClientId          = $script:ClientId
            $params.Tenant            = $script:Tenant
            $params.CertificatePath   = $script:CertPath
            if ($script:CertPassword) { $params.CertificatePassword = $script:CertPassword }
        }
    }

    Connect-PnPOnline @params
}

#endregion

#region ========= Elevation (Owner add/remove) =========
function Add-OwnerIfNeeded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$AdminUpn
    )
    Write-Info "🔐 Elevating on $SiteUrl for $AdminUpn"
    # connect to admin
    Connect-SPO -Url $script:AdminUrl
    $ts = Get-PnPTenantSite -Identity $SiteUrl

    $owners = @()
    if ($ts.Owners) { $owners = @($ts.Owners) }
    if ($owners -notcontains $AdminUpn) {
        $new = $owners + $AdminUpn
        Set-PnPTenantSite -Identity $SiteUrl -Owners $new
        [void]$script:AddedOwnerSites.Add($SiteUrl)
        Write-Good "Added Owner to $SiteUrl"
    } else {
        # still mark it for cleanup so we keep semantics consistent
        [void]$script:AddedOwnerSites.Add($SiteUrl)
    }
}

function Remove-TempOwners {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminUpn
    )
    if ($script:AddedOwnerSites.Count -eq 0) { return }
    Write-Info "🔽 Reverting temporary owner on $($script:AddedOwnerSites.Count) sites..."
    foreach ($site in $script:AddedOwnerSites) {
        try {
            Connect-SPO -Url $script:AdminUrl
            $ts = Get-PnPTenantSite -Identity $site
            $owners = @()
            if ($ts.Owners) { $owners = @($ts.Owners) }
            if ($owners -contains $AdminUpn) {
                $new = $owners | Where-Object { $_ -ne $AdminUpn }
                Set-PnPTenantSite -Identity $site -Owners $new
                Write-Host "Removed $AdminUpn from $site" -ForegroundColor DarkYellow
            }
        } catch {
            Write-Warn2 "Cleanup failed on $site : $($_.Exception.Message)"
        }
    }
    $script:AddedOwnerSites.Clear() | Out-Null
}
#endregion

#region ========= Site & Library helpers =========
function Get-AllSites {
    [CmdletBinding()]
    param(
        [switch]$IncludeOneDrive
    )
    Connect-SPO -Url $script:AdminUrl
    $sites = if ($IncludeOneDrive) {
        Get-PnPTenantSite -IncludeOneDriveSites
    } else {
        Get-PnPTenantSite
    }
    $sites | Where-Object { $_.Template -notlike 'SPSPERS*' -and $_.Template -notlike 'APPCATALOG*' }
}

function Invoke-OnSite {
    <#
      Wraps a site operation with: connect, catch unauthorized -> elevate (optional) -> retry
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$AdminUpn
    )
    try {
        Connect-SPO -Url $SiteUrl
        & $ScriptBlock
    } catch {
        $msg = $_.Exception.Message
        if ($AdminUpn -and ($msg -match 'unauthorized operation|access is denied')) {
            Add-OwnerIfNeeded -SiteUrl $SiteUrl -AdminUpn $AdminUpn
            Connect-SPO -Url $SiteUrl
            & $ScriptBlock
        } else {
            throw
        }
    }
}

function Get-DocumentLibraries {
    [CmdletBinding()]
    param(
        [switch]$AllLibraries
    )
    $lists = Get-PnPList | Where-Object { $_.BaseType -eq 'DocumentLibrary' -and -not $_.Hidden }
    if (-not $AllLibraries) {
        $lists = $lists | Where-Object { $_.Title -eq 'Documents' }
    }
    return $lists
}

#endregion

#region ========= Reports =========

function New-InactiveSitesReport {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$OutPath,
        [string]$AdminUpn
    )
    $rows = [System.Collections.Generic.List[object]]::new()
    $sites = Get-AllSites
    $i=0
    foreach ($s in $sites) {
        Stop-IfCancelled
        $i++
        Write-Info "($i/$($sites.Count)) $($s.Url)"
        try {
            Invoke-OnSite -SiteUrl $s.Url -AdminUpn $AdminUpn -ScriptBlock {
                $web = Get-PnPWeb -Includes LastItemUserModifiedDate
                $rows.Add([pscustomobject]@{
                    SiteUrl      = $s.Url
                    LastModified = $web.LastItemUserModifiedDate
                })
            }
        } catch {
            Write-Warn2 "Skip $($s.Url): $($_.Exception.Message)"
        }
    }
    Export-ReportCsv -Data $rows -Path $OutPath
}

function New-SharingSettingsReport {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$OutPath
    )
    $rows = [System.Collections.Generic.List[object]]::new()
    $sites = Get-AllSites
    $i=0
    foreach ($s in $sites) {
        Stop-IfCancelled
        $i++
        Write-Info "($i/$($sites.Count)) $($s.Url)"
        $rows.Add([pscustomobject]@{
            Url               = $s.Url
            SharingCapability = $s.SharingCapability
        })
    }
    Export-ReportCsv -Data $rows -Path $OutPath
}

function New-FileScanReport {
    <#
      ReportMode: 'Duplicates' or 'AllFiles'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateSet('Duplicates','AllFiles')]$ReportMode,
        [Parameter(Mandatory)][int]$MinFileSizeMB,
        [Parameter(Mandatory)][string]$OutPath,
        [Parameter(Mandatory)][string]$ErrorPath,
        [string]$AdminUpn,
        [switch]$AllLibraries
    )

    $sizeThreshold = $MinFileSizeMB * 1MB
    $results       = [System.Collections.Generic.List[object]]::new()
    $errors        = [System.Collections.Generic.List[object]]::new()

    $sites = Get-AllSites
    $i=0
    foreach ($s in $sites) {
        Stop-IfCancelled
        $i++
        Write-Info "($i/$($sites.Count)) $($s.Url)"

        $scriptBlock = {
            $libs = Get-DocumentLibraries -AllLibraries:$AllLibraries
            foreach ($lib in $libs) {
                Stop-IfCancelled
                Write-Host "  • $($lib.Title)" -ForegroundColor DarkCyan
                try {
                    $items = Get-PnPListItem -List $lib.Title -PageSize 500 -Fields "FileLeafRef","FileRef","Modified","Editor","Author","FSObjType","SMTotalFileStreamSize"
                } catch {
                    Write-Warn2 "   Skipping '$($lib.Title)' (unsupported): $($_.Exception.Message)"
                    continue
                }

                if ($ReportMode -eq 'Duplicates') {
                    $bucket = @{}
                }

                foreach ($it in $items) {
                    Stop-IfCancelled
                    if ($it['FSObjType'] -ne 0) { continue }

                    $name     = $it['FileLeafRef']
                    $size     = [int64]($it.FieldValues['SMTotalFileStreamSize'])
                    if (-not $size) { continue }
                    if ($ReportMode -eq 'AllFiles' -and $size -lt $sizeThreshold) { continue }
                    if ($ReportMode -eq 'Duplicates' -and $size -lt $sizeThreshold) { continue }

                    $row = [pscustomobject]@{
                        SiteURL           = $s.Url
                        Library           = $lib.Title
                        FileName          = $name
                        'File Size (MB)'  = [math]::Round($size/1MB,2)
                        LastModified      = $it['Modified']
                        CreatedBy         = $it['Author'].LookupValue
                        ModifiedBy        = $it['Editor'].LookupValue
                        FolderLocation    = $it['FileRef']
                    }

                    if ($ReportMode -eq 'Duplicates') {
                        $key = "$name|$size"
                        if (-not $bucket.ContainsKey($key)) { $bucket[$key] = @() }
                        $bucket[$key] += $row
                    } else {
                        $results.Add($row) | Out-Null
                    }
                }

                if ($ReportMode -eq 'Duplicates') {
                    foreach ($kv in $bucket.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }) {
                        foreach ($r in $kv.Value) {
                            $results.Add([pscustomobject]@{
                                'Duplicate File'  = $r.FileName
                                'Occurrences'     = $kv.Value.Count
                                'File Size (MB)'  = $r.'File Size (MB)'
                                'Site URL'        = $r.SiteURL
                                'Library'         = $r.Library
                                'Last Modified'   = $r.LastModified
                                'Modified By'     = $r.ModifiedBy
                                'Created By'      = $r.CreatedBy
                                'Folder Location' = $r.FolderLocation
                            }) | Out-Null
                        }
                    }
                }
            }
        }

        try {
            Invoke-OnSite -SiteUrl $s.Url -AdminUpn $AdminUpn -ScriptBlock $scriptBlock
        } catch {
            Write-Bad "   Error on $($s.Url): $($_.Exception.Message)"
            $errors.Add([pscustomobject]@{
                SiteURL     = $s.Url
                ErrorMessage= $_.Exception.Message
            }) | Out-Null
        }
    }

    Export-ReportCsv -Data $results -Path $OutPath
    Export-ReportCsv -Data $errors  -Path $ErrorPath
}

function New-AccessReviewReport {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$OutPath,
        [string]$AdminUpn,
        [string]$SiteUrl,
        [string]$FilterUser # partial match on login/title/email
    )

    $rows  = [System.Collections.Generic.List[object]]::new()
    $sites = if ($SiteUrl) { @(@{ Url = $SiteUrl }) } else { Get-AllSites }
    $needle = $FilterUser?.ToLower().Trim()

    $i=0
    foreach ($s in $sites) {
        Stop-IfCancelled
        $i++
        Write-Info "($i/$($sites.Count)) $($s.Url)"

        $scriptBlock = {
            $groups = Get-PnPGroup
            foreach ($g in $groups) {
                Stop-IfCancelled
                $members = Get-PnPGroupMember -Group $g
                foreach ($m in $members) {
                    $login = ($m.LoginName ?? '').ToLower().Trim()
                    $title = ($m.Title ?? '').ToLower().Trim()
                    $email = ($m.Email ?? '').ToLower().Trim()

                    $match = $true
                    if ($needle) {
                        $match = ($login -like "*$needle*") -or ($title -like "*$needle*") -or ($email -like "*$needle*")
                    }
                    if ($match) {
                        $rows.Add([pscustomobject]@{
                            SiteUrl   = $s.Url
                            GroupName = $g.Title
                            UserName  = $m.Title
                            UserLogin = $m.LoginName
                            UserEmail = $m.Email
                        }) | Out-Null
                    }
                }
            }
        }

        try {
            Invoke-OnSite -SiteUrl $s.Url -AdminUpn $AdminUpn -ScriptBlock $scriptBlock
        } catch {
            Write-Warn2 "Skip $($s.Url): $($_.Exception.Message)"
        }
    }

    if ($rows.Count -gt 0) {
        Export-ReportCsv -Data $rows -Path $OutPath
    } else {
        Write-Warn2 "No matching access entries found."
    }
}

function New-ExternalUsersReport {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$OutPath,
        [switch]$IncludeOneDrive
    )
    $rows  = [System.Collections.Generic.List[object]]::new()
    $sites = Get-AllSites -IncludeOneDrive:$IncludeOneDrive
    $i=0
    foreach ($s in $sites) {
        Stop-IfCancelled
        $i++
        Write-Info "($i/$($sites.Count)) $($s.Url)"
        try {
            Connect-SPO -Url $s.Url
            $ext = Get-PnPExternalUser -PageSize 50 -SiteUrl $s.Url
            foreach ($u in $ext) {
                $rows.Add([pscustomobject]@{
                    SiteUrl     = $s.Url
                    DisplayName = $u.DisplayName
                    Email       = $u.Email
                    AcceptedAs  = $u.AcceptedAs
                    WhenCreated = $u.WhenCreated
                    InvitedBy   = $u.InvitedBy
                    InvitedAs   = $u.InvitedAs
                }) | Out-Null
            }
        } catch {
            Write-Warn2 "Failed on $($s.Url): $($_.Exception.Message)"
        }
    }
    Export-ReportCsv -Data $rows -Path $OutPath
}
#endregion

#region ========= Menu / Entrypoints =========
function Start-SPOReportTool {
    [CmdletBinding()]
    param()

    Ensure-Modules

    Write-Host ""
    Write-Info  "SPO Reports — Refactored"
    Write-Host "Press Q at any time to cancel." -ForegroundColor DarkCyan

    if (-not $script:AdminUrl)  { $script:AdminUrl  = Read-Host "Admin URL (e.g. https://contoso-admin.sharepoint.com)" }
    if (-not $script:ClientId)  { $script:ClientId  = Read-Host "Client ID (Entra app)" }
    $script:AuthMode = 'Interactive' # default UX
    Write-Good "Using Interactive auth. (Switch to AppCert via Set-SPOConnectionSettings if needed.)"

    $menu = @(
        'Inactive Sites',
        'Sharing Settings',
        'Duplicate Files',
        'Large Files',
        'User Access Review',
        'External Users',
        'Update Connection Settings',
        'Exit'
    )

    while ($true) {
        Write-Host ""
        for ($i=0;$i -lt $menu.Count;$i++){ Write-Host "[$i] $($menu[$i])" }
        $sel = Read-Host "Choose (0-$($menu.Count-1))"
        Stop-IfCancelled

        switch ($sel) {
            '0' {
                $path = Resolve-PathForExport -OutDir $null -DefaultFileName "Inactive_Sites_$(Get-Timestamp).csv"
                if (-not $path) { continue }
                $upn  = Read-Host "Temporarily add admin UPN where unauthorized? (blank to skip)"
                try {
                    New-InactiveSitesReport -OutPath $path -AdminUpn $upn
                } finally {
                    if ($upn) { Remove-TempOwners -AdminUpn $upn }
                }
            }
            '1' {
                $path = Resolve-PathForExport -OutDir $null -DefaultFileName "Sharing_Settings_$(Get-Timestamp).csv"
                if (-not $path) { continue }
                New-SharingSettingsReport -OutPath $path
            }
            '2' {
                $path  = Resolve-PathForExport -OutDir $null -DefaultFileName "Duplicate_Files_$(Get-Timestamp).csv"
                $epath = Resolve-PathForExport -OutDir $null -DefaultFileName "Errored_Sites_$(Get-Timestamp).csv"
                if (-not $path -or -not $epath) { continue }
                $minMB = [int](Read-Host "Min file size MB for duplicates (default 100)"); if ($minMB -le 0){$minMB=100}
                $libs  = (Read-Host "Search ALL libraries? (Y/N)").ToUpper() -eq 'Y'
                $upn   = Read-Host "Temporarily add admin UPN where unauthorized? (blank to skip)"
                try {
                    New-FileScanReport -ReportMode Duplicates -MinFileSizeMB $minMB -OutPath $path -ErrorPath $epath -AdminUpn $upn -AllLibraries:$libs
                } finally {
                    if ($upn) { Remove-TempOwners -AdminUpn $upn }
                }
            }
            '3' {
                $path  = Resolve-PathForExport -OutDir $null -DefaultFileName "Large_Files_$(Get-Timestamp).csv"
                $epath = Resolve-PathForExport -OutDir $null -DefaultFileName "Errored_Sites_$(Get-Timestamp).csv"
                if (-not $path -or -not $epath) { continue }
                $minMB = [int](Read-Host "Min file size MB (default 100)"); if ($minMB -le 0){$minMB=100}
                $libs  = (Read-Host "Search ALL libraries? (Y/N)").ToUpper() -eq 'Y'
                $upn   = Read-Host "Temporarily add admin UPN where unauthorized? (blank to skip)"
                try {
                    New-FileScanReport -ReportMode AllFiles -MinFileSizeMB $minMB -OutPath $path -ErrorPath $epath -AdminUpn $upn -AllLibraries:$libs
                } finally {
                    if ($upn) { Remove-TempOwners -AdminUpn $upn }
                }
            }
            '4' {
                $path = Resolve-PathForExport -OutDir $null -DefaultFileName "Access_Review_$(Get-Timestamp).csv"
                if (-not $path) { continue }
                $single = (Read-Host "Single site only? (Y/N)").ToUpper() -eq 'Y'
                $site   = if ($single) { Read-Host "Site URL (full)" } else { $null }
                $filter = (Read-Host "Filter for specific user (email/login/name, blank for all)")
                $upn    = Read-Host "Temporarily add admin UPN where unauthorized? (blank to skip)"
                try {
                    New-AccessReviewReport -OutPath $path -AdminUpn $upn -SiteUrl $site -FilterUser $filter
                } finally {
                    if ($upn) { Remove-TempOwners -AdminUpn $upn }
                }
            }
            '5' {
                $path = Resolve-PathForExport -OutDir $null -DefaultFileName "External_Users_$(Get-Timestamp).csv"
                if (-not $path) { continue }
                $od   = (Read-Host "Include OneDrive sites? (Y/N)").ToUpper() -eq 'Y'
                New-ExternalUsersReport -OutPath $path -IncludeOneDrive:$od
            }
            '6' {
                $new = Read-Host "New Admin URL"
                $cid = Read-Host "New Client ID"
                Set-SPOConnectionSettings -AdminUrl $new -ClientId $cid -AuthMode 'Interactive'
            }
            '7' { return }
            default { Write-Bad "Invalid selection." }
        }
    }
}

#endregion

#region ========= Automation-friendly wrappers =========
function Invoke-SPOInactiveSites {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [string]$AdminUpn,
        [string]$OutDir = "."
    )
    Ensure-Modules
    Set-SPOConnectionSettings -AdminUrl $AdminUrl -ClientId $ClientId -AuthMode 'Interactive'
    $csv = Resolve-PathForExport -OutDir $OutDir -DefaultFileName "Inactive_Sites_$(Get-Timestamp).csv"
    try {
        New-InactiveSitesReport -OutPath $csv -AdminUpn $AdminUpn
    } finally {
        if ($AdminUpn) { Remove-TempOwners -AdminUpn $AdminUpn }
    }
}

function Invoke-SPOSharingSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [string]$OutDir = "."
    )
    Ensure-Modules
    Set-SPOConnectionSettings -AdminUrl $AdminUrl -ClientId $ClientId -AuthMode 'Interactive'
    $csv = Resolve-PathForExport -OutDir $OutDir -DefaultFileName "Sharing_Settings_$(Get-Timestamp).csv"
    New-SharingSettingsReport -OutPath $csv
}

function Invoke-SPODuplicates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [int]$MinFileSizeMB = 100,
        [string]$AdminUpn,
        [string]$OutDir = ".",
        [switch]$AllLibraries
    )
    Ensure-Modules
    Set-SPOConnectionSettings -AdminUrl $AdminUrl -ClientId $ClientId -AuthMode 'Interactive'
    $csv   = Resolve-PathForExport -OutDir $OutDir -DefaultFileName "Duplicate_Files_$(Get-Timestamp).csv"
    $ecsv  = Resolve-PathForExport -OutDir $OutDir -DefaultFileName "Errored_Sites_$(Get-Timestamp).csv"
    try {
        New-FileScanReport -ReportMode Duplicates -MinFileSizeMB $MinFileSizeMB -OutPath $csv -ErrorPath $ecsv -AdminUpn $AdminUpn -AllLibraries:$AllLibraries
    } finally {
        if ($AdminUpn) { Remove-TempOwners -AdminUpn $AdminUpn }
    }
}

function Invoke-SPOLargeFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [int]$MinFileSizeMB = 100,
        [string]$AdminUpn,
        [string]$OutDir = ".",
        [switch]$AllLibraries
    )
    Ensure-Modules
    Set-SPOConnectionSettings -AdminUrl $AdminUrl -ClientId $ClientId -AuthMode 'Interactive'
    $csv   = Resolve-PathForExport -OutDir $OutDir -DefaultFileName "Large_Files_$(Get-Timestamp).csv"
    $ecsv  = Resolve-PathForExport -OutDir $OutDir -DefaultFileName "Errored_Sites_$(Get-Timestamp).csv"
    try {
        New-FileScanReport -ReportMode AllFiles -MinFileSizeMB $MinFileSizeMB -OutPath $csv -ErrorPath $ecsv -AdminUpn $AdminUpn -AllLibraries:$AllLibraries
    } finally {
        if ($AdminUpn) { Remove-TempOwners -AdminUpn $AdminUpn }
    }
}

function Invoke-SPOAccessReview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [string]$SiteUrl,
        [string]$FilterUser,
        [string]$AdminUpn,
        [string]$OutDir = "."
    )
    Ensure-Modules
    Set-SPOConnectionSettings -AdminUrl $AdminUrl -ClientId $ClientId -AuthMode 'Interactive'
    $csv = Resolve-PathForExport -OutDir $OutDir -DefaultFileName "Access_Review_$(Get-Timestamp).csv"
    try {
        New-AccessReviewReport -OutPath $csv -AdminUpn $AdminUpn -SiteUrl $SiteUrl -FilterUser $FilterUser
    } finally {
        if ($AdminUpn) { Remove-TempOwners -AdminUpn $AdminUpn }
    }
}

function Invoke-SPOExternalUsers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AdminUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [string]$OutDir = ".",
        [switch]$IncludeOneDrive
    )
    Ensure-Modules
    Set-SPOConnectionSettings -AdminUrl $AdminUrl -ClientId $ClientId -AuthMode 'Interactive'
    $csv = Resolve-PathForExport -OutDir $OutDir -DefaultFileName "External_Users_$(Get-Timestamp).csv"
    New-ExternalUsersReport -OutPath $csv -IncludeOneDrive:$IncludeOneDrive
}
#endregion

# You can call Start-SPOReportTool for the menu UX:
# Start-SPOReportTool
