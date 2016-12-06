$currentWorkingDirectory = (Get-Item -Path ".\" -Verbose).FullName
[void][Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO")
[void][Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Management.RegisteredServers")
##Using local Invoke-SqlCmd2 until Invoke-SqlCmd2 powershell gallery module is fixed //Jonas 20161122
. .\Invoke-SqlCmd2.ps1
<#
if ((Get-Module -ListAvailable -Name "Invoke-Sqlcmd2") -eq $null) {
    Install-Module -Scope CurrentUser -Name "Invoke-Sqlcmd2" -Force
} else {
    Import-Module -Name "Invoke-Sqlcmd2"
}
#>
 
if ((Test-Path -Path "config.json") -eq $false -or $Init -eq $true) {
    Get-Content -Path ".\config.jsontemplate.json" | Out-File -FilePath "config.json"
}
$ScriptConfiguration = (Get-Content "config.json") -Join "`n" | ConvertFrom-Json
if ($ScriptConfiguration.CMS -ne $null) {
    $cmserver = $ScriptConfiguration.CMS
} else {
    Write-Error "CMS missing in configuration file"
    exit
} 
$server = New-Object Microsoft.SqlServer.Management.Smo.Server $cmserver
$sqlconnection = $server.ConnectionContext.SqlConnectionObject
$InventoryFiles = Get-ChildItem -Filter "*.csv" -Path $ScriptConfiguration.InventoryFilePath
try { $cmstore = new-object Microsoft.SqlServer.Management.RegisteredServers.RegisteredServersStore($sqlconnection)}
catch { throw "Cannot access Central Management Server" }
$server.ConnectionContext.Disconnect()
$dbstore = $cmstore.DatabaseEngineServerGroup
$registeredServers = $dbstore.GetDescendantRegisteredServers()
if ($ScriptConfiguration.UncategorizedServerGroup -ne $null) {
    $uncategorizedServerGroupName = $ScriptConfiguration.UncategorizedServerGroup
} else {
    Write-Error "UncategorizedServerGroup missing in configuration file"
    exit
}
$groupstore = $dbstore
$groupobject = $groupstore.ServerGroups[$uncategorizedServerGroupName]
if ($groupobject -eq $null) {
    Write-Warning "Creating group $uncategorizedServerGroupName"
    $newgroup = New-Object Microsoft.SqlServer.Management.RegisteredServers.ServerGroup($groupstore, $uncategorizedServerGroupName)
    $newgroup.create()
    $groupstore.refresh()
}
$groupstore = $groupstore.ServerGroups[$uncategorizedServerGroupName]
$ErrorLogPath = $ScriptConfiguration.ErrorLogPath
if ($ErrorLogPath -eq $null) {
    Write-Error "Missing error log path in configuration file."
    exit
}
foreach ($InventoryFile in $InventoryFiles) {
    $SQLServerInstances = Import-Csv $inventoryFile.FullName
    foreach ($foundServer in $SQLServerInstances) {
        if ($foundServer.Status -ne "Running") {
            continue
        }
        $displayName = $foundServer.DisplayName
        $machineName = $foundServer.MachineName
        $SQLInstance = ""
        $instanceName = ""
        if ($displayName.Contains("$")) {
            $instanceName = $displayName.Split("$")[1]
        } elseif (($displayName.StartsWith("SQL Server (")) -and ($displayName.EndsWith(")"))) {
            $instanceName = $displayName.Replace("SQL Server (", "")
            $instanceName = $instanceName.Replace(")", "")
        }
        if (($instanceName -ne "") -and ($instanceName -ne "MSSQLSERVER")) {
            $SQLInstance = "$machineName\$instanceName"
        } else {
            $SQLInstance = "$machineName"
        }
        if ($SQInstance -eq $cmserver) {
            continue
        }
        $res = $null
        if ($SQLInstance -ne "" -and $SQLInstance -ne $cmserver) {
            try {
                $recentError = $null
                $res = Invoke-Sqlcmd2 -ServerInstance $SQLInstance -ConnectionTimeout 2 -Query "SELECT 1 AS [Success]" -ErrorVariable recentError -ErrorAction SilentlyContinue
                if ($recentError -ne $null) {
                    $currentDate = Get-Date
                    Out-File -FilePath $ErrorLogPath -Append -InputObject $currentDate
                    Out-File -FilePath $ErrorLogPath -Append -InputObject $SQLInstance
                    Out-File -FilePath $ErrorLogPath -Append -InputObject $recentError
                }
                if ($res.Success -eq "1") {
                    if ($RegisteredServers.ServerName -notcontains $SQLInstance) {
                        $newserver = New-Object Microsoft.SqlServer.Management.RegisteredServers.RegisteredServer($groupstore, $SQLInstance)
                        $newserver.ServerName = $SQLInstance
                        $addedDate = Get-Date
                        $newserver.Description = "Added: $addedDate"
                        $newserver.Create()
                        Write-Host "Added Server $SQLInstance" -ForegroundColor Green
                        $groupstore.Refresh()
                    } else {
                    Write-Debug "Server $SQLInstance already exists. Skipped"
                    }
                }
            }catch {
            Write-Warning $_
            Out-File -FilePath $ErrorLogPath -Append -InputObject $_
            }
        }
    }
}