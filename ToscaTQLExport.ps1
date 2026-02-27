# =============================================================================
# Run-ToscaTQL.ps1
# Tosca Execution History Export — via TCAPI (no UI required)
# Requires: PowerShell 7+, Tosca Commander installed locally
# =============================================================================

Write-Host "Starting Tosca execution history export..."

# ---------------- CONFIG ----------------
$toscaCommanderPath  = "C:\Program Files (x86)\TRICENTIS\Tosca Testsuite\ToscaCommander"
$toscaServerUrl      = "http://localhost/"
$workspacePath       = "C:\Tosca_Projects\Tosca_Workspaces\ToscaCork\ToscaCork.tws"
$outputFile          = "C:\Temp\TQL_ExecutionResults.xml"

# AUTH: fill in ONE of the following
$personalAccessToken = "ew0KICAiQ2xpZW50SWQiOiAiNE04MkQ3YTRIRXl0SEpseDd1MXJSdyIsDQogICJEaXNwbGF5TmFtZSI6ICJUQ0FQSVRlc3QiLA0KICAiQ2xpZW50U2VjcmV0IjogIkNQempTUDBESms2ckg0RUFJNzhQd3dMOW5JekpwbnFrR2RHQ2VNZ2QxeWl3IiwNCiAgIlNjb3BlcyI6IFsNCiAgICAiQXV0aGVudGljYXRpb25TZXJ2aWNlQXBpIiwNCiAgICAiRGV4QXBpIiwNCiAgICAiRmlsZVNlcnZpY2VBcGkiLA0KICAgICJQcm9qZWN0U2VydmljZUFwaSIsDQogICAgIlRlc3REYXRhU2VydmljZUFwaSIsDQogICAgIlRvc2NhQXV0b21hdGlvbk9iamVjdFNlcnZpY2VBcGkiLA0KICAgICJBZG1pbkNvbnNvbGVBcGkiLA0KICAgICJNaWdyYXRpb25TZXJ2aWNlQXBpIiwNCiAgICAiRXhlY3V0aW9uUmVzdWx0U2VydmljZUFwaSIsDQogICAgIlJwYUFwaUdhdGV3YXlBcGkiLA0KICAgICJUZXN0RGVzaWduU3R1ZGlvQXBpIiwNCiAgICAiTmV4dXNBcGkiLA0KICAgICJPc3ZTdHVkaW9BcGkiLA0KICAgICJMaXZlQ29tcGFyZVNlcnZpY2VBcGkiLA0KICAgICJEYXRhSW50ZWdyaXR5U2VydmljZUFwaSIsDQogICAgIk5vdGlmaWNhdGlvblNlcnZpY2VBcGkiLA0KICAgICJFeGFtcGxlU2VydmljZUFwaSIsDQogICAgIlNhcEludGVncmF0aW9uQXBpIiwNCiAgICAiTWJ0U2VydmljZUFwaSIsDQogICAgIkh1YlNlcnZpY2VBcGkiLA0KICAgICJHYXRld2F5QXBpIiwNCiAgICAiTGljZW5zZUFkbWluaXN0cmF0aW9uQXBpIiwNCiAgICAidmlzaW9uYWkubmV4dXMuYXBpIiwNCiAgICAidmlzaW9uYWkuYWdlbnQuYXBpIiwNCiAgICAiYXV0aGVudGljYXRpb24uZ3JvdXBzLnJlYWQiDQogIF0NCn0="   # Personal Access Token  (Tosca Server > Profile > Personal Access Tokens)
$clientId            = ""   # OR: OAuth Client ID    (Tosca Server > Administration > OAuth Clients)
$clientSecret        = ""   #     OAuth Client Secret

# Tosca workspace credentials (same as Tosca Commander login)
$workspaceUser       = "Admin"
$workspacePassword   = "1234567890-="

# -----------------------------------------------------------------------
# TQL QUERIES — uncomment the one you want, or define your own
# -----------------------------------------------------------------------

# All execution log entries (full history)
$tql = "=>SUBPARTS:ExecutionLogEntry"

# Only FAILED results
# $tql = "=>SUBPARTS:ExecutionLogEntry[ExecutionStatus=Failed]"

# Only PASSED results
# $tql = "=>SUBPARTS:ExecutionLogEntry[ExecutionStatus=Passed]"

# Results within a date range (ISO format)
# $tql = "=>SUBPARTS:ExecutionLogEntry[StartedAt>='2025-01-01' AND StartedAt<='2025-12-31']"

# Execution results for a specific test case by name
# $tql = "=>SUBPARTS:ExecutionLogEntry[TestCaseName='My Test Case']"

# Execution lists (top-level containers)
# $tql = "=>SUBPARTS:ExecutionList"

# ----------------------------------------

New-Item -ItemType Directory -Path (Split-Path $outputFile) -Force | Out-Null
$env:PATH = "$toscaCommanderPath;$env:PATH"

# ================================
# AssemblyResolve handler
# ================================
$resolverPath = $toscaCommanderPath
[System.AppDomain]::CurrentDomain.add_AssemblyResolve([System.ResolveEventHandler]{
    param($sender, $args)
    $assemblyName = [System.Reflection.AssemblyName]::new($args.Name).Name
    $candidate = Join-Path $resolverPath "$assemblyName.dll"
    if (Test-Path $candidate) { return [System.Reflection.Assembly]::LoadFrom($candidate) }
    return $null
})

# ================================
# Load assemblies
# ================================
try {
    Write-Host "Loading assemblies..."
    $assemblyObjects = [System.Reflection.Assembly]::LoadFrom((Join-Path $toscaCommanderPath "TCAPIObjects.dll"))
    $assemblyTCAPI   = [System.Reflection.Assembly]::LoadFrom((Join-Path $toscaCommanderPath "TCAPI.dll"))
    $tcapiType       = $assemblyTCAPI.GetType("Tricentis.TCAPI.TCAPI")
    Write-Host "Assemblies loaded OK"
}
catch { Write-Host "FAILED to load assemblies: $_"; exit 1 }

# ================================
# Get or create TCAPI instance
# ================================
try {
    Write-Host "Initialising TCAPI..."
    $tcapi = $tcapiType.GetProperty("Instance").GetValue($null)

    if (-not $tcapi) {
        $connInfoType = $assemblyObjects.GetType("Tricentis.TCAPIObjects.TCAPIConnectionInfo")
        $connInfo     = [Activator]::CreateInstance($connInfoType)
        $connInfo.Url = $toscaServerUrl

        $createMethod = $tcapiType.GetMethods([System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::Public) |
            Where-Object { $_.Name -eq "CreateInstance" -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq "TCAPIConnectionInfo" } |
            Select-Object -First 1
        $createMethod.Invoke($null, @($connInfo))
        $tcapi = $tcapiType.GetProperty("Instance").GetValue($null)
    }

    if (-not $tcapi) { Write-Host "ERROR: TCAPI instance is null"; exit 1 }
    Write-Host "TCAPI ready (v$($tcapi.APIVersionString))"
}
catch { Write-Host "Failed to initialise TCAPI: $_"; if ($_.Exception.InnerException) { Write-Host "Inner: $($_.Exception.InnerException.Message)" }; exit 1 }

# ================================
# Login
# ================================
try {
    Write-Host "Logging in..."
    $loginMethods = $tcapi.GetType().GetMethods() | Where-Object { $_.Name -eq "ToscaServerLogin" }

    if ($personalAccessToken -ne "") {
        $loginMethod = $loginMethods | Where-Object { $_.GetParameters().Count -eq 1 } | Select-Object -First 1
        $loginMethod.Invoke($tcapi, @($personalAccessToken)) | Out-Null
    } elseif ($clientId -ne "" -and $clientSecret -ne "") {
        $loginMethod = $loginMethods | Where-Object { $_.GetParameters().Count -eq 2 } | Select-Object -First 1
        $loginMethod.Invoke($tcapi, @($clientId, $clientSecret)) | Out-Null
    } else {
        Write-Host "ERROR: No auth credentials provided in CONFIG."; exit 1
    }
    Write-Host "Login OK"
}
catch { Write-Host "Login failed: $_"; if ($_.Exception.InnerException) { Write-Host "Inner: $($_.Exception.InnerException.Message)" }; exit 1 }

# ================================
# Open workspace
# ================================
try {
    Write-Host "Opening workspace..."
    $openMethods = $tcapi.GetType().GetMethods() | Where-Object { $_.Name -eq "OpenWorkspace" }
    $openMethod  = $openMethods | Where-Object { $_.GetParameters().Count -eq 4 } | Select-Object -First 1
    if (-not $openMethod) { $openMethod = $openMethods | Sort-Object { $_.GetParameters().Count } | Select-Object -First 1 }
    $openMethod.Invoke($tcapi, @($workspacePath, $workspaceUser, $workspacePassword, 0))
    Write-Host "Workspace opened OK"
}
catch { Write-Host "OpenWorkspace failed: $_"; if ($_.Exception.InnerException) { Write-Host "Inner: $($_.Exception.InnerException.Message)" }; exit 1 }

$workspace = $tcapi.GetType().GetProperty("ActiveWorkspace").GetValue($tcapi)
if (-not $workspace) { Write-Host "ERROR: ActiveWorkspace is null"; exit 1 }

# ================================
# Get project root and run TQL
# ================================
try {
    Write-Host "Running TQL: $tql"
    $project = $workspace.GetType().GetMethod("GetProject").Invoke($workspace, @())
    if (-not $project) { Write-Host "ERROR: GetProject() returned null"; exit 1 }

    $searchMethod = $workspace.GetType().GetMethod("Search")
    $results      = $searchMethod.Invoke($workspace, @($project, $tql))
    Write-Host "Found $($results.Count) result(s)"
}
catch { Write-Host "TQL failed: $_"; if ($_.Exception.InnerException) { Write-Host "Inner: $($_.Exception.InnerException.Message)" }; exit 1 }

# ================================
# Helper: safely read a property
# ================================
function Get-Prop($obj, $propName) {
    try {
        $p = $obj.GetType().GetProperty($propName)
        if ($p) { $v = $p.GetValue($obj); if ($v -ne $null) { return $v.ToString() } }
    } catch {}
    return ""
}

# ================================
# Build XML
# Columns are auto-detected from the first result object,
# then applied to all results — so the output always reflects
# what is actually available in your Tosca version.
# ================================
Write-Host "Building XML..."

$xml     = New-Object System.Xml.XmlDocument
$xmlRoot = $xml.CreateElement("Results")
$xml.AppendChild($xmlRoot) | Out-Null

# Preferred column order for execution log entries
# Any that don't exist on the object are silently skipped
$preferredColumns = @(
    "Name",
    "NodePath",
    "UniqueId",
    "ExecutionStatus",
    "TestCaseName",
    "TestCaseUniqueId",
    "StartedAt",
    "EndedAt",
    "Duration",
    "ExecutionEnvironment",
    "ExecutionListName",
    "ExecutionListUniqueId",
    "CreatedBy",
    "CreatedAt",
    "ModifiedBy",
    "ModifiedAt",
    "Description",
    "Revision"
)

# Detect actual available properties from first result
$availableProps = @()
if ($results.Count -gt 0) {
    $availableProps = $results[0].GetType().GetProperties() | Select-Object -ExpandProperty Name

    # Build final column list: preferred columns first (if they exist), then any remaining
    $columns = @()
    foreach ($col in $preferredColumns) {
        if ($availableProps -contains $col) { $columns += $col }
    }
    foreach ($prop in $availableProps) {
        if ($columns -notcontains $prop) { $columns += $prop }
    }

    Write-Host "Columns in output: $($columns -join ', ')"
} else {
    Write-Host "No results found — empty XML will be written."
    $columns = @()
}

$rowType = if ($results.Count -gt 0) { $results[0].GetType().Name } else { "Result" }

foreach ($item in $results) {
    $node = $xml.CreateElement($rowType)

    foreach ($col in $columns) {
        $val = Get-Prop $item $col
        # Skip complex object references (they stringify as type names)
        if ($val -match "^Tricentis\." -or $val -match "^System\.") { continue }
        $el = $xml.CreateElement($col)
        $el.InnerText = $val
        $node.AppendChild($el) | Out-Null
    }

    $xmlRoot.AppendChild($node) | Out-Null
}

# ================================
# Save XML
# ================================
try {
    $xml.Save($outputFile)
    Write-Host "Export complete: $outputFile"
    Write-Host "Total records: $($results.Count)"
}
catch { Write-Host "Failed to save XML: $_"; exit 1 }

# ================================
# Clean up
# ================================
try {
    $tcapi.GetType().GetMethods() | Where-Object { $_.Name -eq "CloseWorkspace" -and $_.GetParameters().Count -eq 0 } | Select-Object -First 1 | ForEach-Object { $_.Invoke($tcapi, @()) }
    $tcapiType.GetMethods([System.Reflection.BindingFlags]::Static -bor [System.Reflection.BindingFlags]::Public) | Where-Object { $_.Name -eq "CloseInstance" } | Select-Object -First 1 | ForEach-Object { $_.Invoke($null, @()) }
    Write-Host "Closed cleanly"
} catch {}

Write-Host "Done."