# Tosca-TQL-Export-
Connects to a Tosca workspace via the TCAPI (Tosca Commander API) and runs a TQL query to export data to XML
# Tosca TQL Export — README

## What this does

`Run-ToscaTQLExport.ps1` connects to a Tosca workspace via the **TCAPI** (Tosca Commander API) and runs a TQL query to export data to XML — with no UI required. It is designed for bulk export of historical test execution data but can query any Tosca object type.

---

## Requirements

- **PowerShell 7+** (not Windows PowerShell 5 — the Tosca DLLs require .NET 8)
- **Tosca Commander installed locally** at the path in `$toscaCommanderPath`
- Access to your Tosca Server (for authentication) and workspace `.tws` file

---

## ⚙️ What you need to change before running

Open `Run-ToscaTQLExport.ps1`. The CONFIG section is at the top of the file (lines 6–28). You **must** update the following values before the script will work:

---

### 1. Tosca Commander install path
```powershell
$toscaCommanderPath = "C:\Program Files (x86)\TRICENTIS\Tosca Testsuite\ToscaCommander"
```
**Change this if** Tosca Commander is installed in a different location on your machine.
To find the correct path, right-click the Tosca Commander shortcut > Properties > look at the "Start in" or "Target" field.

---

### 2. Tosca Server URL
```powershell
$toscaServerUrl = "http://localhost/"
```
**Change this to** your Tosca Server address. If Tosca Server is on a different machine, it will be something like `http://tosca-server.yourcompany.com/`.
Leave as `http://localhost/` if Tosca Server is running on the same machine.

---

### 3. Workspace file path
```powershell
$workspacePath = "C:\Tosca_Projects\Tosca_Workspaces\ToscaCork\ToscaCork.tws"
```
**Change this to** the full path of your `.tws` workspace file.
You can find this by opening Tosca Commander and checking the workspace path in the title bar or File > Recent Workspaces.

---

### 4. Output file location
```powershell
$outputFile = "C:\Temp\TQL_ExecutionResults.xml"
```
**Change this to** wherever you want the XML file saved. The folder will be created automatically if it doesn't exist.

---

### 5. Authentication — choose ONE method

**Option A: Personal Access Token (recommended)**
```powershell
$personalAccessToken = ""   # <-- paste your token here
```
To get a PAT: log in to Tosca Server in your browser > click your profile/avatar > **Personal Access Tokens** > generate a new token > paste the token string here.

**Option B: OAuth Client ID + Secret**
```powershell
$clientId     = ""   # <-- paste your Client ID here
$clientSecret = ""   # <-- paste your Client Secret here
```
To get these: Tosca Server > **Administration** > **OAuth Clients** > create a new client.

> Leave the unused option blank. If `$personalAccessToken` is filled in, it takes priority.

---

### 6. Workspace login credentials
```powershell
$workspaceUser     = "Admin"
$workspacePassword = "1234567890-="
```
**Change these to** the username and password you normally use when opening this workspace in Tosca Commander.
If your workspace has no password set, leave `$workspacePassword` as `""`.

---

### 7. TQL query (optional — default works out of the box)
```powershell
$tql = "=>SUBPARTS:ExecutionLogEntry"
```
The default query retrieves **all execution history**. You can leave this as-is or customise it. See the **TQL queries** section below for examples.

---

## Running the script

Open a **fresh PowerShell 7** window (search "PowerShell 7" in Start menu — not "Windows PowerShell") and run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
cd "C:\Path\To\Script\Folder"
.\Run-ToscaTQL.ps1
```

To run silently and save output to a log file:

```powershell
.\Run-ToscaTQL.ps1 > C:\Temp\export_log.txt 2>&1
```

To schedule it (e.g. nightly via Task Scheduler):

```
Program:   pwsh.exe
Arguments: -ExecutionPolicy Bypass -File "C:\Path\To\Run-ToscaTQL.ps1"
```

> **Important:** Always use a fresh PowerShell 7 window. If you run the script twice in the same window without it closing cleanly, you may get an "Api already initialized" error.

---

## TQL queries for execution history

Several queries are pre-written in the script (commented out). To use one, delete the `#` at the start of that line and add `#` to the current active `$tql` line.

| Query | Description |
|---|---|
| `=>SUBPARTS:ExecutionLogEntry` | **All execution history** (default) |
| `=>SUBPARTS:ExecutionLogEntry[ExecutionStatus=Failed]` | Failed runs only |
| `=>SUBPARTS:ExecutionLogEntry[ExecutionStatus=Passed]` | Passed runs only |
| `=>SUBPARTS:ExecutionLogEntry[StartedAt>='2025-01-01']` | Results after a specific date |
| `=>SUBPARTS:ExecutionLogEntry[StartedAt>='2025-01-01' AND StartedAt<='2025-12-31']` | Results within a date range |
| `=>SUBPARTS:ExecutionLogEntry[TestCaseName='My Test']` | Results for a specific test case |
| `=>SUBPARTS:ExecutionList` | Top-level execution list containers |
| `=>SUBPARTS:TestCase` | All test cases (not execution results) |

TQL syntax reference: [Tricentis documentation](https://documentation.tricentis.com)

---

<img width="1901" height="1200" alt="image" src="https://github.com/user-attachments/assets/54392a72-4049-42a8-92b2-c02eb99301fc" />


## Output format

The script writes an XML file at `$outputFile`. Each result is a child element of `<Results>`:

```xml
<Results>
  <ExecutionLogEntry>
    <Name>My Test Run</Name>
    <NodePath>/ExecutionLists/Sprint 42/My Test Run</NodePath>
    <UniqueId>01KF3FGGNNCC98DADNTGARBQAB</UniqueId>
    <ExecutionStatus>Passed</ExecutionStatus>
    <TestCaseName>Login Test</TestCaseName>
    <StartedAt>1/16/2026 1:21:58 PM</StartedAt>
    <EndedAt>1/16/2026 1:22:05 PM</EndedAt>
    <Duration>7</Duration>
    <CreatedBy>Admin</CreatedBy>
    ...
  </ExecutionLogEntry>
</Results>
```

---

## Large data volumes

For very large exports (tens of thousands of records), consider:

- **Filtering by date range** in the TQL query to break the export into smaller batches
- Running the script **on the Tosca Server machine itself** to avoid network overhead
- Running with: `pwsh.exe -NoProfile -ExecutionPolicy Bypass -File .\Run-ToscaTQL.ps1`

---

## Troubleshooting

| Error | Fix |
|---|---|
| `Api already initialized` | Open a **fresh PS7 window** before running |
| `Login Failed` | Check `$workspaceUser` and `$workspacePassword` |
| `Failed to authenticate User` | Check your PAT or Client ID/Secret |
| `Assembly not found` | Verify `$toscaCommanderPath` points to your actual Tosca install |
| `Could not load System.Runtime Version=8.0.0.0` | You are running PS5 — must use **PowerShell 7** |
| `No results found` | Check that your TQL query is correct and the workspace contains matching objects |
