# Tosca-TQL-Export-
Connects to a Tosca workspace via the TCAPI (Tosca Commander API) and runs a TQL query to export data to XML
# Tosca TQL Export — README

## What this does

`Run-ToscaTQL.ps1` connects to a Tosca workspace via the **TCAPI** (Tosca Commander API) and runs a TQL query to export data to XML — with no UI required. It is designed for bulk export of historical test execution data but can query any Tosca object type.

---

## Requirements

- **PowerShell 7+** (not Windows PowerShell 5 — the Tosca DLLs require .NET 8)
- **Tosca Commander installed locally** at the path in `$toscaCommanderPath`
- Access to your Tosca Server (for authentication) and workspace `.tws` file

---

## Setup

Open `Run-ToscaTQL.ps1` and fill in the CONFIG section at the top:

| Variable | Description |
|---|---|
| `$toscaCommanderPath` | Path to your Tosca Commander install folder |
| `$toscaServerUrl` | URL of your Tosca Server (e.g. `http://localhost/`) |
| `$workspacePath` | Full path to your `.tws` workspace file |
| `$outputFile` | Where to write the XML output |
| `$personalAccessToken` | PAT from Tosca Server > Profile > Personal Access Tokens |
| `$clientId` / `$clientSecret` | Alternative OAuth credentials from Tosca Server > Administration > OAuth Clients |
| `$workspaceUser` / `$workspacePassword` | Your Tosca Commander workspace login |
| `$tql` | The TQL query to run (see below) |

---

## Running the script

Open a fresh PowerShell 7 window and run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\Run-ToscaTQL.ps1
```

Or to run silently and redirect output to a log file:

```powershell
.\Run-ToscaTQL.ps1 > C:\Temp\export_log.txt 2>&1
```

To schedule it (e.g. nightly via Task Scheduler), use:

```
Program:  pwsh.exe
Arguments: -ExecutionPolicy Bypass -File "C:\Path\To\Run-ToscaTQL.ps1"
```

---

## TQL queries for execution history

The script ships with `=>SUBPARTS:ExecutionLogEntry` as the default query. This retrieves all execution log entries (the historical record of every test run). Several other useful queries are pre-written in the CONFIG section — just uncomment the one you need:

| Query | Description |
|---|---|
| `=>SUBPARTS:ExecutionLogEntry` | All execution history |
| `=>SUBPARTS:ExecutionLogEntry[ExecutionStatus=Failed]` | Failed runs only |
| `=>SUBPARTS:ExecutionLogEntry[ExecutionStatus=Passed]` | Passed runs only |
| `=>SUBPARTS:ExecutionLogEntry[StartedAt>='2025-01-01']` | Results after a date |
| `=>SUBPARTS:ExecutionLogEntry[TestCaseName='My Test']` | Results for a specific test case |
| `=>SUBPARTS:ExecutionList` | Top-level execution list containers |
| `=>SUBPARTS:TestCase` | All test cases (not execution results) |

TQL syntax reference: [Tricentis documentation](https://documentation.tricentis.com)

---

## Output format

The script writes an XML file at `$outputFile`. Each result is a child element of `<Results>`. For execution log entries it looks like:

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

- **Filtering by date range** in the TQL query to break the export into smaller chunks
- Running the script on the Tosca Server machine itself to avoid network overhead
- Increasing PowerShell memory if you hit out-of-memory errors: `pwsh.exe -NoProfile -ExecutionPolicy Bypass -File .\Run-ToscaTQL.ps1`

---

## Troubleshooting

| Error | Fix |
|---|---|
| `Api already initialized` | Open a fresh PS7 window before running |
| `Login Failed` | Check workspace username/password |
| `Failed to authenticate User` | Check PAT or Client ID/Secret |
| `Assembly not found` | Verify `$toscaCommanderPath` points to your Tosca install |
| `Could not load System.Runtime Version=8.0.0.0` | You are running PS5 — switch to PS7 |
