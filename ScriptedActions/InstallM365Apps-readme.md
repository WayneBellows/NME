# Install-M365Apps.ps1

## 🧩 Overview

This PowerShell script installs or removes **Microsoft 365 Apps for Enterprise** (formerly Office 365 ProPlus) on **Azure Virtual Desktop (AVD)** session hosts.  
It automatically detects whether the VM is running a **single-session** or **multi-session** edition of Windows and configures Office accordingly.

---

## ⚙️ Key Features

- 🧠 **Auto-detects AVD edition**
  - **Multi-session** → Enables **Shared Computer Activation (SCA)** and **disables** auto-updates.
  - **Single-session** → Standard per-user activation; **enables** auto-updates.
- 🪄 Supports **adding or removing** components:
  - **Add**: Visio and/or Project.
  - **Remove**: Specific core apps (Word, Excel, Outlook, PowerPoint, OneNote, Access, Publisher).
- 🧱 Uses the **Office Deployment Tool (ODT)** and generates configuration XML on the fly.
- 🧰 Cleans up temporary files automatically.
- 🪣 Friendly for AVD image builds (Nerdio, Packer, Azure Image Builder, etc.).

---

## 🚀 Parameters

| Parameter       | Type       | Required | Description                                                                                 |
|-----------------|------------|----------|---------------------------------------------------------------------------------------------|
| `-Applications` | `String[]` | No       | For `-Type Remove`, list core apps to **exclude**. For `-Type Add`, list `Visio`/`Project`. Leave empty to install the full core suite. |
| `-Version`      | `String`   | Yes      | Office bitness: `"32"` or `"64"`.                                                           |
| `-Type`         | `String`   | Yes      | `"Add"` (Visio/Project) or `"Remove"` (core app exclusions).                                |

> Guardrails: When `-Type Add`, only `Visio`/`Project` are allowed. When `-Type Remove`, only core apps are allowed.

---

## 💡 Behavior by OS Type

| OS Type            | Shared Computer Licensing | Office Updates | Typical Use                         |
|--------------------|---------------------------|----------------|-------------------------------------|
| **Multi-session**  | ✅ Enabled (`SCA=1`)      | ❌ Disabled    | AVD pooled / multi-user hosts       |
| **Single-session** | ❌ Disabled                | ✅ Enabled     | AVD personal / dedicated hosts      |

---

## 🧠 Examples

### 1) Install full suite (no Visio/Project), 64-bit
Installs Word, Excel, PowerPoint, Outlook, OneNote, Access, and Publisher.
```powershell
.\Install-M365Apps.ps1 -Type Remove -Version 64
```

### 2) Exclude specific apps (e.g., Outlook and Publisher)
```powershell
.\Install-M365Apps.ps1 -Applications @("Outlook", "Publisher") -Type Remove -Version 64
```

### 3) Add Visio only
```powershell
.\Install-M365Apps.ps1 -Applications @("Visio") -Type Add -Version 64
```

### 4) Add both Visio and Project
```powershell
.\Install-M365Apps.ps1 -Applications @("Visio", "Project") -Type Add -Version 64
```

### 5) Run directly from GitHub (no download)
```powershell
iex "& { $(Invoke-RestMethod 'https://raw.githubusercontent.com/<your-username>/<your-repo>/main/scripts/Install-M365Apps.ps1') } -Type Remove -Version 64"
```

### 6) Download and run locally
```powershell
$scriptUrl = 'https://raw.githubusercontent.com/<your-username>/<your-repo>/main/scripts/Install-M365Apps.ps1'
$scriptPath = "$env:TEMP\Install-M365Apps.ps1"

Invoke-WebRequest -Uri $scriptUrl -OutFile $scriptPath
powershell -ExecutionPolicy Bypass -File $scriptPath -Type Remove -Version 64
```

### 7) Run with logging
```powershell
.\Install-M365Apps.ps1 -Type Remove -Version 64 | Tee-Object -FilePath "C:\Temp\OfficeInstall.log"
```

## 🧰 Output and Cleanup

- Console logs include timestamps and progress messages.

- Temporary artifacts are removed automatically:

  - C:\Temp\<GUID>
  - C:\AVDImage (if created)

## 🪪 Auto-generated XML (for reference)
### Multi-session
```xml
<Configuration>
  <Add Channel="MonthlyEnterprise"></Add>
  <RemoveMSI />
  <Updates Enabled="FALSE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="SharedComputerLicensing" Value="1" />
</Configuration>
```
### Single-session
```xml
<Configuration>
  <Add Channel="MonthlyEnterprise"></Add>
  <RemoveMSI />
  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
```
