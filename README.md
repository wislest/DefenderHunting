# 🛡️ DefenderHunter

![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?logo=powershell&logoColor=white)
![KQL](https://img.shields.io/badge/KQL-Advanced%20Hunting-0078D4)
![Defender XDR](https://img.shields.io/badge/Microsoft-Defender%20XDR-0078D4?logo=microsoft)
![Sentinel](https://img.shields.io/badge/Microsoft-Sentinel-0089D6?logo=microsoftazure&logoColor=white)
![MITRE ATT&CK](https://img.shields.io/badge/MITRE-ATT%26CK%20mapped-c4162a)
![Detection-as-Code](https://img.shields.io/badge/Detection--as--Code-✓-2ea44f)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

> **Detection-as-Code threat-hunting framework for Microsoft Defender XDR & Sentinel.**
> Automates IOC enrichment and Advanced Hunting across endpoints, with **22 ransomware-family KQL query packs** mapped to MITRE ATT&CK.

---

## ✨ Features

- **Automated IOC hunting** via Microsoft Defender Advanced Hunting (KQL) over the Defender API.
- **Flexible IOC ingestion** — CSV/JSON input with **auto hash-type detection** (MD5 / SHA1 / SHA256).
- **22 ransomware detection packs** under `Queries/` (see below).
- **Modular PowerShell** design (`Modules/`, `Scripts/`) with export pipelines and reporting.
- **Secure by default** — credentials live in `config.json` / `.env` (git-ignored); never hard-coded. See [SECURITY.md](SECURITY.md).

## 🎯 Ransomware coverage (`Queries/`)

`Akira` · `Babuk Locker` · `BlackByte` · `BlackSuit` · `Cactus` · `Cicada3301` · `EMBARGO` · `Fog` · `Hunters International` · `Inc Ransom` · `Interlock` · `Lockbit 3.0` · `Lynx` · `Medusa` · `Nebula` · `Play` · `RansomHub` · `Rhysida` · `STOP/DJVU` · … (plus generic/unknown buckets)

## 🚀 Quick start

**Prerequisites:** PowerShell 5.1+, an Entra ID app registration with Microsoft Defender API permissions (`AdvancedHunting.Read.All`).

```powershell
# 1) Configure Defender API credentials
Copy-Item config.json.example config.json
#   edit config.json -> TenantId, ClientId, ClientSecret

# 2) (optional) Configure export/report settings
Copy-Item export-config.json.example export-config.json
#   edit export-config.json -> WorkspaceId, ExportPath, EmailRecipient, SmtpServer

# 3) Validate connectivity to the Defender API
.\CheckDefenderAPI.ps1

# 4) Hunt a set of IOCs
.\Main.ps1 -IocsFile .\my-iocs.csv

# ...or run the scheduled export workflow
.\Main.ps1 -ConfigPath .\export-config.json
```

## 📑 IOC input format

CSV with three columns; hashes are a JSON object:

```csv
FullPath,FamilyName,FileHashes
C:\temp\malware.exe,Akira,"{""md5"":""…"",""sha1"":""…"",""sha256"":""…""}"
```

Have a simpler `FamilyName,SHA256` list? Convert it automatically (hash type detected by length):

```powershell
.\Convert-IocsFormat.ps1 -InputFile simple-iocs.csv -OutputFile formatted-iocs.csv
```

## 📂 Repository layout

| Path | Purpose |
|------|---------|
| `Main.ps1` | **Entry point** — IOC hunting + scheduled export |
| `CheckDefenderAPI.ps1` | Verify Defender API authentication/connectivity |
| `Convert-IocsFormat.ps1` | Normalize simple IOC lists into the required CSV/JSON format |
| `Queries/` | 22 ransomware-family KQL query packs |
| `Modules/DefenderHunter/` | PowerShell module (export core, queries, IOC hunter) |
| `Scripts/` | Export pipeline scripts |
| `config.json.example`, `export-config.json.example` | Configuration templates (copy → fill → **never commit the real files**) |
| `SECURITY.md` | Credential-handling & security guidance |

## 🔐 Security

- **Never commit `config.json` / `.env` / `export-config.json`** — they hold tenant credentials and are git-ignored by default.
- Rotate the app `ClientSecret` regularly and grant the app **least-privilege** Defender API permissions.
- See [SECURITY.md](SECURITY.md) for full guidance.

## 📄 License

Released under the **[MIT License](LICENSE)**.
