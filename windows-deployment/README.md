# Windows 11 Unattended Deployment + Entra Auto-Enrollment

A zero-touch Windows 11 deployment solution combining an **unattended answer file** with a **WCD provisioning package** — drop it on a USB drive, boot a new PC, and it fully configures itself with no IT interaction.

## What it does

| Step | Tool | Description |
|---|---|---|
| 1. Partition & Install | `autounattend.xml` | Wipes disk, creates EFI + Windows partitions, installs Windows |
| 2. Skip OOBE | `autounattend.xml` | Bypasses all setup screens (EULA, region, account creation) |
| 3. Auto-login | `autounattend.xml` | Logs in as local admin once to apply provisioning package |
| 4. Connect to WiFi | `customizations.xml` (ppkg) | Pre-configures corporate WiFi profile |
| 5. Rename PC | `customizations.xml` (ppkg) | Sets hostname to `OFFICE--<SERIAL>` using device serial number |
| 6. Entra enrollment | `customizations.xml` (ppkg) | Enrolls device into Microsoft Entra ID (Azure AD) via BPRT token |
| 7. Apply policies | `customizations.xml` (ppkg) | Sets sideloading and application management policies |

## Files

- `autounattend.xml` — Windows Setup answer file (place at root of USB/ISO)
- `customizations.xml` — Windows Configuration Designer source (open in WCD to rebuild the `.ppkg`)

## How to use

### Prerequisites
- Windows 11 ISO (create bootable USB with Rufus)
- Windows Configuration Designer (from Microsoft Store)
- An active Bulk Primary Refresh Token from your Entra admin center

### Build the provisioning package
1. Open `customizations.xml` in Windows Configuration Designer
2. Replace all `YOUR_*` placeholders with real values
3. Export → **Provisioning package** (.ppkg)

### Prepare the USB
```
USB Root/
├── autounattend.xml        ← answer file (auto-detected by Windows Setup)
├── Provisioning/
│   └── Autopilot.ppkg     ← WCD package (auto-applied during OOBE)
└── sources/               ← Windows 11 installation files
```

### Deploy
Boot the target PC from USB. Windows Setup detects `autounattend.xml` automatically, installs silently, and applies the provisioning package on first boot.

## Generating a new BPRT token

1. Go to **Entra admin center** → Devices → Bulk enrollment tokens
2. Click **+ New token** → set expiry → copy the token
3. Paste into `customizations.xml` at `<BPRT>` and rebuild the package

> **Security note:** BPRT tokens grant the ability to enroll devices into your tenant. Treat them like passwords — never commit a real token to a public repository. Revoke old/unused tokens regularly.

## Tech used

- Windows Setup answer file format (Microsoft-Windows-Shell-Setup)
- Windows Configuration Designer / WICD
- Microsoft Entra ID (Azure AD) bulk enrollment
- WPA2 WLAN profile provisioning
