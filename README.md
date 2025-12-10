## CertificateExporter

**CertificateExporter** is a lightweight Windows GUI tool that helps you export certificates from the Windows Certificate Store and prepare Linux-ready TLS files with consistent naming and safe export handling.

### Who it's for
Targets IT Pros and admins who need a fast, repeatable way to move certificates from Windows environments to Linux-based services (NGINX, Apache, HAProxy, etc.).

### Certificate sources
Reads certificates from:

- **CurrentUser\Personal** (`Cert:\CurrentUser\My`)
- **LocalMachine\Personal** (`Cert:\LocalMachine\My`)

### Validation
- Filters for **certificates with private keys** (required for PFX export).

### Export workflows
- **Export Selected → PFX only**
- **Export Selected → PFX → Linux files**
- **Convert Existing PFX → Linux files**

### Linux output (via OpenSSL)
Generates:

- `<name>-cert.pem` *(leaf certificate)*
- `<name>-privkey.key` *(private key)*
- `<name>-chain.cer` *(intermediate chain)*
- `<name>-fullchain.cer` *(leaf + intermediates)*

### Folder naming
- Export folder naming is based on the certificate **SSL Common Name (CN)** to keep files organized and human-readable.

### Safe overwrite behavior
- Prompts if export files exist.
- Allows:
  - **Overwrite**
  - **Add timestamp**
- Supports canceling without partial exports.

### User experience
- **Progress indicator** built-in for clear feedback during export/conversion steps.
- **Password prompts are GUI-based** (no console prompts), with confirmation and optional show/hide.

### Logging
Verbose logging runs behind the scenes:

- Logs are automatically stored in `.\Logs`
- New runs archive older logs with timestamps

### OpenSSL detection
- Tries system **PATH** first.
- Falls back to common install locations including:
  - `C:\Program Files\OpenSSL-Win64\bin`

---

## Requirements

### OpenSSL
This tool relies on **OpenSSL** to convert PFX files into Linux-ready certificate/key/chain outputs.

Recommended Windows distribution:
- OpenSSL for Windows (Win64)

You can download OpenSSL here:
- https://slproweb.com/products/Win32OpenSSL.html

> Tip: If you install OpenSSL to the default location, the tool should auto-detect it.  
> Otherwise, set the path manually using the **OpenSSL path** field in the UI.

### PowerShell version
Recommended:
- **Windows PowerShell 5.1** (best compatibility for WinForms on Windows)
- or **PowerShell 7.2+** on Windows

### Permissions / Run as Administrator
You **do not** need admin rights for:
- Browsing and exporting certificates from **CurrentUser\Personal**

You **should run as Administrator** if you plan to:
- Access or export from **LocalMachine\Personal**
- Export certs where the private key ACL requires elevated rights

If you experience access errors when selecting a LocalMachine certificate, re-run the tool with:
- **Run as Administrator**
