# MFA-DomainSleuth

A PowerShell tool that scans your email inbox and identifies which sender domains support Multi-Factor Authentication (MFA), helping you prioritize which accounts to secure with MFA.

## 📋 Overview

MFA-DomainSleuth analyzes emails from your mailbox, extracts sender domains, and cross-references them with the [2FA Directory](https://2fa.directory/) to identify which services support MFA/2FA. The results are presented in an easy-to-read HTML report.

## 🎯 Versions Available

### 1. **MFA-DomainSleuthv5.ps1** (Classic - Outlook COM)
- Uses Outlook COM Object (legacy method)
- Works with classic Outlook desktop application
- Requires Outlook to be installed locally
- **Best for:** Traditional Outlook desktop users

### 2. **MFA-DomainSleuth-Graph.ps1** (Modern - Microsoft Graph API) ⭐ **RECOMMENDED**
- Uses Microsoft Graph API
- Works with **new Outlook** and Microsoft 365
- Cloud-based authentication (OAuth 2.0)
- More secure and future-proof
- No local Outlook installation required
- **Best for:** Microsoft 365 users, new Outlook, and modern environments

## 🚀 Quick Start

### Prerequisites

**For Classic Version (MFA-DomainSleuthv5.ps1):**
- Windows PowerShell 5.1 or later
- Microsoft Outlook desktop application installed
- Active Outlook profile configured

**For Graph Version (MFA-DomainSleuth-Graph.ps1):**
- Windows PowerShell 5.1 or PowerShell 7+
- Microsoft 365 account (Exchange Online)
- Internet connection
- Microsoft Graph PowerShell SDK (auto-installed by script)

### Installation

1. Clone or download this repository:
```powershell
git clone https://github.com/yourusername/MFA-DomainSleuth.git
cd MFA-DomainSleuth
```

2. Ensure PowerShell execution policy allows script execution:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 📖 Usage

### Using the Graph API Version (Recommended)

1. **Run the script:**
```powershell
.\MFA-DomainSleuth-Graph.ps1
```

2. **First-time setup:**
   - The script will automatically install Microsoft Graph PowerShell modules if not present
   - You'll be prompted to sign in with your Microsoft 365 account
   - Grant the requested permissions (Mail.Read, Mail.ReadBasic)

3. **Follow the prompts:**
   - Enter the number of days to check back
   - Choose whether to show domains that don't support MFA
   - Select the folder to scan (Inbox, Sent Items, etc.)

4. **Review the report:**
   - The script generates `MFA-DomainSleuth.html` in the script directory
   - Open it in your web browser to see the results

### Using the Classic COM Version

1. **Ensure Outlook is installed and configured**

2. **Run the script:**
```powershell
.\MFA-DomainSleuthv5.ps1
```

3. **Select folder:** A dialog will appear to pick the Outlook folder to scan

4. **Follow the prompts:** Enter the number of days and MFA display preference

## 🔧 Microsoft Graph Setup (Graph Version)

### Automatic Setup (Default)
The script uses the Microsoft Graph PowerShell SDK with delegated permissions, which handles authentication automatically. No Azure AD app registration is required for personal use.

### Enterprise/Organization Setup (Optional)
For enterprise environments, administrators can pre-register an Azure AD application:

1. **Register an App in Azure AD:**
   - Go to [Azure Portal](https://portal.azure.com/)
   - Navigate to Azure Active Directory → App registrations
   - Click "New registration"
   - Name: "MFA-DomainSleuth"
   - Supported account types: "Accounts in this organizational directory only"
   - Click "Register"

2. **Configure API Permissions:**
   - Go to "API permissions"
   - Add permissions → Microsoft Graph → Delegated permissions
   - Add: `Mail.Read` and `Mail.ReadBasic`
   - Grant admin consent (if required by your organization)

3. **Note the Application (client) ID** for enterprise deployment

## 📊 Report Features

The generated HTML report includes:

- **Matched Domain:** The base domain of the sender
- **Supports MFA:** Yes/No indicator (color-coded)
- **Email Address:** Sender's email address
- **Documentation:** Link to MFA setup instructions
- **MFA Methods:** Available authentication methods (SMS, TOTP, hardware keys, etc.)
- **Custom Software/Hardware:** Supported authenticator apps and physical tokens
- **Recovery URL:** Account recovery information
- **Additional Domains:** Related domains for the same service
- **Keywords & Notes:** Additional service information

## 🎨 Color Coding

- **Green rows:** Domain supports MFA ✅
- **Red rows:** Domain does not support MFA ❌

## 🔒 Security & Privacy

### Graph API Version:
- Uses OAuth 2.0 for secure authentication
- Requires only read permissions (Mail.Read, Mail.ReadBasic)
- No passwords or credentials stored
- Tokens are managed by Microsoft Graph SDK
- Complies with Microsoft 365 security policies

### Classic COM Version:
- Uses local Outlook profile credentials
- No data sent to external services except the 2FA Directory API
- All processing done locally

### Both Versions:
- Email content is NOT uploaded or shared
- Only sender email addresses are processed
- Report is generated locally on your computer

## 🛠️ Troubleshooting

### Graph API Version

**Issue:** "Microsoft Graph PowerShell SDK is not installed"
- **Solution:** The script will auto-install the required modules. Ensure you have internet access and permissions to install PowerShell modules.

**Issue:** "Failed to connect to Microsoft Graph"
- **Solution:**
  - Ensure you have a valid Microsoft 365 account
  - Check your internet connection
  - Verify you granted the requested permissions during sign-in

**Issue:** "No emails found"
- **Solution:**
  - Try a longer date range
  - Verify you selected the correct folder
  - Check that your mailbox has emails in the specified timeframe

### Classic COM Version

**Issue:** "Cannot create Outlook COM object"
- **Solution:** Ensure Microsoft Outlook is installed and properly configured

**Issue:** Slow performance with large mailboxes
- **Solution:** Reduce the number of days to check or select a specific folder

## 🔄 Migrating from COM to Graph Version

The Graph API version provides the same functionality with these improvements:

1. ✅ **Cloud-based:** No local Outlook installation required
2. ✅ **Modern auth:** OAuth 2.0 instead of COM object dependency
3. ✅ **New Outlook compatible:** Works with Microsoft 365 and new Outlook
4. ✅ **Better error handling:** More detailed feedback
5. ✅ **Enhanced reporting:** Additional metadata in the report

Simply run `MFA-DomainSleuth-Graph.ps1` instead of the classic version!

## 📝 Data Source

This tool uses the [2FA Directory API](https://2fa.directory/) to determine MFA support. The directory is community-maintained and regularly updated.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This tool is provided as-is for informational purposes. Always verify MFA availability directly with service providers. The accuracy of MFA support information depends on the 2FA Directory database.

## 🔗 Related Resources

- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/)
- [2FA Directory](https://2fa.directory/)
- [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/)
- [Multi-Factor Authentication Best Practices](https://www.microsoft.com/en-us/security/business/identity-access/mfa-multi-factor-authentication)

---

**Made with ❤️ to help secure your online accounts**
