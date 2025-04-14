# Send-MailKitMessageHC

This PowerShell function `Send-MailKitMessageHC` allows you to send emails using the powerful [MailKit](https://github.com/jstedfast/MailKit) and [MimeKit](https://github.com/jstedfast/MimeKit) .NET libraries. It provides a flexible way to construct and send emails with features like attachments, HTML bodies, and various SMTP connection options.

## Parameters

- SMTP server name and port
- SMTP connection security (**None**, **Auto**, **SslOnConnect**, **StartTlsWhenAvailable** and **StartTls**)
- **Credentials** for authenticated SMTP servers
- **To**, **From**, **Subject**, **Attachements**, **Body**, **Priority**, ..

### -MailKitAssemblyPath

The path to the MailKit assembly (`MailKit.dll`).

**Example:** `'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'`

### -MimeKitAssemblyPath

The path to the MImeKit assembly (`MimeKit.dll`).

**Example:** `'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'`

### Requirements:

### Install MailKit
```powershell
Install-Package -Name 'MailKit' -Source "https://www.nuget.org/api/v2" -SkipDependencies -Scope AllUsers
```

### Install MimeKit
```powershell
Install-Package -Name 'MimeKit' -Source "https://www.nuget.org/api/v2" -SkipDependencies -Scope AllUsers
```
