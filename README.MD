# Send-MailKitMessageHC

This PowerShell function `Send-MailKitMessageHC` allows you to send emails using the powerful [MailKit](https://github.com/jstedfast/MailKit) and [MimeKit](https://github.com/jstedfast/MimeKit) .NET libraries. It provides a flexible way to construct and send emails with features like attachments, HTML bodies, and various SMTP connection options.

## Synopsis

Send an email using MailKit and MimeKit assemblies.

## Description

This function sends an email using the MailKit and MimeKit assemblies. It requires these assemblies to be installed before calling the function. The function supports various features, including:

- Specifying the SMTP server name and port.
- Choosing the SMTP connection security (None, Auto, SslOnConnect, StartTlsWhenAvailable, StartTls).
- Providing credentials for authenticated SMTP servers.
- Setting the sender and recipient email addresses.
- Defining the email body (HTML supported).
- Setting the email subject.
- Attaching multiple files to the email.
- Setting the email priority (Low, Normal, High).
- Limiting the maximum size of attachments.

## Parameters

### -MailKitAssemblyPath

The path to the MailKit assembly (`MailKit.dll`).

**Example:** `'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'`

### -MimeKitAssemblyPath

The path to the MImeKit assembly (`MimeKit.dll`).

**Example:** `'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'`

**Installation:**

```powershell
Install-Package -Name 'MailKit' -Source "[https://www.nuget.org/api/v2](https://www.nuget.org/api/v2)" -SkipDependencies -Scope AllUsers
Install-Package -Name 'MimeKit' -Source "https://www.nuget.org/api/v2" -SkipDependencies -Scope AllUsers
```
