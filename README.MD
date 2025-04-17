# Send-MailKitMessageHC

This PowerShell function `Send-MailKitMessageHC` allows you to send emails using the powerful [MailKit](https://github.com/jstedfast/MailKit) and [MimeKit](https://github.com/jstedfast/MimeKit) .NET libraries. It provides a flexible way to construct and send emails with features like attachments, HTML bodies, and various SMTP connection options.

### Example: Send an authenticated email

```powershell
$SmtpUserName = 'smtpUser'
$SmtpPassword = 'smtpPassword'

$securePassword = ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($SmtpUserName, $securePassword)

$params = @{
    SmtpServerName      = 'SMT_SERVER@example.com'
    SmtpPort            = 587
    SmtpConnectionType  = 'StartTls'
    Credential          = $credential
    From                = 'm@example.com'
    To                  = '007@example.com'
    Body                = '<p>Mission details in attachment</p>'
    Subject             = 'For your eyes only'
    Priority            = 'High'
    Attachments         = @('c:\Mission.ppt', 'c:\ID.pdf')
    MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
    MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
}

Send-MailKitMessageHC @params
```

### Example: Send an email without authentication

```powershell
$params = @{
    SmtpServerName      = 'SMT_SERVER@example.com'
    SmtpPort            = 25
    From                = 'hacker@example.com'
    FromDisplayName     = 'White hat hacker'
    Bcc                 = @('james@example.com', 'mike@example.com')
    Body                = '<h1>You have been hacked</h1>'
    Subject             = 'Oops'
    MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
    MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
}

Send-MailKitMessageHC @params
```

## Requirements

### Install MailKit and MimeKit assembly

```powershell
$params = @{
    Source           = 'https://www.nuget.org/api/v2'
    SkipDependencies = $true
    Scope            = 'AllUsers'
}
Install-Package @params -Name 'MailKit'
Install-Package @params -Name 'MimeKit'
```
