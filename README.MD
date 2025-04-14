# Send-MailKitMessageHC

This PowerShell function `Send-MailKitMessageHC` allows you to send emails using the powerful [MailKit](https://github.com/jstedfast/MailKit) and [MimeKit](https://github.com/jstedfast/MimeKit) .NET libraries. It provides a flexible way to construct and send emails with features like attachments, HTML bodies, and various SMTP connection options.

### Example: Send an anonymous email

```powershell
$params = @{
    MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
    MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
    SmtpServerName      = 'SMT_SERVER@example.com'
    SmtpPort            = 25
    From                = 'bob@example.com'
    To                  = @('james@example.com', 'mike@example.com')
    Body                = '<h1>Welcome</h1><p>Welcome email</p>'
    Subject             = 'Welcome'
}

Send-MailKitMessageHC @params
```

### Example: Send an authenticated email

```powershell
$SmtpUserName = 'smtpUser'
$SmtpPassword = 'smtpPassword'

$securePassword = ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($SmtpUserName, $securePassword)

$params = @{
    MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
    MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
    SmtpServerName      = 'SMT_SERVER@example.com'
    SmtpPort            = 587
    SmtpConnectionType  = 'StartTls'
    Credential          = $credential
    From                = 'm@example.com'
    To                  = '007@example.com'
    Body                = '<p>See attachment for your next mission</p>'
    Subject             = 'For your eyes only'
    Priority            = 'High'
    Attachments         = 'c:\Secret file.txt'
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
Install-Package @params -Name 'MimeKit'
Install-Package @params -Name 'MailKit'
```
