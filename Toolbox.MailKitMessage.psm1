function Send-MailKitMessageHC {
    <#
        .SYNOPSIS
            Send an email using MailKit and MimeKit assemblies.

        .DESCRIPTION
            This function sends an email using the MailKit and MimeKit
            assemblies. It requires the assemblies to be installed before
            calling the function.

        .PARAMETER MailKitAssemblyPath
            The path to the MailKit assembly.

            Example: 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'

            Install-Package -Name 'MailKit' -Source "https://www.nuget.org/api/v2" -SkipDependencies -Scope AllUsers

        .PARAMETER MimeKitAssemblyPath
            The path to the MimeKit assembly.

            Example: 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'

            Install-Package -Name 'MimeKit' -Source "https://www.nuget.org/api/v2" -SkipDependencies -Scope AllUsers

        .PARAMETER SmtpServerName
            The name of the SMTP server.

        .PARAMETER SmtpPort
            The port of the SMTP server.

        .PARAMETER SmtpConnectionType
            The connection type for the SMTP server.

            Valid values are:
            - 'None'
            - 'Auto'
            - 'SslOnConnect'
            - 'StartTlsWhenAvailable'
            - 'StartTls'

        .PARAMETER Credential
            The credential object containing the username and password.

        .PARAMETER From
            The sender's email address.

        .PARAMETER To
            The recipient's email address.

        .PARAMETER Body
            The body of the email, HTML is supported.

        .PARAMETER Subject
            The subject of the email.

        .PARAMETER Attachments
            An array of file paths to attach to the email.

        .PARAMETER Priority
            The email priority.

            Valid values are:
            - 'Low'
            - 'Normal'
            - 'High'

        .EXAMPLE
            # Send an email with StartTls and credential

            $SmtpUserName = "smptUser"
            $SmtpPassword = "smtpPasswrod"

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

        .EXAMPLE
            # Send an anonymous email

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
    #>

    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [string]$MailKitAssemblyPath,
        [parameter(Mandatory)]
        [string]$MimeKitAssemblyPath,
        [parameter(Mandatory)]
        [string]$SmtpServerName,
        [parameter(Mandatory)]
        [ValidateRange(0, 65535)]
        [int]$SmtpPort,
        [parameter(Mandatory)]
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
        [string]$From,
        [parameter(Mandatory)]
        [string[]]$To,
        [parameter(Mandatory)]
        [string]$Body,
        [parameter(Mandatory)]
        [string]$Subject,
        [int]$MaxAttachmentSize = 20MB,
        [ValidateSet(
            'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
        )]
        [string]$SmtpConnectionType = 'None',
        [ValidateSet('Normal', 'Low', 'High')]
        [string]$Priority = 'Normal',
        [string[]]$Attachments,
        [PSCredential]$Credential
    )

    begin {
        function Test-IsAssemblyLoaded {
            param (
                [String]$Name
            )
            foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
                if ($assembly.FullName -like "$Name, Version=*") {
                    return $true
                }
            }
            return $false
        }

        function Add-Attachments {
            param (
                [string[]]$Attachments,
                [MimeKit.Multipart]$BodyMultiPart
            )

            $attachmentList = New-Object System.Collections.ArrayList($null)

            $tempFolder = "$env:TEMP\Send-MailKitMessageHC {0}" -f (Get-Random)
            $totalSizeAttachments = 0

            foreach (
                $attachmentPath in
                $Attachments | Sort-Object -Unique
            ) {
                try {
                    #region Test if file exists
                    try {
                        $attachmentItem = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop

                        if ($attachmentItem.PSIsContainer) {
                            Write-Warning "Attachment '$attachmentPath' is a folder, not a file"
                            continue
                        }
                    }
                    catch {
                        Write-Warning "Attachment '$attachmentPath' not found"
                        continue
                    }
                    #endregion

                    $totalSizeAttachments += $attachmentItem.Length

                    if ($attachmentItem.Extension -eq '.xlsx') {
                        #region Copy Excel file, open file cannot be sent
                        if (-not(Test-Path $tempFolder)) {
                            $null = New-Item $tempFolder -ItemType 'Directory'
                        }

                        $params = @{
                            LiteralPath = $attachmentItem.FullName
                            Destination = $tempFolder
                            PassThru    = $true
                            ErrorAction = 'Stop'
                        }

                        $copiedItem = Copy-Item @params

                        $null = $attachmentList.Add($copiedItem)
                        #endregion
                    }
                    else {
                        $null = $attachmentList.Add($attachmentItem)
                    }

                    #region Check size of attachments
                    if ($totalSizeAttachments -ge $MaxAttachmentSize) {
                        $M = "The maximum allowed attachment size of {0} MB has been exceeded ({1} MB). No attachments were added to the email. Check the log folder for details." -f
                        ([math]::Round(($MaxAttachmentSize / 1MB))),
                        ([math]::Round(($totalSizeAttachments / 1MB), 2))

                        Write-Warning $M

                        return [PSCustomObject]@{
                            AttachmentLimitExceededMessage = $M
                        }
                    }
                }
                catch {
                    Write-Warning "Failed to add attachment '$attachmentPath': $_"
                }
            }
            #endregion

            foreach (
                $attachmentItem in
                $attachmentList
            ) {
                try {
                    Write-Verbose "Add mail attachment '$($attachmentItem.Name)'"

                    $attachment = New-Object MimeKit.MimePart

                    $attachment.Content = New-Object MimeKit.MimeContent(
                        [System.IO.File]::OpenRead($attachmentItem.FullName)
                    )

                    $attachment.ContentDisposition = New-Object MimeKit.ContentDisposition

                    $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64

                    $attachment.FileName = $attachmentItem.Name

                    $bodyMultiPart.Add($attachment)
                }
                catch {
                    Write-Warning "Failed to add attachment '$attachmentItem': $_"
                }
            }
        }

        try {
            #region Test To
            foreach ($email in $To) {
                if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                    throw "To email address '$email' not valid."
                }
            }
            #endregion

            #region Load MimeKit assembly
            if (-not(Test-IsAssemblyLoaded -Name 'MimeKit')) {
                try {
                    Write-Verbose "Load MimeKit assembly '$MimeKitAssemblyPath'"
                    Add-Type -Path $MimeKitAssemblyPath
                }
                catch {
                    throw "Failed to load MimeKit assembly '$MimeKitAssemblyPath': $_"
                }
            }
            #endregion

            #region Load MailKit assembly
            if (-not(Test-IsAssemblyLoaded -Name 'MailKit')) {
                try {
                    Write-Verbose "Load MailKit assembly '$MailKitAssemblyPath'"
                    Add-Type -Path $MailKitAssemblyPath
                }
                catch {
                    throw "Failed to load MailKit assembly '$MailKitAssemblyPath': $_"
                }
            }
            #endregion
        }
        catch {
            throw "Failed to send email to '$To': $_"
        }
    }

    process {
        try {
            $message = New-Object -TypeName 'MimeKit.MimeMessage'

            #region Create body with attachments
            $bodyPart = New-Object MimeKit.TextPart('html')
            $bodyPart.Text = $Body

            $bodyMultiPart = New-Object MimeKit.Multipart('mixed')
            $bodyMultiPart.Add($bodyPart)

            if ($Attachments) {
                $params = @{
                    Attachments   = $Attachments
                    BodyMultiPart = $bodyMultiPart
                }
                $addAttachments = Add-Attachments @params

                if ($addAttachments.AttachmentLimitExceededMessage) {
                    $bodyPart.Text += '<p><i>{0}</i></p>' -f
                    $addAttachments.AttachmentLimitExceededMessage
                }
            }

            $message.Body = $bodyMultiPart
            #endregion

            $message.From.Add($From)

            foreach ($email in $To) {
                $message.To.Add($email)
            }

            $message.Subject = $Subject

            #region Set priority
            switch ($Priority) {
                'Low' {
                    $message.Headers.Add('X-Priority', '5 (Lowest)')
                    break
                }
                'Normal' {
                    $message.Headers.Add('X-Priority', '3 (Normal)')
                    break
                }
                'High' {
                    $message.Headers.Add('X-Priority', '1 (Highest)')
                    break
                }
                default {
                    throw "Priority type '$_' not supported"
                }
            }
            #endregion

            $smtp = New-Object -TypeName 'MailKit.Net.Smtp.SmtpClient'

            $smtp.Connect(
                $SmtpServerName, $SmtpPort,
                [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType
            )

            if ($Credential) {
                $smtp.Authenticate(
                    $Credential.UserName,
                    $Credential.GetNetworkCredential().Password
                )
            }

            Write-Verbose "Send mail to '$To' with subject '$Subject'"

            $null = $smtp.Send($message)
            $smtp.Disconnect($true)
            $smtp.Dispose()
        }
        catch {
            throw "Failed to send email to '$To': $_"
        }
    }
}

Export-ModuleMember -Function * -Alias *