<#
.SYNOPSIS
    Creates and sends an Outlook mail message.

.DESCRIPTION
    Creates an Outlook mail message, then sends it.  Optionally, preview the message instead of sending it.

.PARAMETER To
    The recipient(s).  Separate mutliple recipients with a semi-colon (;).

.PARAMETER CC
    The "courtesy copy" recipient(s).  Separate mutliple recipients with a semi-colon (;).

.PARAMETER BCC
    The "blind courtesy copy" recipient(s).  Separate mutliple recipients with a semi-colon (;).

.PARAMETER Subject
    The message's subject.

.PARAMETER Body
    The message's body (supports embedded HTML).

.PARAMETER Attachments
    Comma-delimited list of attachments.

.PARAMETER Preview
    If set, the message will be displayed instead of sent.

.EXAMPLE
    PS > Send-Mail "recipient@domain0.tld;recipient@domain1.tld", "the subject", "the message", "path/to/attachment0;path/to/attachmentN"

Send the message without interaction.

.EXAMPLE
    PS > Send-Mail "recipient@domain0.tld;recipient@domain1.tld", "the subject", "the message", "path/to/attachment0;path/to/attachmentN" -Preview

Show the message.
#>

Function Send-Mail {

    [cmdletbinding()]
    PARAM (
        [Parameter(Mandatory=$True)]
        [String[]] $To,
        
        [Parameter(Mandatory=$True)]
        [Alias('s')]
        [String] $Subject,
        
        [Parameter(Mandatory=$False)]
        [String] $Body,
        
        [Parameter(Mandatory=$False)]
        [String[]] $CC,
        
        [Parameter(Mandatory=$False)]
        [String[]] $BCC,
        
        [Parameter(Mandatory=$False)]
        [Alias('a')]
        [String[]] $Attachments,

        [Parameter(Mandatory=$False)]
        [Alias('p')]
        [switch] $Preview
    )
    
    BEGIN {
        Write-Debug "$($MyInvocation.MyCommand.Name)::Begin"

        Write-Debug "To: $To"
        Write-Debug "Subject: $Subject"
        Write-Debug "Body: $Body"
        Write-Debug "CC: $CC"
        Write-Debug "Attachments: $Attachments"
        Write-Debug "Preview: $Preview"

        try {
            # activate existing instance
            Write-Verbose "Activating existing Outlook instance..."
            $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            $Outlook.ActiveWindow().Activate()
        }
        catch [System.Runtime.InteropServices.COMException],[System.Management.Automation.MethodInvocationException] {
            Write-Verbose "Opening new Outlook instance..."
            # open new instance

            $Outlook = New-Object -Com Outlook.Application
            # $Outlook = New-Object Outlook.Application

            # Outlook's UI isn't required to send a message
            # $namespace = $Outlook.GetNamespace("MAPI")
            # $folder = $namespace.GetDefaultFolder("olFolderInbox")
            # $explorer = $folder.GetExplorer()
            # $explorer.Display() 

            # eliminate race conditions (http://stackoverflow.com/a/461327/134367)
            # Start-Sleep -sec 4

        }
        catch [Exception] {
            # Write-Error $_.Exception.ToString()
            Throw
        }

    }
    
    PROCESS {
        Write-Debug "$($MyInvocation.MyCommand.Name)::Process"

        Try {
            # throw New-Object System.Exception 'This is an error'

            Write-Verbose "Creating message..."
            # OlItemType.olMailItem=0
            $Mail = $Outlook.CreateItem(0)

            # convert spaces and commas to semi-colons
            $Mail.To = $to -join ";"
            $Mail.Cc = $cc -join ";"
            $Mail.Bcc = $bcc -join ";"
            $Mail.Subject = $subject

            # Outlook.OlBodyFormat.olFormatUnspecified=0
            # Outlook.OlBodyFormat.olFormatPlain=1
            # Outlook.OlBodyFormat.olFormatHTML=2
            # Outlook.OlBodyFormat.olFormatRichText=3
            
            $Mail.Display()
            $Signature = $Mail.Body
            Write-Verbose "Signature: $Signature"

            #$Mail.BodyFormat = 0 # [Outlook.OlBodyFormat.olFormatHTML]
            $Mail.Body = $Body + $Signature
            # $Mail.HTMLBody = "<HTML><BODY>" + $Body + $Signature + "</BODY></HTML>" 

            # add attachments
            if ($Attachments -ne $null) {
                Foreach ($File In $Attachments) {
                    Write-Verbose "Attaching $File ..."
                    $Mail.Attachments.Add($File) | out-null
                }
            } # if

            # show window or send message
            if ($preview) { $Mail.Display() } 
            else { $Mail.Send() }

        } # try
        Catch [Exception] {
            # Write-Host $_.Exception.ToString()
            Throw
        }

    }
    
    END { 
        Write-Debug "$($MyInvocation.MyCommand.Name)::End"

        if ($Preview -eq $False) { 
            Write-Verbose "Quitting..."
            $Outlook.Quit() 
        }
    }
    
}

Set-Alias sm Send-Mail
