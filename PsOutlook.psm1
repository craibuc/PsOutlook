# Add-Type -assembly "Microsoft.Office.Interop.Outlook"

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
    Send-Mail "recipient@domain0.tld;recipient@domain1.tld", "the subject", "the message", "path/to/attachment0;path/to/attachmentN"
    
#>

Function Send-Mail {

    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$True,Position=0)][String] $To,
        [Parameter(Mandatory=$True,Position=1)][String] $Subject,
        [Parameter(Mandatory=$True,Position=2)][String] $Body,
        [Parameter(Mandatory=$False,Position=3)][String] $CC,
        [Parameter(Mandatory=$False,Position=4)][String] $BCC,
        [Parameter(Mandatory=$False,Position=5)][String] $Attachments,
        [Parameter(Mandatory=$False,Position=6)][switch] $Preview
    )
    
    begin {
        Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"

        try {
            # activate existing instance
            $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            $Outlook.ActiveWindow().Activate()
        }
        catch [System.Management.Automation.MethodInvocationException] {
            # open new instance
            $Outlook = new-object -com Outlook.Application
            # $Outlook = New-Object Outlook.Application
            $namespace = $Outlook.GetNamespace("MAPI")
            $folder = $namespace.GetDefaultFolder("olFolderInbox")
            $explorer = $folder.GetExplorer()
            $explorer.Display() 
        }
        catch [Exception] {
            Write-Host $_.Exception.ToString()
        }

    }
    
    process {

        # try {
        #     # activate existing instance
        #     $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        #     $Outlook.ActiveWindow().Activate()
        # }
        # catch [System.Management.Automation.MethodInvocationException] {
        #     # open new instance
        #     #$Outlook = new-object -com Outlook.Application
        #     $Outlook = New-Object Outlook.Application
        #     $namespace = $Outlook.GetNamespace("MAPI")
        #     $folder = $namespace.GetDefaultFolder("olFolderInbox")
        #     $explorer = $folder.GetExplorer()
        #     $explorer.Display() 
        # }
        # catch [Exception] {
        #     Write-Host $_.Exception.ToString()
        # }
        
        # OlItemType.olMailItem = 0
        $Mail = $Outlook.CreateItem(0)

        # convert spaces and commas to semi-colons
        $Mail.To = $to.Replace(",",";").Replace(" ",";")
        $Mail.Cc = $cc.Replace(",",";").Replace(" ",";")
        $Mail.Bcc = $bcc.Replace(",",";").Replace(" ",";")

        $Mail.Subject = $subject

        # Outlook.OlBodyFormat.olFormatUnspecified=0
        # Outlook.OlBodyFormat.olFormatPlain=1
        # Outlook.OlBodyFormat.olFormatHTML = 2
        # Outlook.OlBodyFormat.olFormatRichText=3

        $Mail.BodyFormat = 2 # [Outlook.OlBodyFormat.olFormatHTML]
        $Mail.Body = $body
        $Mail.HTMLBody = "<HTML><BODY>" + $body + "</BODY></HTML>"

        # split attachment string; loop and attach
        #$File = "C:\Users\cb2215\Desktop\dx.sql"
        #$Mail.Attachments.Add($File) | out-null

        if ($preview) {
            Write-Debug "Verbose..."
            $Mail.Display()
        }
        else {
            Write-Debug "Verbose..."
            $Mail.Send()
        }
        
    }
    
    end {
        Write-Verbose "$($MyInvocation.MyCommand.Name)::End"
    }
    
}

Export-ModuleMember Send-Mail
Set-Alias sm Send-Mail
Export-ModuleMember -Alias sm