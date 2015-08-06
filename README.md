# PsOutlook

PowerShell wrapper of selected Microsoft Outlook functionality.

## Installation

* Download the [latest version of PsOutlook](https://github.com/craibuc/PsOutlook/releases/latest)
* Unzip
* Copy the PsOutlook folder to `C:\Users\<account>\Documents\WindowsPowerShell\Modules`
* Type `PS>  Import-Module PsOutlook -Force` in script's directory or add to `$Profile`

## Usage

```powershell
# send the message without interaction
PS> Send-Mail -To @("recipient@domain0.tld","recipient@domain1.tld") -Subject "the subject" -Body "the message" -Attachments @("path/to/attachment0","path/to/attachmentN")
```

```powershell
# show the message
PS> Send-Mail -To @("recipient@domain0.tld","recipient@domain1.tld") -Subject "the subject" -Body "the message" -Attachments @("path/to/attachment0","path/to/attachmentN") -Preview
```

```powershell
# alternative syntax
$Message = @{
  To=@("recipient@domain0.tld","recipient@domain1.tld")
  Subject="the subject"
  Body="the message" 
  Attachments=@("path/to/attachment0","path/to/attachmentN")
  Preview=$True
}

PS> Send-Mail @Message
```

## Contributors

* [Craig Buchanan](https://github.com/craibuc) - Author
