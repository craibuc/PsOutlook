# PsOutlook

PowerShell wrapper of selected Microsoft Outlook functionality.

## Installation

* Download the [latest version of PsOutlook](https://github.com/craibuc/PsOutlook/releases)
* Unzip
* Copy the PsOutlook folder to `C:\Users\<account>\Documents\WindowsPowerShell\Modules`
* Add `Import-Module PsOutlook -Force` to script

## Usage

```powershell
# send the message without interaction
PS> Send-Mail "recipient@domain0.tld;recipient@domain1.tld", "the subject", "the message", "path/to/attachment0;path/to/attachmentN"
```

```powershell
# show the message
PS> Send-Mail "recipient@domain0.tld;recipient@domain1.tld", "the subject", "the message", "path/to/attachment0;path/to/attachmentN" -Preview
```

## Contributors

* [Craig Buchanan](https://github.com/craibuc) - Author
