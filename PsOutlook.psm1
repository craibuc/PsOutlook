Write-Host "Importing module PsOutlook..."

#
# load (dot-source) *.PS1 files, excluding unit-test scripts (*.Tests.*), and disabled scripts (__*)
#
Get-ChildItem "$PSScriptRoot\Functions\*.ps1" | 
    Where-Object { $_.Name -like '*.ps1' -and $_.Name -notlike '__*' -and $_.Name -notlike '*.Tests*' } | 
    % { . $_ }

Export-ModuleMember Send-Mail
Export-ModuleMember -Alias sm