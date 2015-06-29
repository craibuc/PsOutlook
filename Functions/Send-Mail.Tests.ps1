$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".")
. "$here\$sut"

Describe "Send-Mail" {

    $p0 = "psoutlook_0@mailinator.com"
    $p1 = "psoutlook_1@mailinator.com"
    $p2 = "psoutlook_2@mailinator.com"

    $plainText = "PsOutlook"
    $htmlText = "<h1>PsOutlook</h1>"

    $file0 = New-Item "TestDrive:\Desktop\File0.txt" -Type File -Force
    Set-Content $file0 -value "AbCdEfGhIjKlMnOpQrStUvWxYz"

    $file1 = New-Item "TestDrive:\Desktop\File1.txt" -Type File -Force
    Set-Content $file1 -value "aBcDeFgHiJkLmNoPqRsTuVwXyZ"

    It "Fails without a To and Subject" {
        { Send-Mail } | Should Throw
    }  

    # It "Sends a plain-text message" {
    #     Send-Mail $p0 "Sends a plain-text message" $plainText -preview
    # }

    # It "Sends a HTML message" {
    #     Send-Mail $p0 "Sends a HTML message" $htmlText -preview
    # }

    # It "Sends a message with attachments" {
    #     Send-Mail $p0 "Sends a message with attachments"  $plainText -Attachments @($file0,$file1) -preview
    # }

    # It "Sends a plain-text message to multiple recipients" {
    #     Send-Mail @($p0,$p1,$p2) "Sends a plain-text message to multiple recipients" $plainText -preview
    # }

    # It "Sends a plain-text message to multiple recipient types" {
    #     Send-Mail $p0 "Sends a plain-text message to multiple recipient types" $plainText -cc $p1 -bcc $p2 -preview
    # }

}

Describe "Aliases" {

    It "Send-Mail alias should exist" {
        (Get-Alias -Definition Send-Mail).name | Should Be "sm"
    }

}