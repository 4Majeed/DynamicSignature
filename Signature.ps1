$User = $env:UserName
$FileName = "signature"
$FileExtension = "htm","txt"
$Path = "<server share>\Signature"
$PathSignature = "C:\users\" + $user + "\TempSignature"
$PathSignatureTemplates = "$Path"
$AppSignatures =$env:APPDATA + "\Microsoft\Signatures"

$DisplayName = ([adsisearcher]"(&(objectClass=user)(samaccountname=$user))").FindOne().Properties['displayname']
$Title = ([adsisearcher]"(&(objectClass=user)(samaccountname=$user))").FindOne().Properties['title']
$Email = ([adsisearcher]"(&(objectClass=user)(samaccountname=$user))").FindOne().Properties['mail']
$Extension = ([adsisearcher]"(&(objectClass=user)(samaccountname=$user))").FindOne().Properties['ipphone']

New-Item -Path "$PathSignature" -ItemType Container –Force
foreach ($Ext in $FileExtension)
{
Copy-Item -Force "$Path\$FileName.$Ext" "$PathSignature\$FileName.$Ext"
}
Copy-Item -Force "$Path\$FileName.rtf" "$PathSignature\$FileName.rtf"

foreach ($Ext in $FileExtension)
{
(Get-Content "$PathSignature\$FileName.$Ext") | Foreach-Object {
$_`
-replace '@Name', $DisplayName `
-replace '@Title', $Title `
-replace '@Ext', $Extension `
-replace '@Email', $Email `
} | Set-Content "$PathSignature\$FileName.$Ext" -encoding utf8
}

(Get-Content "$PathSignature\$FileName.rtf") | Foreach-Object {
$_`
-replace '@Name', $DisplayName `
-replace '@Title', $Title `
-replace '@Ext', $Extension `
-replace '@Email', $Email `
} | Set-Content "$PathSignature\$FileName.$Ext"

foreach ($Ext in $FileExtension)
{
Copy-Item -Force "$PathSignature\$FileName.$Ext" "$AppSignatures\$User.$Ext"
write-host "$PathSignature\$FileName.$Ext"
write-host "$AppSignatures\$User.$Ext"
}

Copy-Item -Force "$PathSignature\$FileName.rtf" "$AppSignatures\$User.rtf"
write-host "$PathSignature\$FileName.rtf"
write-host "$AppSignatures\$User.rtf"

copy-item -force "$Path\Signature_files" "$AppSignatures\Signature_files" -recurse
write-host "$AppSignatures\Signature_files"


If (Test-Path HKCU:'\Software\Microsoft\Office\15.0') {
Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Setup -Name First-Run -Force -ErrorAction SilentlyContinue -Verbose
New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'ReplySignature' -Value $User -PropertyType 'String' -Force
New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'NewSignature' -Value $User -PropertyType 'String' -Force
}

If (Test-Path HKCU:'\Software\Microsoft\Office\16.0') {
Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup -Name First-Run -Force -ErrorAction SilentlyContinue -Verbose
New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $User -PropertyType 'String' -Force
New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $User -PropertyType 'String' -Force
}
