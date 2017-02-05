param (
    [string]$alias = $( Read-Host "Input alias name" ),
    [string]$user = $( Read-Host "Input user name" ),
    [string]$rights = "Editor",
    [switch]$remove = $false
 )


Add-MailboxFolderPermission -Identity ((Get-MailboxFolderstatistics -Identity $alias| Where-Object {$_.FolderType -eq "Calendar"}| Select -exp Identity) -replace "\\", ":\") -User $user -accessRights $rights
if ($remove) {
Remove-MailboxFolderPermission -Identity ((Get-MailboxFolderstatistics -Identity $alias| Where-Object {$_.FolderType -eq "Calendar"}| Select -exp Identity) -replace "\\", ":\") -User $user 
}