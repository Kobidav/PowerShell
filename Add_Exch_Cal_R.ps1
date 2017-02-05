
function Write-Color([String[]]$Text, [ConsoleColor[]]$Color) {
    for ($i = 0; $i -lt $Text.Length; $i++) {
        Write-Host $Text[$i] -Foreground $Color[$i] -NoNewLine
    }
    Write-Host
}


#****************************************************************

$GroupToAdd = "Calendar_sec_rights"
$GroupToCheck = "Calendar_adv_rights"
$Rights = "Editor"
#***********************************************************

$Mb = Get-mailbox
$Mb | Foreach { $MbAlias= $_.Alias

$ADAlias = switch -Wildcard ($MbAlias)
    { "DiscoverySearchMailbox{D919BA05-46A6-415f-80AD-7E09334BB852}" {"T43"}
      default     {$MbAlias} 
      }

Get-ADPrincipalGroupMembership $AdAlias | foreach  {$UGroup = $_.Name
            if ($UGroup -eq $GroupToCheck)
                 {Write-host "**********************************" 

                
                 Write-color -Text "Found user ", $AdAlias, " In Group ", $UGroup, ". Check Folder Rights" -Color Gray, Yellow, Gray, Yellow, Gray
                 Write-host



$flag=0
Get-MailboxFolderPermission -Identity ((Get-MailboxFolderstatistics -Identity $MbAlias | Where-Object {$_.FolderType -eq "Calendar"}| Select -exp Identity) -replace "\\", ":\") | Select User | foreach { $n = $_.User

                 if ([string]$n -eq $GroupToAdd )
        {
        Write-host "Found rights"  $n ([string]$m) "in Calendar of" $MbAlias "Skip"
        Write-host "**********************************"
        $flag = 1
        }


                   }
 if ($flag -eq 0)
 { Write-color -text  "Rights ", $GroupToAdd,  " not Find in Calendar of ", $MbAlias, ". Adding..." -Color Gray, Yellow, Gray, Yellow, Gray
 Add-MailboxFolderPermission -Identity ((Get-MailboxFolderstatistics -Identity $mbalias| Where-Object {$_.FolderType -eq "Calendar"}| Select -exp Identity) -replace "\\", ":\") -User $GroupToAdd -accessRights $Rights
 Write-host "**********************************" 
 Write-host}}}}

Write-host " Start Checking...."
$Mb = Get-mailbox
$Mb | Foreach { $MbAlias= $_.Alias

$FolderPath = (Get-MailboxFolderstatistics -Identity $MbAlias | Where-Object {$_.FolderType -eq "Calendar"}| Select -Exp Identity) -replace "\\", ":\"
Get-MailboxFolderPermission -Identity $FolderPath | Select User | foreach { $n = $_.User

                 if ([string]$n -eq $GroupToAdd )
        {
        Write-host "**********************************"
        Write-color -Text "Found rights ", $GroupToAdd, " in Calendar of ", $MbAlias, ".Checking Folder..." -Color Gray, Yellow, Gray, Yellow, Gray
        Write-host 
        $ADAlias = switch -Wildcard ($MbAlias)
    { "DiscoverySearchMailbox{D919BA05-46A6-415f-80AD-7E09334BB852}" {"T43"}
      default     {$MbAlias} 
      }
$flag =0
 
Get-ADPrincipalGroupMembership $AdAlias | foreach  {$UGroup = $_.Name

            if ($UGroup -eq $GroupToCheck)
                 {
                  
                 Write-color -Text "Found user ", $AdAlias, " In Group ", $UGroup, ". All Rights" -Color Gray, Yellow, Gray, Yellow, Gray
                 Write-host "**********************************" 
                 $flag = 1
                 
                 
 
                 }}                             
 if ($flag -eq 0)
 { Write-color -text  "User ", $GroupToAdd,  " not Find in Group ", $GroupToCheck, ". Remove..." -Color Gray, Yellow, Gray, Yellow, Gray
 Remove-MailboxFolderPermission -Identity ([string]$FolderPath) -User $GroupToAdd 
 Remove-MailboxFolderPermission -Identity ([string]$FolderPath) -User $GroupToAdd 
 Write-host "**********************************" 
 Write-host}    
                 

                 }}}


