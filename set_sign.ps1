$Template_folder = "E:\IT\scripts\PS\set_sign\"


Function Get_Table([string]$Group_Name, [string]$Adv, [string]$Partner)
{
$Table = (Get-ADGroupMember  $Group_Name | foreach { Get-Aduser $_.SamAccountName  -Properties mail | select Name, mail, SamAccountName, @{Label="Group"; Expression={$Group_Name}}})

$Table | foreach { 
 $Full_Name=$_.Name
 $Mail = $_.mail
 
 
 
$AccName = $_.SamAccountName

Write-Host $Full_Name, $Mail, $AccName, $Group_Name , $Adv , $Partner

New-Item ($Template_folder + $AccName) -type directory -Force

(Get-Content ($Template_folder + 'sign_template\t_e_m_p_l_a_t_e.htm')) -replace('t_e_m_p_l_a_t_e', $AccName) -replace('t_e_m_p_n_a_m_e', $Full_Name) -replace('t_e_m_p_a_d_v', $Adv) -replace('t_e_m_p_e_m_a_i_l', $Mail) -replace('t_e_m_p_a_r_t_n_e_r', $Partner) | Set-Content ($Template_folder + $AccName + "\" + $AccName + ".htm")
New-Item ($Template_folder + $AccName + "\" + $AccName + "_files") -type directory -Force
Copy-Item ($Template_folder + "sign_template\t_e_m_p_l_a_t_e_files\*") -Destination ($Template_folder + $AccName + "\" + $AccName + "_files") -Force
(Get-Content ($Template_folder + $AccName + '\' + $AccName + '_files\filelist.xml')) -replace('t_e_m_p_l_a_t_e', $AccName)| Set-Content ($Template_folder + $AccName + '\' + $AccName + '_files\filelist.xml')



} # Create Sing files from AD groups

}
Function Put_Data($Put_Name,$Put_Adv, $Put_Prof, $Put_mail, $Put_sid)
{ 
Write-Host $Put_Name,$Put_Prof, $Put_mail

New-Item ($Template_folder + $Put_sid) -type directory -Force

(Get-Content ( $Template_folder + '\sign_template\t_e_m_p_l_a_t_e.htm')) -replace('t_e_m_p_l_a_t_e', $Put_sid) -replace('t_e_m_p_n_a_m_e', $Put_Name) -replace('t_e_m_p_a_d_v', $Put_Adv) -replace('t_e_m_p_e_m_a_i_l', $Put_mail) -replace('t_e_m_p_a_r_t_n_e_r', $Put_Prof) | Set-Content ($Template_folder + $Put_sid + '\' + $Put_sid + '.htm')
New-Item ($Template_folder + $Put_sid + "\" + $Put_sid + "_files") -type directory -Force
Copy-Item ($Template_folder + "sign_template\t_e_m_p_l_a_t_e_files\*") -Destination ($Template_folder + $Put_sid + "\" + $Put_sid + "_files") -Force
(Get-Content ($Template_folder + $Put_sid + '\' + $Put_sid + '_files\filelist.xml')) -replace('t_e_m_p_l_a_t_e', $Put_sid)| Set-Content ($Template_folder + $Put_sid + '\' + $Put_sid + '_files\filelist.xml')



   } # Create Sing fils manualy
Function Send_Data
{
$d = [DateTime]::Today.AddDays(-60)
Get-ADComputer -Filter '(LastLogonDate -gt $d) -and ((OperatingSystem -eq "Windows 10 Pro") -or (OperatingSystem -eq "Windows 7 Professional") -or (OperatingSystem -like "*XP*"))'  | foreach {
$CompName = $_.name
$UserName = [string]((Get-WmiObject win32_computersystem -comp $CompName).Username) -replace "RMBLAW" -replace "\\"
Write-Host $UserName

if ( Test-Path ($Template_folder + $UserName)){
if ((Test-connection $CompName -count 2 -quiet) -eq "True")
{
Write-Host "Sending to " $CompName " for " $UserName -ForeGroundColor Green 
$Remote_s_folder =("\\" + $CompName + "\C$\Users\" + $UserName + "\AppData\Roaming\Microsoft\Signatures") 
if ( Test-Path ($Remote_s_Folder + "\" + $UserName + "_files"))
 {Write-host "Add to zip"
  Compress-Archive ($Remote_s_Folder + "\" + $UserName + "_files") -CompressionLevel Optimal  ($Remote_s_Folder + "\" + $UserName + ".zip") -force
  Compress-Archive ($Remote_s_Folder + "\" + $UserName + ".htm") -CompressionLevel Optimal  ($Remote_s_Folder + "\" + $UserName + ".zip") -Update
Write-host "remove old copy"
Remove-item ($Remote_s_Folder + "\" + $UserName + "_files") -Recurse
Remove-item ($Remote_s_Folder + "\" + $UserName + ".htm")
}
  
Copy-Item -Recurse -Force ($Template_folder + $UserName + "\*") $Remote_s_folder
}
else 
{ Write-Host "No Connection" $CompName $UserName }
}
else 
{ Write-Host "No Folder" $CompName $UserName} 
 }
 } # Send Files


Get_Table 'legal clerks' ' ' 'Legal Intern'
Get_Table 'advocates' ', Adv.' ' '
Get_Table 'Partners all' ', Adv.' 'Partner'  


#Send_Data
