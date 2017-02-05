Function AddACL($Path){
$Acl = Get-Acl $Path
$Ar = New-Object  system.security.accesscontrol.filesystemaccessrule("Domain Users","FullControl","Allow")
$Acl.SetAccessRule($Ar)
Set-Acl $Path $Acl
}



$num = 1
$commit_folder = "C:\Bincopy\ads.ini"
#$commit_shortcut = "E:\commitsys\COMMIT\Bincopy\CommitStart.exe.lnk"
$d = [DateTime]::Today.AddDays(-60) 
$AllComp = (Get-ADComputer -Filter '(LastLogonDate -gt $d) -and ((OperatingSystem -eq "Windows 7 Professional") -or (OperatingSystem -like "*XP*") -or (OperatingSystem -like "*10*"))'  -Properties OperatingSystem, LastLogonDate | select name)
Write-Host "List of active computers gotten" -ForeGroundColor Green 
$AllComp | foreach { $a=$_.name
$c = "\\" + $a + "\c$\Bincopy\"
#$d = "\\" + $a + "\c$\Users\Public\Desktop"
if (Test-Path ($c + "ads.ini")) {
$Exit = [string]$num + ". Comp " + $a +  ": file ads.ini exist"}
Else {
Copy-Item $commit_folder $c
#Copy-Item $commit_shortcut $d
#AddAcl ($c + "\Bincopy")
#AddAcl ($d + "\CommitStart.exe.lnk")
$exit = [string]$num + ". Comp " +$a + ": ads.ini coped"}

Write-Host $Exit -ForeGroundColor Green
$num = $num +1
}