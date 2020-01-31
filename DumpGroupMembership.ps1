#import-module ActiveDirectory

$Groups = (Get-AdGroup -filter * | Where {$_.name -like "**"} | select name -expandproperty name)

$Table = @()

$Record = [ordered]@{
"Group Name" = ""
#"Name" = ""
"Username" = ""
}

Foreach ($Group in $Groups)
{
if ($Group -ne "Domain Computers")
{
$Arrayofmembers = Get-ADGroupMember -identity $Group -Recursive | select samaccountname

foreach ($Member in $Arrayofmembers)
{
$Record."Group Name" = $Group
#$Record."Name" = $Member.name
$Record."UserName" = $Member.samaccountname
$objRecord = New-Object PSObject -property $Record
$Table += $objrecord

}
}
}

$Table | export-csv "C:\NALR\SecurityGroups.csv" -NoTypeInformation

cscript.exe C:\NALR\ImportGroups.vbs