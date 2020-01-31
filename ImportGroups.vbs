Dim adoconn
Dim rs
Dim str
Dim AllMembers 'Data from CSV

Set adoconn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
adoconn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\NALR\NALR.accdb;"

'Check existing entries in database
str = "Delete * from GroupMembers;"
rs.Open str, adoconn, 3, 3 'OpenType, LockType

'rs.Close

AllMembers = GetFile("C:\NALR\SecurityGroups.csv")
AllMembers = replace(AllMembers,"""Group Name"",""UserName""","")

str = "Select * from GroupMembers where 1 = 0;"
rs.Open str, adoconn, 3, 3 'OpenType, LockType

do while len(AllMembers) > 3
  AllMembers = right(AllMembers,len(AllMembers)-3)
  rs.AddNew 0, mid(AllMembers,1,instr(1,AllMembers,"""",1)-1)
  'msgbox "'" & mid(AllMembers,1,20) & "'"
  'msgbox mid(AllMembers,1,instr(1,AllMembers,"""",1)-1)
  'msgbox instr(1,AllMembers,"""",1)
  AllMembers = right(AllMembers,len(AllMembers)-instr(1,AllMembers,"""",1)-2)
  rs("UserName") = mid(AllMembers,1,instr(1,AllMembers,"""",1)-1)
  'msgbox mid(AllMembers,1,instr(1,AllMembers,"""",1)-1)
  'msgbox instr(1,AllMembers,"""",1)
  AllMembers = right(AllMembers,len(AllMembers)-instr(1,AllMembers,"""",1))
  
  rs.update
loop

Set adoconn = Nothing
Set rs = Nothing

'Read text file
function GetFile(FileName)
  If FileName<>"" Then
    Dim FS, FileStream
    Set FS = CreateObject("Scripting.FileSystemObject")
      on error resume Next
      Set FileStream = FS.OpenTextFile(FileName)
      GetFile = FileStream.ReadAll
  End If
End Function