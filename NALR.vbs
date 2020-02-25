Dim adoconn
Dim rs
Dim str
Dim adata 'Big list of all permissions
dim FoundAll, FoundChange, FoundAdd, FoundDel, FoundAny 'List of permissions that have changed
dim datadelstr 'temp string for data deletion checking
dim i, j 'Counters
Dim outputl 'Output to email
Dim EmailRecpt 'Email Recipients
Set WShell = CreateObject("wscript.shell")

''''''Edit These''''''

EmailRecpt = "admin@company.com" 'Who to send to (comma separated)
strTargetPath = "D:\" 'Drive or path to audit
'Search for "ignored" for ignored folders

''''''End if editable''''''

'On Error Resume Next
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
'Common Permissions
Const FOLDER_FULL_CONTROL = 2032127
Const FOLDER_MODIFY = 1245631
Const FOLDER_READ_ONLY = 1179785
Const FOLDER_READ_CONTENT_EXECUTE =  1179817 
Const FOLDER_READ_CONTENT_EXECUTE_WRITE =  1180095 
Const FOLDER_WRITE = 1179926
Const FOLDER_READ_WRITE = 1180063 
'Special Permissions
Const FOLDER_LIST_DIRECTORY = 1
Const FOLDER_ADD_FILE = 2
Const FOLDER_ADD_SUBDIRECTORY = 4
Const FOLDER_READ_EA = 8
Const FOLDER_WRITE_EA = 16
Const FOLDER_EXECUTE = 32
Const FOLDER_DELETE_CHILD = 64
Const FOLDER_READ_ATTRIBUTES = 128
Const FOLDER_WRITE_ATTRIBUTES = 256
Const FOLDER_DELETE = 65536
Const FOLDER_READ_CONTROL = 131072
Const FOLDER_WRITE_DAC = 262144
Const FOLDER_WRITE_OWNER = 524288
Const FOLDER_SYNCHRONIZE = 1048576
'INHERIT
'Const FOLDER_OBJECT_INHERIT_ACE = 1
'Const FOLDER_CONTAINER_INHERIT_ACE = 2
'Const FOLDER_NO_PROPAGATE_INHERIT_ACE = 4
'Const FOLDER_INHERIT_ONLY_ACE = 8
Const FOLDER_INHERITED_ACE = 16
'ACL Control
Const SE_DACL_PRESENT = 4
Const ACCESS_ALLOWED_ACE_TYPE = 0
Const ACCESS_DENIED_ACE_TYPE  = 1

strComputer = "."
adata = ""
FoundChange = ""
FoundAny = 0

If WScript.Arguments.Count = 3 Then
	strTargetPath = WScript.Arguments.Item(0)
	strOutFile = WScript.Arguments.Item(1)
	strdrop = WScript.Arguments.Item(2)
Elseif WScript.Arguments.Count = 2 Then
	strTargetPath = WScript.Arguments.Item(0)
	strOutFile = WScript.Arguments.Item(1)
	strdrop = ""
Else
	'wscript.echo "Run at CMD Prompt: cscript List_Sec_Folder_v2.vbs c:\PastaTeste Outlog.txt"
	'wscript.echo "To drop Inheritance run: cscript List_Sec_Folder_v2.vbs c:\PastaTeste Outlog.txt /dropInherit"
	'wscript.quit
	'strTargetPath = "C:\largefolder\"
	strOutFile = "LFPerm.txt"
	strdrop = "/dropInherit"
End If

If Trim(strTargetPath) = "" or Trim(strOutFile) = "" Then
	wscript.echo "Run at CMD Prompt: cscript List_Sec_Folder_v2.vbs c:\PastaTeste Outlog.txt"
	wscript.echo "To drop Inheritance run: cscript List_Sec_Folder_v2.vbs c:\PastaTeste Outlog.txt /dropInherit"
	wscript.quit
End If 

'Wscript.echo "Start Process"
'Wscript.echo "Root Folder Target: " & strTargetPath

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutFile = objFSO.OpenTextFile(strOutFile, ForWriting, True)
objOutFile.Writeline "Date;Time;Folder;Group / User Name;Common Permission's;Special Permission's;Access Type;Inheritance;Error's"

Set adoconn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
'adoconn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\LargeFolder\NewMods Test files\Access Level Review\NALR.accdb;"
adoconn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=NALR.accdb;"

ShowSubACL objFSO.GetFolder(strTargetPath)
ShowSubfolders objFSO.GetFolder(strTargetPath)

'Wscript.echo "Finished Process"
objOutFile.Close

'Check existing entries in database
str = "select * from DataDump;"
rs.Open str, adoconn, 3, 3 'OpenType, LockType

if not rs.eof then 
  rs.MoveFirst
  do while not rs.eof
    datadelstr = """" & rs("Folder") & """;" & """" & rs("GrouporName") & """;" & rs("CommonPerm") & ";"
    if instr(1,adata,datadelstr,1) = 0 then 'Add code to check to make sure folder was indeed scanned
	  FoundChange = FoundChange & rs("Folder") & "|" & "Removed" & "|" & rs("GrouporName") & "|" & rs("CommonPerm") & "|" & "None|"
	  rs.delete
	end if
	rs.movenext
  loop
end if
rs.close

'Write Change Log
do while len(FoundChange) > 0
  str = "select * from ChangeLog where 0 = 1;"
  rs.Open str, adoconn, 3, 3 'OpenType, LockType
  rs.AddNew 1, left(FoundChange,instr(1,FoundChange,"|")-1)
  FoundChange = mid(FoundChange,instr(1,FoundChange,"|")+1,len(FoundChange)-instr(1,FoundChange,"|"))
  rs("TypeChange") = left(FoundChange,instr(1,FoundChange,"|")-1)
  FoundChange = mid(FoundChange,instr(1,FoundChange,"|")+1,len(FoundChange)-instr(1,FoundChange,"|"))
  rs("UserPerm") = left(FoundChange,instr(1,FoundChange,"|")-1)
  FoundChange = mid(FoundChange,instr(1,FoundChange,"|")+1,len(FoundChange)-instr(1,FoundChange,"|"))
  rs("PrevPerm") = left(FoundChange,instr(1,FoundChange,"|")-1)
  FoundChange = mid(FoundChange,instr(1,FoundChange,"|")+1,len(FoundChange)-instr(1,FoundChange,"|"))
  rs("NewPerm") = left(FoundChange,instr(1,FoundChange,"|")-1)
  FoundChange = mid(FoundChange,instr(1,FoundChange,"|")+1,len(FoundChange)-instr(1,FoundChange,"|"))
  rs("WhoChange") = ""
  rs("AckChange") = ""
  rs("CommentChange") = ""
  rs("DateChanged") = Date()

  rs.Update
  rs.close

  FoundAny = 1
Loop
  
if FoundAny = 0 then
  FoundAll = "There were no permission changes since the last scan."
else
  'Header Info
  FoundAll = FoundAll & "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body> <table>" & vbcrlf
  FoundAll = FoundAll & "<tr>" & vbcrlf
  FoundAll = FoundAll & "  <th>Folder</th>" & vbcrlf
  FoundAll = FoundAll & "  <th>Type of Change</th>" & vbcrlf
  FoundAll = FoundAll & "  <th>User Affected</th>" & vbcrlf
  FoundAll = FoundAll & "  <th>Previous Permission</th>" & vbcrlf
  FoundAll = FoundAll & "  <th>New Permission</th>" & vbcrlf
  FoundAll = FoundAll & "  <th>Date Changed</th>" & vbcrlf
  FoundAll = FoundAll & "</tr>" & vbcrlf
  
  str = "select * from ChangeLog where DateChanged=Date()"
  rs.Open str, adoconn, 2, 1 'OpenType, LockType
  
  if not rs.eof then 
    rs.MoveFirst
	do while not rs.eof
	  FoundAll = FoundAll & "<tr>" & vbcrlf
      FoundAll = FoundAll & "  <td>" & rs("Folder") & "</td>" & vbcrlf
	  FoundAll = FoundAll & "  <td>" & rs("TypeChange") & "</td>" & vbcrlf
	  FoundAll = FoundAll & "  <td>" & rs("UserPerm") & "</td>" & vbcrlf
	  FoundAll = FoundAll & "  <td>" & rs("PrevPerm") & "</td>" & vbcrlf
	  FoundAll = FoundAll & "  <td>" & rs("NewPerm") & "</td>" & vbcrlf
	  FoundAll = FoundAll & "  <td>" & rs("DateChanged") & "</td>" & vbcrlf	  
      FoundAll = FoundAll & "</tr>" & vbcrlf
	  rs.MoveNext
    loop
  end if
  rs.close
end if
FoundAll = FoundAll & "</table>"
FoundAll = FoundAll & "<br><a href=""http://www.intranet.com/mis-only/nalr/index.asp?filter=Search&StartDate=" & format(date(), "YYYYMMDD") & "&EndDate=" & format(date(), "YYYYMMDD") & """>Acknowledge Changes</a>"
FoundAll = FoundAll & "</body></html>"
WriteFile "PermChanges.txt", FoundAll

outputl = FoundAll
SendMail EmailRecpt, "NALR: Daily Report"

Set adoconn = Nothing
Set rs = Nothing

Sub ShowSubFolders(Folder)
	On Error Resume Next
    For Each Subfolder in Folder.SubFolders
	  if not UCase(Subfolder.Path) = UCase("D:\Home\z_Profiles") then 'Ignored list
		'Wscript.Echo Subfolder.Path
		'Wscript.echo "Get ACL on Path: " & Subfolder.Path
		ShowSubACL(Subfolder.Path)
		ShowSubFolders Subfolder
		If Err.Number = 0 Then
			strErros = "No Error's"
		ElseIf Err.Number = 451 Then
			strErros = "No Error's"
			Err.clear
		Else
			strErros = "Cod.: " & Err.Number & " Desc.: " & Err.description
			objOutFile.Writeline Date() & ";" & Time() & ";" & FolderPerm & ";" & "" & "" & "" & ";" & "" & ";" & "" & ";" & ""  & ";" & "" & ";" & strErros
			Err.clear
		End If
	  end if
    Next
End Sub

Sub ShowSubACL(FolderPerm)
	On Error Resume Next
	strCPerm = ""
	strSPerm = ""
	strTypePerm = "" 
	strInherit = ""
	strErros = ""
	Set objWMIService = GetObject("winmgmts:")
	Set objFolderSecuritySettings = objWMIService.Get("Win32_LogicalFileSecuritySetting='" & FolderPerm & "'")
	intRetVal = objFolderSecuritySettings.GetSecurityDescriptor(objSD)
	intControlFlags = objSD.ControlFlags
	If intControlFlags AND SE_DACL_PRESENT Then
		arrACEs = objSD.DACL
		If Err.Number = 0 Then
			strErros = "No Error's"
		Else
			strErros = "Cod.: " & Err.Number & " Desc.: " & Err.description
			objOutFile.Writeline Date() & ";" & Time() & ";""" & FolderPerm  & """;" & "" & "" & "" & ";" & "" & ";" & "" & ";" & ""  & ";" & "" & ";" & strErros
			Err.clear
		End If
		For Each objACE in arrACEs
			
			'ACL Type
			If objACE.AceType = ACCESS_ALLOWED_ACE_TYPE Then strTypePerm = "Allowed"
			If objACE.AceType = ACCESS_DENIED_ACE_TYPE Then strTypePerm = "Denied"
			
			'Inherit
			If objAce.AceFlags AND FOLDER_INHERITED_ACE Then
				strInherit = "Yes"
			Else
				strInherit = "No"
			End if
			
			'Common Permissions
			If objACE.AccessMask = FOLDER_FULL_CONTROL Then 
				strCPerm = "Full Control"
			ElseIf objACE.AccessMask = FOLDER_MODIFY Then 
				strCPerm = "Modify"
			ElseIf objACE.AccessMask = FOLDER_READ_CONTENT_EXECUTE_WRITE Then
				strCPerm = "Read & Execute / List Folder Contents (folders only) + Write"
			ElseIf objACE.AccessMask = FOLDER_READ_CONTENT_EXECUTE Then 
				strCPerm = "Read & Execute / List Folder Contents (folders only)"
			ElseIf objACE.AccessMask = FOLDER_READ_WRITE Then
				strCPerm = "Read + Write"
			ElseIf objACE.AccessMask = FOLDER_READ_ONLY Then 
				strCPerm = "Read Only"
			ElseIf objACE.AccessMask = FOLDER_WRITE Then 
				strCPerm = "Write"
			Else
				strCPerm = "Special"
			End If
			
			'Special Permissions
			strSPerm = ""
			If objACE.AccessMask and FOLDER_EXECUTE Then strSPerm = strSPerm & "Traverse Folder/Execute File, "
			If objACE.AccessMask and FOLDER_LIST_DIRECTORY Then strSPerm = strSPerm & "List Folder/Read Data, "
			If objACE.AccessMask and FOLDER_READ_ATTRIBUTES Then strSPerm = strSPerm & "Read Attributes, "
			If objACE.AccessMask and FOLDER_READ_EA Then strSPerm = strSPerm & "Read Extended Attributes, "
			If objACE.AccessMask and FOLDER_ADD_FILE Then strSPerm = strSPerm & "Create Files/Write Data, "
			If objACE.AccessMask and FOLDER_ADD_SUBDIRECTORY Then strSPerm = strSPerm & "Create Folders/Append Data"
			If objACE.AccessMask and FOLDER_WRITE_ATTRIBUTES Then strSPerm = strSPerm & "Write Attributes, "
			If objACE.AccessMask and FOLDER_WRITE_EA Then strSPerm = strSPerm & "Write Extended Attributes, "
			If objACE.AccessMask and FOLDER_DELETE_CHILD Then strSPerm = strSPerm & "Delete Subfolders and Files, "
			If objACE.AccessMask and FOLDER_DELETE Then strSPerm = strSPerm & "Delete, "
			If objACE.AccessMask and FOLDER_READ_CONTROL Then strSPerm = strSPerm & "Read Permissions, "
			If objACE.AccessMask and FOLDER_WRITE_DAC Then strSPerm = strSPerm & "Change Permissions, "
			If objACE.AccessMask and FOLDER_WRITE_OWNER Then strSPerm = strSPerm & "Take Ownership, "
			If objACE.AccessMask and FOLDER_SYNCHRONIZE Then strSPerm = strSPerm & "Synchronize, "
			If trim(strSPerm) <> "" then strSPerm =  left(strSPerm, len(strSPerm)-2)

			If UCase(strdrop) = UCase("/dropInherit") and objAce.AceFlags AND FOLDER_INHERITED_ACE and not UCase(FolderPerm) = UCase(strTargetPath) Then
			  'Wscript.echo "Dropped ACL Inheritance " & FolderPerm
			Else	
				'Wscript.echo "Get ACL on Path: " & FolderPerm
				'Date;Time;Folder;Group / User Name;Common Permission's;Special Permission's;Access Type;Inherit;Error's
				objOutFile.Writeline Date() & ";" & Time() & ";""" & FolderPerm & """;""" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name & """;" & strCPerm & ";" & strSPerm & ";" & strTypePerm  & ";" & strInherit & ";" & strErros
				'wscript.echo objACE.Trustee.Name & " " & objACE.AceFlags
				
				'Add to database
				str = "select * from DataDump where Folder='" & FolderPerm & "' and GrouporName='" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name & "';"
                rs.Open str, adoconn, 3, 3 'OpenType, LockType
				'msgbox "first " & FolderPerm & " '" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name & "'"
				i = 0
				if not rs.eof then 
                  rs.MoveFirst
				  do while not rs.eof
				    'msgbox "found " & rs("GrouporName") & " = " & objACE.Trustee.Domain & "\" & objACE.Trustee.Name
				    if rs("CommonPerm") = strCPerm then
					  if not rs("SpecialPerm") = strSPerm then 'Minor
						FoundChange = FolderPerm & "|" & "Special" & "|" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name & "|" & rs("SpecialPerm") & "|" & strSPerm
						rs("SpecialPerm") = strSPerm
					  end if
					  if not rs("AccessType") = strTypePerm then 'Major
					    FoundChange = FolderPerm & "|" & "Access Type" & "|" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name & "|" & rs("AccessType") & "|" & strTypePerm
					    rs("AccessType") = strTypePerm
					  end if
					  if not rs("Inheritance") = strInherit then 'Should never happen
					    FoundChange = FolderPerm & "|" & "Access Type" & "|" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name & "|" & rs("Inheritance") & "|" & strInherit
						rs("Inheritance") = strInherit
					  end if
					  
					  'See if it is the top folder
					  if instr(1,adata,FolderPerm,1) = 0 then
					    rs("TopFolder") = True
					  else
					    rs("TopFolder") = False
					  end if
					  
					  rs.Update
					  i = 1
					end if
					rs.movenext
				  loop
				  'msgbox "End of loop"
				end if
				if i = 0 then
				  'msgbox "add new"
				  rs.AddNew 1, FolderPerm
				  rs("GrouporName") = objACE.Trustee.Domain & "\" & objACE.Trustee.Name
				  rs("CommonPerm") = strCPerm
                  rs("SpecialPerm") = strSPerm
				  rs("AccessType") = strTypePerm
				  rs("Inheritance") = strInherit
				  'See if it is the top folder
				  if instr(1,adata,FolderPerm,1) = 0 then
				    rs("TopFolder") = True
				  else
					rs("TopFolder") = False
				  end if
      	            
	              rs.Update
				  FoundChange = FolderPerm & "|" & "New Access" & "|" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name & "|" & "None" & "|" & strCPerm
				end if
				'msgbox "all done"
				rs.close
				adata = adata & Date() & ";" & Time() & ";""" & FolderPerm & """;""" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name & """;" & strCPerm & ";" & strSPerm & ";" & strTypePerm  & ";" & strInherit & ";" & strErros & vbcrlf
				
				'Write changes
				if len(FoundChange) > 0 then
				  str = "select * from ChangeLog where 0 = 1;"
                  rs.Open str, adoconn, 3, 3 'OpenType, LockType
				  rs.AddNew 1, left(FoundChange,instr(1,FoundChange,"|")-1)
				  FoundChange = mid(FoundChange,instr(1,FoundChange,"|")+1,len(FoundChange)-instr(1,FoundChange,"|"))
				  rs("TypeChange") = left(FoundChange,instr(1,FoundChange,"|")-1)
				  FoundChange = mid(FoundChange,instr(1,FoundChange,"|")+1,len(FoundChange)-instr(1,FoundChange,"|"))
				  rs("UserPerm") = left(FoundChange,instr(1,FoundChange,"|")-1)
				  FoundChange = mid(FoundChange,instr(1,FoundChange,"|")+1,len(FoundChange)-instr(1,FoundChange,"|"))
                  rs("PrevPerm") = left(FoundChange,instr(1,FoundChange,"|")-1)
				  FoundChange = mid(FoundChange,instr(1,FoundChange,"|")+1,len(FoundChange)-instr(1,FoundChange,"|"))
				  rs("NewPerm") = FoundChange
				  rs("WhoChange") = ""
				  rs("AckChange") = ""
				  rs("CommentChange") = ""
				  rs("DateChanged") = Date()
				  
				  rs.Update
				  rs.close
				  
				  FoundChange = ""
				  FoundAny = 1
				end if
			End if	
		Next
	End If
End Sub

'Write string As a text file.
function WriteFile(FileName, Contents)
  Dim OutStream, FS

  on error resume Next
  Set FS = CreateObject("Scripting.FileSystemObject")
    Set OutStream = FS.OpenTextFile(FileName, 2, True)
    OutStream.Write Contents
End Function

Function SendMail(TextRcv,TextSubject)
  Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
  Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 

  Const cdoAnonymous = 0 'Do not authenticate
  Const cdoBasic = 1 'basic (clear-text) authentication
  Const cdoNTLM = 2 'NTLM

  Set objMessage = CreateObject("CDO.Message") 
  objMessage.Subject = TextSubject 
  objMessage.From = "admin@company.com" 
  objMessage.To = TextRcv
  objMessage.HTMLBody = outputl

  '==This section provides the configuration information for the remote SMTP server.

  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

  'Name or IP of Remote SMTP Server
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "webmail.company.com"

  'Type of authentication, NONE, Basic (Base64 encoded), NTLM
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoAnonymous

  'Server port (typically 25)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

  'Use SSL for the connection (False or True)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

  'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

  objMessage.Configuration.Fields.Update

  '==End remote SMTP server configuration section==

  objMessage.Send
End Function

Function Format(vExpression, sFormat)
  Dim nExpression
  nExpression = sFormat
  
  if isnull(vExpression) = False then
    if instr(1,sFormat,"Y") > 0 or instr(1,sFormat,"M") > 0 or instr(1,sFormat,"D") > 0 or instr(1,sFormat,"H") > 0 or instr(1,sFormat,"S") > 0 then 'Time/Date Format
      vExpression = cdate(vExpression)
	  if instr(1,sFormat,"AM/PM") > 0 and int(hour(vExpression)) > 12 then
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression)-12,2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)-12) '1 character hour
		nExpression = replace(nExpression,"AM/PM","PM") 'If if its afternoon, its PM
	  else
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression),2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)) '1 character hour
		if int(hour(vExpression)) = 12 then nExpression = replace(nExpression,"AM/PM","PM") '12 noon is PM while anything else in this section is AM (fixed 04/19/2019 thanks to our HR Dept.)
		nExpression = replace(nExpression,"AM/PM","AM") 'If its not PM, its AM
	  end if
	  nExpression = replace(nExpression,":MM",":" & right("00" & minute(vExpression),2)) '2 character minute
	  nExpression = replace(nExpression,"SS",right("00" & second(vExpression),2)) '2 character second
	  nExpression = replace(nExpression,"YYYY",year(vExpression)) '4 character year
	  nExpression = replace(nExpression,"YY",right(year(vExpression),2)) '2 character year
	  nExpression = replace(nExpression,"DD",right("00" & day(vExpression),2)) '2 character day
	  nExpression = replace(nExpression,"D",day(vExpression)) '(N)N format day
	  nExpression = replace(nExpression,"MMM",left(MonthName(month(vExpression)),3)) '3 character month name
	  if instr(1,sFormat,"MM") > 0 then
	    nExpression = replace(nExpression,"MM",right("00" & month(vExpression),2)) '2 character month
	  else
	    nExpression = replace(nExpression,"M",month(vExpression)) '(N)N format month
	  end if
    elseif instr(1,sFormat,"N") > 0 then 'Number format
	  nExpression = vExpression
	  if instr(1,sFormat,".") > 0 then 'Decimal format
	    if instr(1,nExpression,".") > 0 then 'Both have decimals
		  do while instr(1,sFormat,".") > instr(1,nExpression,".")
		    nExpression = "0" & nExpression
		  loop
		  if len(nExpression)-instr(1,nExpression,".") >= len(sFormat)-instr(1,sFormat,".") then
		    nExpression = left(nExpression,instr(1,nExpression,".")+len(sFormat)-instr(1,sFormat,"."))
	      else
		    do while len(nExpression)-instr(1,nExpression,".") < len(sFormat)-instr(1,sFormat,".")
			  nExpression = nExpression & "0"
			loop
	      end if
		else
		  nExpression = nExpression & "."
		  do while len(nExpression) < len(sFormat)
			nExpression = nExpression & "0"
		  loop
	    end if
	  else
		do while len(nExpression) < sFormat
		  nExpression = "0" and nExpression
		loop
	  end if
	else
      msgbox "Formating issue on page. Unrecognized format: " & sFormat
	end if
	
	Format = nExpression
  end if
End Function