<html>
<head>
<title>Network Access Review</title>
<%
'Edit these for your organization
Dim DatabaseNALR

DatabaseNALR = "C:\NALR\NALR.accdb" 'Database location
AdminGrp = "MIS" 'Name of the domain admin group (group that can acknowlege changes)
%>
<link rel="stylesheet" href="dhtmlwindow.css" type="text/css" />
<script src="dhtmlwindow.js" type="text/javascript"></script>
</head>
<body bgcolor="white" text="black">
<%
' Script to get the domain and name of the logged on user.
Private strDomain
Private strUserName
Private wshNetwork

Set wshNetwork = CreateObject("WScript.Network")

strDomain = wshNetwork.UserDomain
strUserName = lcase(wshNetwork.UserName)

Set wshNetwork = Nothing
%>
<center>
Options: <a href="index.asp">Back to NALR</a>
<br><br>

<!-- Begin form code -->
<form name="form" method="post" action="investigate_results.asp">
  Folder: <input type="text" name="Folder" maxlength="50">&#160;&#160;
  
  Name (or group): <input type="text" name="Name" maxlength="50">
  <small>Enum Grps<input type="checkbox" name="Group"></small>&#160;&#160;
  
  Inject group: <input type="text" name="InjectGrp" maxlength="50">&#160;&#160;

  Permission: 
    <select name="Permission">
      <option value="Any">Any</option>
      <option value="RX">Read and Execute</option>
      <option value="Modify">Modify</option>
	  <option value="FullControl">Full Control</option>
	  <option value="Editable">Editable (Modify or Full)</option>
	  <option value="Special">Special/Other</option>
    </select>&#160;&#160;

  <input type="submit" name="Submit" value="Search">
</form>
<!-- End form code -->
</center>

<table style="font-size:14px;" BORDER="1" CELLSPACING="1" CELLPADDING="1" WIDTH="100%" bgcolor="ffffff">
<font size=2>
<tr>
  <td><center><b>Folder</b></center></td>
  <td><center><b>User Affected</center></b></td>
  <td><center><b>Permission</center></b></td>
  <td><center><b>Details</center></b></td>
</tr>

<% 
'Dimension variables
Dim adoCon         'Holds the Database Connection Object
Dim rsCheckIn    'Holds the recordset for the records in the database
Dim strSQL         'Holds the SQL query to query the database
Dim OrderSQL 'The order by field at the end of the SQL string
Dim i
Dim inRec, outRec, dateRec
dim filter

'Read in the filter string
filter = Request.QueryString("filter")
name = Request.QueryString("name")
folder = Request.QueryString("folder")
EGrp = Request.QueryString("EGrp")
InjectGrp = Request.QueryString("injectgrp")
permission = Request.QueryString("permission")
acknowleged = Request.QueryString("acknowleged")

'Create an ADO connection object
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & DatabaseNALR & ";Persist Security Info=False"

'Create an ADO recordset object
Set rsCheckIn = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "select * from DataDump where 1 = 0;"

'Filters
if filter = "Search" then

  strSQL = "select * from DataDump where "
  if EGrp = "on" and len(name)>1 then
    strSQL = strSQL & "(GrouporName = '" & strDomain & "\" & name & "'"
    rsCheckIn.Open "select * from GroupMembers where UserName = '" & name & "'", adoCon
	Do While not rsCheckIn.EOF 
	  strSQL = strSQL & " or GrouporName = '" & strDomain & "\" & rsCheckIn("GroupName") & "'"
	  rsCheckIn.MoveNext
	loop
	if len(InjectGrp) > 1 then strSQL = strSQL & " or GrouporName = '" & strDomain & "\" & InjectGrp & "'" 'To inject group
	strSQL = strSQL & ")"
	OrderSQL =  "order by Folder,GrouporName"
	rsCheckIn.Close
  elseif EGrp = "" and len(name)>1 then
    strSQL = strSQL & "(GrouporName = '" & strDomain & "\" & name & "'"
	if len(InjectGrp) > 1 then strSQL = strSQL & " or GrouporName = '" & strDomain & "\" & InjectGrp & "'" 'To inject group
	strSQL = strSQL & ")"
	OrderSQL = "order by Folder,GrouporName"
  end if
  if len(folder) > 1 then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
    strSQL = strSQL & "Folder like '%" & folder & "%'"
  end if
  if not permission = "" and not permission = "Any" then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
	if permission = "RX" then
	  strSQL = strSQL & "CommonPerm = 'Read & Execute / List Folder Contents (folders only)'"
	elseif permission = "Modify" then
	  strSQL = strSQL & "CommonPerm = 'Modify'"
	elseif permission = "FullControl" then
	  strSQL = strSQL & "CommonPerm = 'Full Control'"
	elseif permission = "Editable" then
	  strSQL = strSQL & "(CommonPerm = 'Full Control' or CommonPerm = 'Modify')"
	else
	  strSQL = strSQL & "(CommonPerm <> 'Read & Execute / List Folder Contents (folders only)' and CommonPerm <> 'Modify' and CommonPerm <> 'Full Control')"
	end if
  end if

  if right(strSQL, 6) = "where " then strSQL = strSQL & "1 = 0;"
end if
strSQL = strSQL & " " & OrderSQL 'Add the order by info
'response.write(strSQL)

rsCheckIn.LockType = 1
rsCheckIn.CursorType = 2

'Open the recordset with the SQL query 
rsCheckIn.Open strSQL, adoCon

'Loop through the recordset 
i = 0
inRec = 0
outRec = 0
Do While not rsCheckIn.EOF 
     
    'Write the HTML to display the current record in the recordset
    if 1 = 1 then

      'Display injects in red
	  if lcase(rsCheckIn("GrouporName")) = lcase(strDomain & "\" & InjectGrp) then
	    Response.Write ("<tr bgcolor=""FF6347"">") 'Tomato Red
	    if i = 0 then i = 1 else i = 0
	  else
	    'Display the records in Green Bar
        if i = 0 then
          Response.Write ("<tr>")
          i = 1
        else
          Response.Write ("<tr bgcolor=""CCFFCC"">") 'Green
          i = 0
        end if
	  end if

      'dim testz
	  'testz = rsCheckIn("Folder")
	  Response.Write ("<td>")
      Response.Write (rsCheckIn("Folder")) 
      Response.Write ("</td>")
	  Response.Write ("<td><center>")
      Response.Write (rsCheckIn("GrouporName"))
      Response.Write ("<td><center>") 
      Response.Write (rsCheckIn("CommonPerm"))
      Response.Write ("</center></td>")
      Response.Write ("<td><center>") 
      Response.Write ("<a href=""#"" OnClick=""netwin=dhtmlwindow.open('SpecialPermissions', 'ajax', 'nalr_special.asp?ID=" & rsCheckIn("ID") & "', 'Special Permissions', 'width=500px,height=300px,left=210px,top=50px,resize=1,scrolling=1'); return false"">View Special</a>")
	  Response.Write (" <a href=""index.asp?filter=Search&folder=" & rsCheckIn("Folder") & "&name=" & replace(rsCheckIn("GrouporName"),strDomain & "\","") & """>Audit</a>")
      Response.Write ("</center></td>")
  
      Response.Write ("</tr>")
    end if
  'Move to the next record in the recordset 
  rsCheckIn.MoveNext 
Loop

'Reset server objects
rsCheckIn.Close
Set rsCheckIn = Nothing
Set adoCon = Nothing


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
      response.write "Formating issue on page. Unrecognised format: " & sFormat
	end if
	
	Format = nExpression
  end if
End Function
%>
</font>
</table>
</body>
</html> 