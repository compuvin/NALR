<html>
<head>
<title>Network Access Review</title>
<%
'Edit these for your organization
Dim DatabaseNALR

DatabaseNALR = "C:\NALR\NALR.accdb" 'Database location
AdminGrp = "MIS" 'Name of the domain admin group (group that can acknowlege changes)
%>
<script language="JavaScript">
  var d = new Date();
  function showhidedate()
  {
    if (document.getElementById("DateRange").style.display == "none")
    {
	  document.getElementById("DateRange").style.display = "block";
	  document.getElementById("StartDate").value = d.getMonth()+1 + "/" + d.getDate() + "/" + d.getYear();
	  document.getElementById("EndDate").value = d.getMonth()+1 + "/" + d.getDate() + "/" + d.getYear();
    }
	else
	{
	  document.getElementById("DateRange").style.display = "none";
	  document.getElementById("StartDate").value = "";
	  document.getElementById("EndDate").value = "";
	}
  }
</script>
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
Options: <a href="#" onclick=showhidedate()>Edit date range</a> | <a href="investigate.asp">Investigate user</a> | <a href="index.asp">Display last 7 days</a>
<br><br>

<!-- Begin form code -->
<form name="form" method="post" action="nalr_results.asp">
  <div id='DateRange' style='display:none'>
    Start Date: <input type="text" name="StartDate" id="StartDate" maxlength="10">&#160;&#160; End Date: <input type="text" name="EndDate" id="EndDate" maxlength="10">
    <br><br>
  </div>
  Folder: <input type="text" name="Folder" maxlength="50">&#160;&#160;
  
  Name (or group): <input type="text" name="Name" maxlength="50">
  <small>Enum Grps<input type="checkbox" name="Group"></small>&#160;&#160;

  Type of Change: 
    <select name="TypeChange">
      <option value="Any">Any</option>
      <option value="New Access">New Access</option>
      <option value="Removed">Removed</option>
      <option value="Other">Other</option>
    </select>&#160;&#160;

  Permission: 
    <select name="Permission">
      <option value="Any">Any</option>
      <option value="RX">Read and Execute</option>
      <option value="Modify">Modify</option>
	  <option value="FullControl">Full Control</option>
	  <option value="Editable">Editable (Modify or Full)</option>
	  <option value="Special">Special/Other</option>
    </select>&#160;&#160;

  Acknowleged: 
    <select name="Acknowleged">
      <option value="Any">Any</option>
      <option value="Yes">Yes</option>
      <option value="No">No</option>
    </select>&#160;&#160;

  <input type="submit" name="Submit" value="Search">
</form>
<!-- End form code -->
</center>

<table style="font-size:14px;" BORDER="1" CELLSPACING="1" CELLPADDING="1" WIDTH="100%" bgcolor="ffffff">
<font size=2>
<tr>
  <td><center><b>Folder</b></center></td>
  <td><center><b>Type of Change</b></center></td>
  <td><center><b>User Affected</center></b></td>
  <td><center><b>Previous Permission</center></b></td>
  <td><center><b>New Permission</center></b></td>
  <td><center><b>Who Changed</center></b></td>
  <td><center><b>Acknowledged By</center></b></td>
    <!--<td><center><b>CommentChange</center></b></td>-->
  <td><center><b>Date Changed</center></b></td>
  <td><center><b>Edit</center></b></td>
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
typechange = Request.QueryString("typechange")
permission = Request.QueryString("permission")
acknowleged = Request.QueryString("acknowleged")
StartDate = Request.QueryString("StartDate")
EndDate = Request.QueryString("EndDate")

'Create an ADO connection object
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & DatabaseNALR

'Create an ADO recordset object
Set rsCheckIn = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "select * from ChangeLog where format(DateChanged, 'YYYYMMDD') > " & format(date()-7, "YYYYMMDD")
OrderSQL = "order by DateChanged,Folder"

'Filters
if filter = "Search" then

  strSQL = "select * from ChangeLog where "
  if EGrp = "on" and len(name)>1 then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
    strSQL = strSQL & "(UserPerm = '" & strDomain & "\" & name & "'"
    rsCheckIn.Open "select * from GroupMembers where UserName = '" & name & "'", adoCon
	Do While not rsCheckIn.EOF 
	  strSQL = strSQL & " or UserPerm = '" & strDomain & "\" & rsCheckIn("GroupName") & "'"
	  rsCheckIn.MoveNext
	loop
	strSQL = strSQL & ")"
	OrderSQL = "order by Folder,UserPerm"
	rsCheckIn.Close
  elseif EGrp = "" and len(name)>1 then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
    strSQL = strSQL & "UserPerm = '" & strDomain & "\" & name & "'"
	OrderSQL = "order by Folder,UserPerm"
  end if
  if len(folder) > 1 then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
    strSQL = strSQL & "Folder like '%" & folder & "%'"
  end if
  if not permission = "" and not permission = "Any" then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
	if permission = "RX" then
	  strSQL = strSQL & "(NewPerm = 'Read & Execute / List Folder Contents (folders only)' or PrevPerm = 'Read & Execute / List Folder Contents (folders only)')"
	elseif permission = "Modify" then
	  strSQL = strSQL & "(NewPerm = 'Modify' or PrevPerm = 'Modify')"
	elseif permission = "FullControl" then
	  strSQL = strSQL & "(NewPerm = 'Full Control' or PrevPerm = 'Full Control')"
	elseif permission = "Editable" then 'Anyone who can change/edit a file (Modify and Full Control)
	  strSQL = strSQL & "(NewPerm = 'Modify' or PrevPerm = 'Modify' or NewPerm = 'Full Control' or PrevPerm = 'Full Control')"
	else
	  strSQL = strSQL & "((NewPerm <> 'Read & Execute / List Folder Contents (folders only)' and NewPerm <> 'Modify' and NewPerm <> 'Full Control' and NewPerm <> 'None') or (PrevPerm <> 'Read & Execute / List Folder Contents (folders only)' and PrevPerm <> 'Modify' and PrevPerm <> 'Full Control' and PrevPerm <> 'None'))"
	end if
  end if
  if acknowleged = "Yes" then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
	strSQL = strSQL & "Len(AckChange & '') > 1"
  elseif acknowleged = "No" then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
	strSQL = strSQL & "Len(AckChange & '') < 1"
  end if
  if not typechange = "" and not typechange = "Any" then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
	if typechange = "Other" then
	  strSQL = strSQL & "(TypeChange <> 'New Access' and TypeChange <> 'Removed')"
	else
	  strSQL = strSQL & "TypeChange = '" & typechange & "'"
	end if
  end if
  if not StartDate = "" and not StartDate = "Any" then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
	strSQL = strSQL & "format(DateChanged, 'YYYYMMDD') >= " & StartDate
	OrderSQL = "order by DateChanged,Folder"
  end if
  if not EndDate = "" and not EndDate = "Any" then
    if not right(strSQL, 6) = "where " then strSQL = strSQL & " and "
	strSQL = strSQL & "format(DateChanged, 'YYYYMMDD') <= " & EndDate
	OrderSQL = "order by DateChanged,Folder"
  end if

  if right(strSQL, 6) = "where " then
    strSQL = strSQL & "format(DateChanged, 'YYYYMMDD') > " & format(date()-7, "YYYYMMDD")
	OrderSQL = "order by DateChanged,Folder"
  end if
end if
strSQL = strSQL & " " & OrderSQL 'Add the order by info
'response.write(strSQL)

rsCheckIn.LockType = 1 'ReadOnly

'Open the recordset with the SQL query 
rsCheckIn.Open strSQL, adoCon

'Loop through the recordset 
i = 0
inRec = 0
outRec = 0
Do While not rsCheckIn.EOF 
           
    'Write the HTML to display the current record in the recordset
    if 1 = 1 then

      'Display the records in Green Bar
      if i = 0 then
        Response.Write ("<tr>")
        i = 1
      else
        Response.Write ("<tr bgcolor=""CCFFCC"">") 'Green
        i = 0
      end if

      Response.Write ("<td>")
      Response.Write (rsCheckIn("Folder")) 
      Response.Write ("</td>")
      Response.Write ("<td><center>") 
      Response.Write (rsCheckIn("TypeChange"))
      Response.Write ("</center></td>")
	  Response.Write ("<td><center>")
      Response.Write (rsCheckIn("UserPerm"))
      Response.Write ("<td><center>") 
      Response.Write (rsCheckIn("PrevPerm"))
      Response.Write ("</center></td>")
      Response.Write ("<td><center>") 
      Response.Write (rsCheckIn("NewPerm"))
      Response.Write ("</center></td>")
      Response.Write ("<td><center>")
      if rsCheckIn("WhoChange") & "" = "" then
	    Response.Write "-"
	  else
	    Response.Write (rsCheckIn("WhoChange"))
	  end if
      Response.Write ("</center></td>")
	  Response.Write ("<td><center>") 
      if rsCheckIn("AckChange") & "" = "" then
	    Response.Write "-"
	  else
	    Response.Write (rsCheckIn("AckChange"))
	  end if
      Response.Write ("</center></td>")
	  Response.Write ("<td><center>") 
      Response.Write (rsCheckIn("DateChanged"))
      Response.Write ("</center></td>")
      Response.Write ("<td><center>") 
      Response.Write ("<a href=""nalr_update.asp?ID=" & rsCheckIn("ID") & """>Edit</a>")
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