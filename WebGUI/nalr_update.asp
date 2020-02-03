<%
'Edit these for your organization
Dim DatabaseNALR

DatabaseNALR = "C:\NALR\NALR.accdb" 'Database location
AdminGrp = "MIS" 'Name of the domain admin group (group that can acknowlege changes)
%>
<%
'Dimension variables
Dim adoCon 			'Holds the Database Connection Object
Dim rsCheckIn			'Holds the recordset for the record to be updated
Dim strSQL			'Holds the SQL query for the database
Dim lngRecordNo			'Holds the record number to be updated
dim WebAddr
dim RecDate
dim AdminList, AdminTemp

'Read in the record number to be updated
lngRecordNo = Request.QueryString("ID")
WebAddr = Request.ServerVariables("HTTP_REFERER")

' Script to get the domain and name of the logged on user.
Private strDomain
Private strUserName
Private wshNetwork

Set wshNetwork = CreateObject("WScript.Network")

strDomain = wshNetwork.UserDomain
strUserName = lcase(wshNetwork.UserName)

Set wshNetwork = Nothing


Set fso = Server.CreateObject("scripting.FileSystemObject")

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & DatabaseNALR

'Create an ADO recordset object
Set rsCheckIn = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT * FROM ChangeLog WHERE ID=" & lngRecordNo

'Set the lock type so that the record is locked by ADO when it is updated
rsCheckIn.LockType = 1

'Quickly get the admins
rsCheckIn.Open "select * from GroupMembers where GroupName = '" & AdminGrp & "'", adoCon

Do While not rsCheckIn.EOF 
	AdminList = AdminList & rsCheckIn("UserName") & "|"
	rsCheckIn.MoveNext
loop
rsCheckIn.Close

'Set the lock type so that the record is locked by ADO when it is updated
rsCheckIn.LockType = 1

'Open the recordset with the SQL query 
rsCheckIn.Open strSQL, adoCon
%>
<html>
<head>
<title>NALR Update Form</title>
<script>
function CheckAdmin() {
    if (document.getElementById("WhoChange").value == document.getElementById("AckChange").value)
	{
      document.getElementById("AckChange").value = "Comments Only";
	  document.getElementById("AckChangeSub").value = "Comments Only";
	}
	else
	{
	  document.getElementById("AckChange").value="<%if len(rsCheckIn("AckChange")) > 1 then response.write(rsCheckIn("AckChange")) else response.write(strUserName)%>";
      document.getElementById("AckChangeSub").value = document.getElementById("AckChange").value;
	}
}
</script>
</head>
<body bgcolor="white" text="black">
<!-- Begin form code -->
<form name="form" method="post" action="nalr_entry.asp">
  <b>Folder:</b> <% = rsCheckIn("Folder") %>
  <br><br>
  <b>Type of Change:</b> <% = rsCheckIn("TypeChange") %>
  <br>
  <b>New Permission:</b> <% = rsCheckIn("UserPerm") %>
  <br>
  <b>Previous Permission:</b> <% = rsCheckIn("PrevPerm") %>
  <br>
  <b>New Permission:</b> <% = rsCheckIn("NewPerm") %>
  <br><br>
  <b>Who Made the Change:</b>
  <select onchange=CheckAdmin() name="WhoChange" id="WhoChange">
      <option value=""></option>
	  <%
	    AdminTemp = AdminList
		do while len(AdminList) > 1
		  if mid(AdminList,1,instr(1,AdminList,"|",1)-1) = rsCheckIn("WhoChange") then
		    response.write("<option selected value='" & mid(AdminList,1,instr(1,AdminList,"|",1)-1) & "'>" & mid(AdminList,1,instr(1,AdminList,"|",1)-1) & "</option>")
		  else
		    response.write("<option value='" & mid(AdminList,1,instr(1,AdminList,"|",1)-1) & "'>" & mid(AdminList,1,instr(1,AdminList,"|",1)-1) & "</option>")
		  end if
		  AdminList = right(AdminList,len(AdminList)-instr(1,AdminList,"|",1))
		loop
		AdminList = AdminTemp
		if rsCheckIn("WhoChange") = "Folder Moved/Renamed" then
		  response.write("<option selected value=""Folder Moved/Renamed"">Folder Moved/Renamed</option>")
		else
		  response.write("<option value=""Folder Moved/Renamed"">Folder Moved/Renamed</option>")
		End if
	  %>
    </select>&#160;&#160;  
  <br>
  <b>Acknowledged By:</b> &#160;
  <input type="text" name="AckChange" id="AckChange" disabled maxlength="50" value="<%if len(rsCheckIn("AckChange")) > 1 then response.write(rsCheckIn("AckChange")) else response.write(strUserName)%>">
  <input type="hidden" name="AckChangeSub" Id="AckChangeSub" value="<%if len(rsCheckIn("AckChange")) > 1 then response.write(rsCheckIn("AckChange")) else response.write(strUserName)%>">
  <script>CheckAdmin()</script>
  <br><br>
  <b>Comments:</b> <textarea name="CommentChange" cols=40 rows=5 wrap=physical maxlength="250"><% = rsCheckIn("CommentChange") %></textarea>  

  <br><br>

  <input type="hidden" name="WebAddr" value="<% = WebAddr %>">
  <input type="hidden" name="Login" value="<% = rsCheckIn("ID") %>">
  <input type="Button" value="Cancel" onclick="history.back()"> 

  <input type="submit" name="Submit" value="Submit">
</form>
<!-- End form code -->
</body>
</html>
<%
'Reset server objects
rsCheckIn.Close
Set rsCheckIn = Nothing
Set adoCon = Nothing
%>
