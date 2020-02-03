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
dim RecDate

'Read in the record number to be updated
lngRecordNo = Request.QueryString("ID")

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & DatabaseNALR

'Create an ADO recordset object
Set rsCheckIn = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT * FROM DataDump WHERE ID=" & lngRecordNo

'Set the lock type so that the record is locked by ADO when it is updated
rsCheckIn.LockType = 1

'Open the recordset with the SQL query 
rsCheckIn.Open strSQL, adoCon

rsCheckIn.MoveFirst
%>
<html>
<head>
<title>Special Permissions</title>
</head>
<body bgcolor="white" text="black">
<%response.write("Folder: <b>" & rsCheckIn("Folder") & "</b><br><br>" & rsCheckIn("SpecialPerm"))%>
</body>
</html>
<%
'Reset server objects
rsCheckIn.Close
Set rsCheckIn = Nothing
Set adoCon = Nothing
%>
