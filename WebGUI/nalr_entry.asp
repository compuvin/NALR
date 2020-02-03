<%
'Edit these for your organization
Dim DatabaseNALR

DatabaseNALR = "C:\NALR\NALR.accdb" 'Database location
AdminGrp = "MIS" 'Name of the domain admin group (group that can acknowlege changes)
%>
<%
'Dimension variables
Dim adoCon 			'Holds the Database Connection Object
Dim rsUpdateEntry		'Holds the recordset for the record to be updated
Dim strSQL			'Holds the SQL query for the database
Dim lngRecordNo			'Holds the record number to be updated

if len(Request.Form("WhoChange")) < 1 or _
   len(Request.Form("CommentChange")) < 1 or _
   len(Request.Form("CommentChange")) > 254 then
  
  response.write("<script type='text/JavaScript'>alert('Please provide all of the required information on the form. The comments section cannot be longer than 254 charactors.')</script>") 
  Response.write("<script>history.back()</script>")
else

  'Read in the record number to be updated
  lngRecordNo = Request.Form("Login")

  'Create an ADO connection odject
  Set adoCon = Server.CreateObject("ADODB.Connection")

  'Set an active connection to the Connection object using a DSN-less connection
  adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & DatabaseNALR

  'Create an ADO recordset object
  Set rsUpdateEntry = Server.CreateObject("ADODB.Recordset")

  'Initialise the strSQL variable with an SQL statement to query the database
  strSQL = "SELECT * FROM ChangeLog WHERE ID=" & lngRecordNo

  'Set the cursor type we are using so we can navigate through the recordset
  rsUpdateEntry.CursorType = 2

  'Set the lock type so that the record is locked by ADO when it is updated
  rsUpdateEntry.LockType = 3

  'Open the tblComments table using the SQL query held in the strSQL varaiable
  rsUpdateEntry.Open strSQL, adoCon

  'Update the record in the recordset
  rsUpdateEntry.Fields("WhoChange") = Request.Form("WhoChange")
  if Request.Form("AckChangeSub") = "Comments Only" then
    rsUpdateEntry.Fields("AckChange") = ""
  else
    rsUpdateEntry.Fields("AckChange") = Request.Form("AckChangeSub")
  end if
  rsUpdateEntry.Fields("CommentChange") = Request.Form("CommentChange")


  'Write the updated recordset to the database
  rsUpdateEntry.Update

  'Reset server objects
  rsUpdateEntry.Close
  Set rsUpdateEntry = Nothing
  Set adoCon = Nothing

  'Return to the update select page in case another record needs deleting
  Response.Redirect Request.Form("WebAddr")
end if
%>