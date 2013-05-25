<%
'Dimension variables
Dim adoCon          'Holds the Database Connection Object
Dim rsUser   'Holds the recordset for the new record to be added
Dim strSQL 		'Holds the SQL query to query the database
		
'Create an ADO connection object
Set adoCon = Server.CreateObject("ADODB.Connection") 
'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("mydb.mdb")
'Create an ADO recordset object
Set rsUser = Server.CreateObject("ADODB.Recordset")	
'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT * FROM dist;"
'Set the cursor type we are using so we can navigate through the recordset
rsUser.CursorType = 2
'Set the lock type so that the record is locked by ADO when it is updated
rsUser.LockType = 3

'Open the recordset with the SQL query 
rsUser.Open strSQL, adoCon
'Tell the recordset we are adding a new record to it
rsUser.AddNew	

'Add a new record to the recordset
rsUser.Fields("shop_name") = Request.Form("name")
rsUser.Fields("address") = Request.Form("uname")
rsUser.Fields("med_comp") = Request.Form("med")
rsUser.Fields("pdate") = Request.Form("date")


'Write the updated recordset to the database
rsUser.Update


'Reset server objects
rsUser.Close
Set rsUser = Nothing
Set adoCon = Nothing

'Redirect to the guestbook.asp page
Response.Redirect "index.html"
%>



