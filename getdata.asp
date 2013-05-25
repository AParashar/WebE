
<html>
<head>
<title>My First ASP Page</title>
</head>
<body bgcolor="white" text="black">

<% 
'Dimension variables
Dim adoCon       'Holds the Database Connection Object
Dim rsUser   	 'Holds the recordset for the records in the database
Dim strSQL       'Holds the SQL query to query the database
Dim pass
Dim user
Dim passf
Dim userf
passf=Request.Form("pass")
userf=Request.Form("user")
Set adoCon = Server.CreateObject("ADODB.Connection")
adoCon.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("mydb.mdb")
Set rsUser = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT * FROM registration;"
rsUser.Open strSQL, adoCon
Do While not rsUser.EOF 
    user=rsUser("username") 
    pass=rsUser("password") 
    'Move to the next record in the recordset 
    rsUser.MoveNext 
	If user=userf And pass=passf Then
		rsUser.Close
		Set rsUser = Nothing
		Set adoCon = Nothing
		'Redirect to the Home.html page
		Response.Write("Welcome "&user)
		Response.Redirect "medicine.html"
	End If
Loop
Response.Write("<b><br><br><br><br><br><br><br><br><br><br><br><Font size='20'><center>Wrong email id or Password</font>")
Response.Write("<b><br><center><a href='login.html'>Click to go back</a>")
'Reset server objects
rsUser.Close
Set rsUser = Nothing
Set adoCon = Nothing
Response.Redirect "pur.html"
%>

</body>
</html>