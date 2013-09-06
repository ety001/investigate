<%
Set conn=server.createobject("adodb.connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath("Database/investigate.mdb")
%>