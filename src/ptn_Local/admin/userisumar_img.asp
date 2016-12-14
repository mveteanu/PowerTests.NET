<%@ Language=VBScript %>
<%
 Set ch = Server.CreateObject("VMAObjects.ASPChart") 
 ch.DefineCanvas "Data chart", 250,180
 ch.AddBar "Locked students:",CInt(Request.QueryString("SB"))
 ch.AddBar "Students:",CInt(Request.QueryString("S"))
 ch.AddBar "Locked professors:", CInt(Request.QueryString("PB"))
 ch.AddBar "Professors:", CInt(Request.QueryString("P"))
 ch.AddBar "Locked admins:", CInt(Request.QueryString("AB"))
 ch.AddBar "Administrators:", CInt(Request.QueryString("A"))
 ch.GenerateChart 
 set ch=nothing
%>
