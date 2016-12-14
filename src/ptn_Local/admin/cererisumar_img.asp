<%@ Language=VBScript %>
<%
 Set ch = Server.CreateObject("VMAObjects.ASPChart") 
 ch.DefineCanvas "Data chart", 250,180
 ch.AddBar "Students:",CInt(Request.QueryString("S"))
 ch.AddBar "Professors:", CInt(Request.QueryString("P"))
 ch.AddBar "Administrators:", CInt(Request.QueryString("A"))
 ch.GenerateChart 
 set ch=nothing
%>
