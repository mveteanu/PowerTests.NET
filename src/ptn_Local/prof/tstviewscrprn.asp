<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/tests.asp" -->
<%
 Dim qs, PBViewer
 
 Response.Buffer = True
 Response.Expires = -1

 set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 set tst1 = New PTNTestGenerator
 tst1.LoadTest Request.QueryString("TstID"), cn
 qs = Join(tst1.GetGeneratedTest,",")
 set tst1 = nothing
 cn.Close
 set cn = nothing
 
 If Request.QueryString("prn")="" then
  PBViewer = "pbviewscr.asp"
 Else
  PBViewer = "pbviewprn.asp"
 End If 
 
 Response.Redirect PBViewer & "?PBIDs=" & qs
%>