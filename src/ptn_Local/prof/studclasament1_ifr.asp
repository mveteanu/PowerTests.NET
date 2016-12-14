<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 ' Tipareste raportul final
 Sub PrintClasamentReport
  Dim cn, TestID
  
  Set cn = Server.CreateObject("ADODB.Connection")
  cn.Open Application("DSN")
  TestID = Request.QueryString("TestID")
  If TestID="" then
    PrintCursInfo Session("CursID"), cn
    PrintStudentsTopTable Session("CursID"), cn
  Else
    PrintTestInfo TestID, cn
    PrintStudentsTopTableByTest TestID, cn
  End If  
  cn.Close
  set cn = nothing
 End Sub
 
 ' Tipareste headerul raportului cu informatii despre curs
 Sub PrintCursInfo(CursID, objCon)
  Dim myCmd, rs
  
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetCursByID"
  Set rs = myCmd.Execute(,CLng(CursID))
  Set myCmd = Nothing

  With Response
    .Write "<br><table border=0 align=center width=100% cellspacing=0 bgcolor=#f0f0f0 style='border-top:1px solid black;border-bottom:1px solid black;'><tr><td>"
    .Write "<b>Class name:</b> " & rs.Fields("numecurs").Value & "<br>"
    .Write "<b>Professor name:</b> " & rs.Fields("profname").Value & "<br>"
    .Write "<b>Enrolled students:</b> " & rs.Fields("StudentiInscrisi").Value & "<br>"
    .Write "<b>Sorted by:</b> Summary of all test results"
    .Write "</td></tr></table><br>"
  End With

  rs.Close
  Set rs = nothing
 End Sub

 ' Tipareste headerul raportului cu informatii despre test
 Sub PrintTestInfo(TestIDL, objCon)
  Const SQLSelTst = "SELECT TBTests.*, TBCursuri.numecurs FROM TBCursuri INNER JOIN TBTests ON TBCursuri.id_curs = TBTests.id_curs WHERE TBTests.id_test=@1"
  Dim rs
  
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetTstRezInfo"
  Set rs = myCmd.Execute(,CLng(TestIDL))
  Set myCmd = Nothing

  With Response
    .Write "<br><table border=0 align=center width=100% cellspacing=0 bgcolor=#f0f0f0 style='border-top:1px solid black;border-bottom:1px solid black;'><tr><td>"
    .Write "<b>Test name:</b> " & rs.Fields("numetest").Value & "<br>"
    .Write "<b>Course name:</b> " & rs.Fields("numecurs").Value & "<br>"
    .Write "<b>Professor name:</b> " & rs.Fields("profname").Value & "<br>"
    .Write "<b>Sorted by:</b> Summary of test results"
    .Write "</td></tr></table><br>"
  End With

  rs.Close
  Set rs = nothing
 End Sub
 
 ' Tipareste tabelul cu clasamentul dupa media notelor de la curs
 Sub PrintStudentsTopTable(CursID, objCon)
  Dim myCmd, rs, contor
  
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetClasamentByCurs"
  Set rs = myCmd.Execute(,CLng(CursID))
  Set myCmd = Nothing
  
  If rs.EOF then
   Response.Write "Insufficient data (maybe no test was completed yet)."
  Else
   With Response
    .Write "<table border=0 cellspacing=0 width=100% >"
    .Write "<tr bgcolor=#f0f0f0><td width=5% ><b>No</b></td><td width=25% ><b>Student name</b></td><td width=25% ><b>Completed tests</b></td><td width=25% ><b>Abandoned tests</b></td><td width=20% ><b>Final score</b></td></tr>"
   End With 
   contor = 1
   do until rs.EOF
    With Response
     .Write "<tr><td>"& contor &".</td><td>"& rs.Fields("numestud").value &"</td><td>"& NullToZero(rs.Fields("TesteRezolvate").value) &"</td><td>"& NullToZero(rs.Fields("TesteNefinalizate").value) &"</td><td>"& CompareToString2(rs.Fields("MediaNote").value,null,"Undefined") &"</td></tr>"
    End With
    contor = contor + 1
    rs.MoveNext
   loop
   Response.Write "</table>"
  End If
  
  rs.Close
  set rs = nothing
 End Sub

 ' Tipareste tabelul cu clasamentul dupa media notelor de la curs
 ' dupa un anumit test
 Sub PrintStudentsTopTableByTest(TestIDL, objCon)
  Dim myCmd, rs, contor
  
  Set myCmd = Server.CreateObject("ADODB.Command")
  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetClasamentByTest"
  Set rs = myCmd.Execute(,CLng(TestIDL))
  Set myCmd = Nothing
  
  If rs.EOF then
   Response.Write "Insufficient data (maybe no test was completed yet)."
  Else
   With Response
    .Write "<table border=0 cellspacing=0 width=100% >"
    .Write "<tr bgcolor=#f0f0f0><td width=5% ><b>No</b></td><td width=25% ><b>Student name</b></td><td width=25% ><b>Completed tests</b></td><td width=25% ><b>Abandoned tests</b></td><td width=20% ><b>Final score</b></td></tr>"
   End With 
   contor = 1
   do until rs.EOF
    With Response
     .Write "<tr><td>"& contor &".</td><td>"& rs.Fields("numestud").value &"</td><td>"& NullToZero(rs.Fields("NrRezolvari").value) &"</td><td>"& NullToZero(rs.Fields("NrNeterminari").value) &"</td><td>"& CompareToString2(rs.Fields("MediaNote").value,null,"Undefined") &"</td></tr>"
    End With
    contor = contor + 1
    rs.MoveNext
   loop
   Response.Write "</table>"
  End If
  
  rs.Close
  set rs = nothing
 End Sub
%>
<html>
<head>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<body unselectable="on" class=White style="behavior:url('../_clientscripts/application.htc');">

<%
PrintClasamentReport
%>

<script language=vbscript>
 Sub Window_onload
  window.parent.HandleIframeLoading false
 End Sub 
</script>

</body>
</html>
