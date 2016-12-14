<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 Response.Buffer = True
 Response.Expires = -1

 ' Tipareste in pagina HTML tabelul cu rezultatele
 Sub PrintResultsReport
  Dim cn, rs
 
  set cn = Server.CreateObject("ADODB.Connection")
  cn.Open Application("DSN")
  PrintStudentInfo cn
  If CInt(Request.QueryString("tip")) = 2 then 
    PrintResultsGroupedByNote cn
  Else
    PrintResultsGroupedByTest cn
  End If
  PrintAbandonedTests cn
  cn.Close
  set cn = nothing 
 End Sub


 ' Tipareste headerul rasportului cu informatii despre student si curs
 Sub PrintStudentInfo(objCon)
  Dim myCmd, rs
  Dim NumeStud, NumeCurs, NumeProf, d1,d2
  
  Set myCmd = Server.CreateObject("ADODB.Command")

  Set myCmd.ActiveConnection = objCon
  myCmd.CommandText = "GetUserByID"
  Set rs = myCmd.Execute(,CLng(Request.QueryString("userid")))
  NumeStud = rs.Fields("nume").Value & " " & rs.Fields("prenume").Value
  rs.Close
  myCmd.Parameters.Delete(0)

  myCmd.CommandText = "GetCursByID"
  Set rs = myCmd.Execute(,CLng(Session("CursID")))
  NumeCurs = rs.Fields("numecurs").Value
  NumeProf = rs.Fields("profname").Value
  rs.Close
  myCmd.Parameters.Delete(0)

  Set rs = nothing
  Set myCmd = Nothing

  d1 = Request.QueryString("d1")
  d2 = Request.QueryString("d2")
  With Response
    .Write "<br><table border=0 align=center width=100% cellspacing=0 bgcolor=#f0f0f0 style='border-top:1px solid black;border-bottom:1px solid black;'><tr><td>"
    .Write "<b>Student name:</b> " & NumeStud & "<br>"
    .Write "<b>Course name:</b> " & NumeCurs & "<br>"
    .Write "<b>Professor name:</b> " & NumeProf & "<br>"
    .Write "<b>Time range:</b> "
    If d1 = "" then 
      .Write "All dates" 
    ElseIf d2 = "" then
      .Write "> " & d1
    Else
      .Write d1 & " - " & d2
    End If  
    .Write "</td></tr></table></br>"
  End With
 End Sub


' Tipareste testele abandonate grupate dupa test
Sub PrintAbandonedTests(objCon)
	Dim LastTestID, TestID
	
	set rs = Server.CreateObject("ADODB.Recordset")
	Set rs.ActiveConnection = objCon
	rs.CursorLocation   = adUseClient
	rs.CursorType       = adOpenStatic
	rs.LockType         = adLockOptimistic

	rs.Open GetSQLQuery(false, 1)
    Response.Write "<font size=+1 color=#000064>Abandoned tests</font><br><br>" & vbCrLf
    If rs.EOF then 
     Response.Write "No test was abandoned in this time range!"
    Else
     LastTestID = -1
     Response.Write "<table border=0 cellspacing=0 cellpadding=10 width=100% >" & vbCrLf & vbCrLf
     do until rs.EOF
      TestID = rs.Fields("id_test").value
      If LastTestID <> TestID then
        With Response
         If LastTestID<>-1 then .Write "</table></td></tr>" & vbCrLf & vbCrLf
         .Write "<tr><td colspan=2 valign=bottom align=left><b>Test name:</b> " & rs.Fields("numetest").value & "</td></tr>" & vbCrLf
         .Write "<tr><td width=10% >&nbsp;</td><td align=left valign=top><table cellspacing=0 border=0 width=36% >" & vbCrLf
         .Write "<tr bgcolor=#f0f0f0><td width=20% ><b>Solving date</b></td><td width=16% ><b>Total time</b></td></tr>" & vbCrLf
         .Write "<tr><td>"& DToSR(rs.Fields("datastarttest").value, "DD.MM.YYYY") &"</td><td>"& rs.Fields("timp").value &"</td>" & vbCrLf
        End With
        LastTestID = TestID
      Else
        Response.Write "<tr><td>"& DToSR(rs.Fields("datastarttest").value, "DD.MM.YYYY") &"</td><td>"& rs.Fields("timp").value &"</td>" & vbCrLf
      End If
      rs.MoveNext
      If (rs.EOF) and (LastTestID <> -1) then  Response.Write "</table></td></tr>" & vbCrLf & vbCrLf
     loop
     Response.Write "</table>" & vbCrLf
    End If 
	rs.Close

	set rs = nothing
End Sub


 ' Tipareste pagina cu rezultate grupate dupa test
 Sub PrintResultsGroupedByTest(objCon)
	Dim LastTestID, TestID
	
	set rs = Server.CreateObject("ADODB.Recordset")
	Set rs.ActiveConnection = objCon
	rs.CursorLocation   = adUseClient
	rs.CursorType       = adOpenStatic
	rs.LockType         = adLockOptimistic
	rs.Open GetSQLQuery(true, 1)
    Response.Write "<font size=+1 color=#000064>Completed tests</font><br><br>" & vbCrLf
    If rs.EOF then 
     Response.Write "No data available for specified time range!"
    Else
     LastTestID = -1
     Response.Write "<table border=0 cellspacing=0 cellpadding=10 width=100% >" & vbCrLf & vbCrLf
     do until rs.EOF
      TestID = rs.Fields("id_test").value
      If LastTestID <> TestID then
        With Response
         If LastTestID<>-1 then .Write "</table></td></tr>" & vbCrLf & vbCrLf
         .Write "<tr><td colspan=2 valign=bottom align=left><b>Test name:</b> " & rs.Fields("numetest").value & "</td></tr>" & vbCrLf
         .Write "<tr><td width=10% >&nbsp;</td><td align=left valign=top><table cellspacing=0 border=0 width=100% >" & vbCrLf
         .Write "<tr bgcolor=#f0f0f0><td width=20% ><b>Solving date</b></td><td width=16% ><b>Total time</b></td><td width=16% ><b>Used time</b></td><td width=16% ><b>Test points</b></td><td width=16% ><b>Obtained points</b></td><td width=16% ><b>Score</b></td></tr>" & vbCrLf
         .Write "<tr><td>"& DToSR(rs.Fields("datastarttest").value, "DD.MM.YYYY") &"</td><td>"& rs.Fields("timp").value &"</td><td>"& rs.Fields("timpfolosit").value &"</td><td>"& rs.Fields("punctetest").value &"</td><td>"& rs.Fields("puncteobtinute").value &"</td><td>"& rs.Fields("nota").value &"</td>" & vbCrLf
        End With
        LastTestID = TestID
      Else
        Response.Write "<tr><td>"& DToSR(rs.Fields("datastarttest").value, "DD.MM.YYYY") &"</td><td>"& rs.Fields("timp").value &"</td><td>"& rs.Fields("timpfolosit").value &"</td><td>"& rs.Fields("punctetest").value &"</td><td>"& rs.Fields("puncteobtinute").value &"</td><td>"& rs.Fields("nota").value &"</td>" & vbCrLf
      End If
      rs.MoveNext
      If (rs.EOF) and (LastTestID <> -1) then  Response.Write "</table></td></tr>" & vbCrLf & vbCrLf
     loop
     Response.Write "</table>" & vbCrLf
    End If 
	rs.Close

	Response.Write "<br><br>"
	
	set rs = nothing
 End Sub


 ' Tipareste pagina cu rezultate grupate dupa nota
 Sub PrintResultsGroupedByNote(objCon)
	Dim LastTestNota, TestNota
	
	set rs = Server.CreateObject("ADODB.Recordset")
	Set rs.ActiveConnection = objCon
	rs.CursorLocation   = adUseClient
	rs.CursorType       = adOpenStatic
	rs.LockType         = adLockOptimistic
	rs.Open GetSQLQuery(true, 2)

    Response.Write "<font size=+1 color=#000064>Completed tests</font><br><br>" & vbCrLf
    If rs.EOF then 
     Response.Write "No data available for specified time range!"
    Else
     LastTestNota = -1
     Response.Write "<table border=0 cellspacing=0 cellpadding=10 width=100% >" & vbCrLf & vbCrLf
     do until rs.EOF
      TestNota = rs.Fields("nota").value
      If LastTestNota <> TestNota then
        With Response
         If LastTestNota<>-1 then .Write "</table></td></tr>" & vbCrLf & vbCrLf
         .Write "<tr><td colspan=2 valign=bottom align=left><b>Obtained score:</b> " & rs.Fields("nota").value & "</td></tr>" & vbCrLf
         .Write "<tr><td width=10% >&nbsp;</td><td align=left valign=top><table cellspacing=0 border=0 width=100% >" & vbCrLf
         .Write "<tr bgcolor=#f0f0f0><td width=20% ><b>Test name</b></td><td width=16% ><b>Solving date</b></td><td width=16% ><b>Total time</b></td><td width=16% ><b>Used time</b></td><td width=16% ><b>Test points</b></td><td width=16% ><b>Obtained score</b></td></tr>" & vbCrLf
         .Write "<tr><td>"& rs.Fields("numetest").value &"</td><td>"& DToSR(rs.Fields("datastarttest").value, "DD.MM.YYYY") &"</td><td>"& rs.Fields("timp").value &"</td><td>"& rs.Fields("timpfolosit").value &"</td><td>"& rs.Fields("punctetest").value &"</td><td>"& rs.Fields("puncteobtinute").value &"</td>" & vbCrLf
        End With
        LastTestNota = TestNota
      Else
        Response.Write "<tr><td>"& rs.Fields("numetest").value &"</td><td>"& DToSR(rs.Fields("datastarttest").value, "DD.MM.YYYY") &"</td><td>"& rs.Fields("timp").value &"</td><td>"& rs.Fields("timpfolosit").value &"</td><td>"& rs.Fields("punctetest").value &"</td><td>"& rs.Fields("puncteobtinute").value &"</td>" & vbCrLf
      End If
      rs.MoveNext
      If (rs.EOF) and (LastTestNota <> -1) then  Response.Write "</table></td></tr>" & vbCrLf & vbCrLf
     loop
     Response.Write "</table>" & vbCrLf
    End If 
	rs.Close
 
	Response.Write "<br><br>"
	
	set rs = nothing
 End Sub


 ' Contruieste interogarea SQL folosita pentru extragerea informatiilor
 ' referitoare la rezultatele obtinute de un anumit student la teste
 Function GetSQLQuery(bFinishedTests, iOrderby)
	Const SQLSelMach1 = "SELECT id_fisarezultate, id_test, numetest, timp, datastarttest, dataendtest, punctetest, puncteobtinute, Round((10*puncteobtinute)/punctetest,2) AS nota, DateDiff(""n"", datastarttest, dataendtest) AS timpfolosit FROM TBStudentsResults WHERE (dataendtest Is Not Null) AND (id_user=@iduser) AND (id_test IN (@idtests))@datasupp" 
	Const SQLSelMach2 = "SELECT id_fisarezultate, id_test, numetest, timp, datastarttest FROM TBStudentsResults WHERE (dataendtest Is Null) AND (id_user=@iduser) AND (id_test IN (@idtests))@datasupp" 
	Dim DataS1, DataA1, DataS2, DataA2, re
   
	DataS1 = Request.QueryString("d1")
	If DataS1 <> "" then DataA1 = Split(DataS1, ".", -1,1) 
	DataS2 = Request.QueryString("d2")
	If DataS2 <> "" then DataA2 = Split(DataS2, ".", -1,1) 
 
    If bFinishedTests then re = SQLSelMach1 else re = SQLSelMach2
	re = Replace(Replace(re, "@idtests", Request.QueryString("tstids")), "@iduser", Request.QueryString("userid"))
	If IsEmpty(DataS1) then
	  re = Replace(re, "@datasupp", "")
	ElseIf IsEmpty(DataS2) then
	  re = Replace(re, "@datasupp", " AND (datastarttest > DateSerial("& DataA1(2) &","& DataA1(1) &","& DataA1(0) &"))")
	Else
	  re = Replace(re, "@datasupp", " AND (datastarttest BETWEEN DateSerial("& DataA1(2) &","& DataA1(1) &","& DataA1(0) &") and DateSerial("& DataA2(2) &","& DataA2(1) &","& DataA2(0) &"))")
	End If

    Select Case iOrderby
     Case 1 re = re & " ORDER BY numetest, datastarttest"
     Case 2 re = re & " ORDER BY Round((10*puncteobtinute)/punctetest,2) DESC, datastarttest"
    End Select
   
    GetSQLQuery = re
 End Function
%>
<html>
<head>
  <link rel="stylesheet" type="text/css" href="../css/ptn.css">
</head>
<body unselectable="on" class=White style="behavior:url('../_clientscripts/application.htc');">

<%
PrintResultsReport
%>

<script language=vbscript>
 Sub Window_onload
  window.parent.HandleIframeLoading false
 End Sub 
</script>

</body>
</html>
