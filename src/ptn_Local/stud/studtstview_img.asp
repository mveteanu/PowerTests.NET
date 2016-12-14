<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/utils.asp" -->
<%
 set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")

 Select Case Request.QueryString("imgid")
  case 1 PrintTwoBarImage "Points summary", "Test points:", "Obtained points:", Request.QueryString("v1"), Request.QueryString("v2")
  case 2 PrintTwoBarImage "Time summary", "Permited time:", "Used time:", Request.QueryString("v1"), Request.QueryString("v2")
  case 3 PrintPbResultsImage "Points / question", Request.QueryString("packedpbs"), cn
 End Select

 cn.Close
 set cn = nothing

 Sub PrintTwoBarImage(title, name1, name2, val1, val2)
    Dim ch
    
	Set ch = Server.CreateObject("VMAObjects.ASPChart") 
	ch.DefineCanvas title, 250,150
	ch.AddBar name2, CDbl(val2)
	ch.AddBar name1, CDbl(val1)
	ch.GenerateChart 
	set ch=nothing
 End Sub
 
 Sub PrintPbResultsImage(title, pbpakedvals, objCon)
  Const SQLSel = "SELECT * FROM TBProblems WHERE id_problem IN (@1)"
  Dim dic, pbids, rs, ch
    
  Set dic = UnpackDictFromString(pbpakedvals,"=", ";")
  pbids   = DictkeysToCSVList(dic)
  Set ch = Server.CreateObject("VMAObjects.ASPChart") 
  ch.DefineCanvas title, 500, dic.Count * 40

  set rs = Server.CreateObject("ADODB.Recordset")
  Set rs.ActiveConnection = objCon
  rs.CursorLocation = adUseClient
  rs.CursorType = adOpenStatic
  rs.Open Replace(SQLSel, "@1", pbids)
  rs.MoveLast
  do until rs.BOF
   ch.AddBar rs.Fields("numeproblema").value, dic.Item(CStr(rs.Fields("id_problem").value))
   rs.MovePrevious
  loop
  ch.GenerateChart 
  set ch=nothing
  rs.Close
  set rs = nothing
 End Sub
%>