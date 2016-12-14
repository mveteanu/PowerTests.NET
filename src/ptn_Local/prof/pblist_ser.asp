<%@ Language=VBScript %>
<%
 Const MsgDelOK = """Selected questions were deleted."", vbOkOnly+vbInformation"
 Const MsgRenOK = """Selected question was renamed."", vbOkOnly+vbInformation"
 Const MsgError = """DB errors occured."", vbOkOnly+vbCritical"
 Dim Mesaj, ReqAct
 
 on error resume next
 
 set cn = server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
  
 ReqAct = Request.Form("SelectAction")
 if ReqAct="del" then
   DeleteProblems Request.Form("SelectList"), cn
   Mesaj = MsgDelOK
 elseif ReqAct="ren" then
   RenameProblem Request.Form("SelectList"), Request.Form("SelectValue"), cn
   Mesaj = MsgRenOK
 end if
 
 cn.Close
 set cn = nothing
 
 if Err.number <> 0 then
   Mesaj = MsgError
   Err.Clear 
 end if
 
 Sub DeleteProblems(pbids, objCon)
   Const SQLDel = "DELETE FROM TBProblems WHERE id_problem IN (@1)"
   objCon.Execute Replace(SQLDel, "@1", pbids)
 End Sub

 Sub RenameProblem(pbid, newname, objCon)
   Const SQLRen = "UPDATE TBProblems SET numeproblema = '@1' WHERE id_problem = @2"
   objCon.Execute Replace(Replace(SQLRen, "@2", CStr(pbid)), "@1", newname)
 End Sub
%>
<body>
<script language=vbscript>
  window.parent.ReloadTDC
  msgbox <%=Mesaj%>
</script>
</body>
