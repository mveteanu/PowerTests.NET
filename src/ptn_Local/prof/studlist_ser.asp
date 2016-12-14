<%@ Language=VBScript %>
<%
 Const MsgLockOK      = """Selected students were locked."", vbOkOnly+vbInformation"
 Const MsgUnLockOK    = """Selected students were unlocked."", vbOkOnly+vbInformation"
 Const MsgDelSubscrOK = """Selected students were un-enrolled."", vbOkOnly+vbInformation"
 Const MsgError       = """DB errors occured"", vbOkOnly+vbCritical" 
 
 Dim Mesaj, ReqAct

 on error resume next
 
 set cn = server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
  
 ReqAct = Request.Form("SelectAction")
 if ReqAct="lock" then
   LockStuds Request.Form("SelectList"), true, cn
   Mesaj = MsgLockOK
 elseif ReqAct="unlock" then
   LockStuds Request.Form("SelectList"), false, cn
   Mesaj = MsgUnLockOK
 elseif ReqAct="delsubscript" then
   DelSubscript Request.Form("SelectList"), cn
   Mesaj = MsgDelSubscrOK
 End If
 
 cn.Close
 set cn = nothing
 
 if Err.number <> 0 then
   Mesaj = MsgError
   Err.Clear 
 end if
 
 Sub LockStuds(studids, locktype, objCon)
  Const SQLLock = "UPDATE TBStudentiLaCursuri SET locked=@1 WHERE id_studcurs IN(@2)"
  objCon.Execute Replace(Replace(SQLLock, "@2", studids), "@1", CStr(locktype))
 End Sub

 Sub DelSubscript(studids, objCon)
  Const SQLDelS = "DELETE FROM TBStudentiLaCursuri WHERE id_studcurs IN(@1)"
  objCon.Execute Replace(SQLDelS, "@1", studids)
 End Sub
%>
<body>
<script language=vbscript>
  window.parent.ReloadTDC
  msgbox <%=Mesaj%>
</script>
</body>
