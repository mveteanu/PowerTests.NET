<%@ Language=VBScript %>
<%
 Const MsgDelOK = """Selected categories were deleted successfully."", vbOKOnly+vbInformation"
 Const MsgRenOK = """Selected category was successfully renamed."", vbOKOnly+vbInformation"
 Const MsgAddOK = """Categoy added successfully."", vbOKOnly+vbInformation"
 Const MsgEdtOK = """Category was edited successfully."", vbOKOnly+vbInformation"
 Const MsgErr   = """DB errors occured."", vbOKOnly+vbCritical"
 
 Dim RequestAction, RequestList, RequestValue
 Dim cn
 
 RequestAction = Request.Form("SelectAction")
 RequestList   = Request.Form("SelectList")
 RequestValue  = Request.Form("SelectValue")
 
 on error resume next
 
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Select Case LCase(RequestAction)
   case "del" DeleteCategs RequestList, cn
              Mesaj = MsgDelOK
   case "ren" RenameCateg RequestList, RequestValue, cn
              Mesaj = MsgRenOK
   case "add" AddCateg Session("CursID"), RequestValue, cn
              Mesaj = MsgAddOK
   case "edt" AddPBInCateg RequestList, RequestValue, cn
              Mesaj = MsgEdtOK
 End Select

 cn.Close
 Set cn = nothing

 If Err.number <> 0 then
   Err.Clear 
   Mesaj = MsgErr
 End If

 
 ' Introduce lista de probleme specificata prin PBIDs
 ' in categoria categID
 Sub AddPBInCateg(categID, PBIDs, objCon)
  Const SQLDelPrev = "DELETE FROM TBProblemsInCateg WHERE id_categpb=@1"
  Const SQLInsNew  = "INSERT INTO TBProblemsInCateg(id_categpb, id_problem) VALUES (@1,@2)"
  Dim PBIDAr, PBIDIt
  
  objCon.Execute Replace(SQLDelPrev, "@1", CStr(categID))
  
  If PBIDs<>"" then
    PBIDAr = Split(PBIDs, ",", -1, 1)
    For Each PBIDIt in PBIDAr
      objCon.Execute Replace(Replace(SQLInsNew, "@2", CStr(PBIDIt)), "@1", CStr(categID))
    Next
  End If
  
 End Sub

 ' Sterge categoriile specificate prin lista de ID-uri: categsID
 Sub DeleteCategs(categsID, objCon)
  Const SQLDel = "DELETE FROM TBProblemsCategories WHERE id_categpb IN (@1)"
  objCon.Execute Replace(SQLDel, "@1", categsID)
 End Sub
 
 ' Redenumeste categoria specificata prin categID cu numele newname
 Sub RenameCateg(categID, newname, objCon)
  Const SQLRen = "UPDATE TBProblemsCategories SET numecateg='@1' WHERE id_categpb=@2"
  objCon.Execute Replace(Replace(SQLRen, "@1", newname), "@2", categID)
 End Sub
 
 ' Adauga la cursul cursid categoria cu numele categname
 Sub AddCateg(cursid, categname, objCon)
  Const SQLAdd = "INSERT INTO TBProblemsCategories(numecateg, id_curs) VALUES('@1', @2)"
  objCon.Execute Replace(Replace(SQLAdd, "@2", cursid), "@1", categname)
 End Sub
%>
<body>
<script language=vbscript>
  window.parent.ReloadTDC
  msgbox <%=Mesaj%>
</script>
</body>
