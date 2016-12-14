<%@ Language=VBScript %>
<%
 Const MsgDelOK = """Selected categories were deleted."", vbOKOnly+vbInformation"
 Const MsgRenOK = """Selected category was renamed."", vbOKOnly+vbInformation"
 Const MsgAddOK = """Category was created successfully."", vbOKOnly+vbInformation"
 Const MsgEdtOK = """Category was edited."", vbOKOnly+vbInformation"
 Const MsgPubOK = """Category changed visibility attribute."", vbOKOnly+vbInformation"
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
   case "edt" AddTstInCateg RequestList, RequestValue, cn
              Mesaj = MsgEdtOK
   case "pub" ChangeCategTstVisibility Request.Form("SelectList"), Request.Form("SelectValue"), cn
              Mesaj = MsgPubOK
 End Select

 cn.Close
 Set cn = nothing

 If Err.number <> 0 then
   Err.Clear 
   Mesaj = MsgErr
 End If

 
 ' Introduce lista de teste specificata prin TstIDs
 ' in categoria categID
 Sub AddTstInCateg(categID, TstIDs, objCon)
  Const SQLDelPrev = "DELETE FROM TBTestsInCateg WHERE id_categtst=@1"
  Const SQLInsNew  = "INSERT INTO TBTestsInCateg(id_categtst, id_test) VALUES (@1,@2)"
  Dim TstIDAr, TstIDIt
  
  objCon.Execute Replace(SQLDelPrev, "@1", CStr(categID))
  
  If TstIDs<>"" then
    TstIDAr = Split(TstIDs, ",", -1, 1)
    For Each TstIDIt in TstIDAr
      objCon.Execute Replace(Replace(SQLInsNew, "@2", CStr(TstIDIt)), "@1", CStr(categID))
    Next
  End If
  
 End Sub

 ' Sterge categoriile specificate prin lista de ID-uri: categsID
 Sub DeleteCategs(categsID, objCon)
  Const SQLDel = "DELETE FROM TBTestsCategories WHERE id_categtst IN (@1)"
  objCon.Execute Replace(SQLDel, "@1", categsID)
 End Sub

 ' Modifica vizibilitatea categoriilor specificate
 Sub ChangeCategTstVisibility(categIDs, vis, objCon)
  Const SQLChPub = "UPDATE TBTestsCategories SET categtstpublica=@1 WHERE id_categtst IN (@2)"
  objCon.Execute Replace(Replace(SQLChPub, "@2", categIDs), "@1", vis)
 End Sub
 
 ' Redenumeste categoria specificata prin categID cu numele newname
 Sub RenameCateg(categID, newname, objCon)
  Const SQLRen = "UPDATE TBTestsCategories SET numecateg='@1' WHERE id_categtst=@2"
  objCon.Execute Replace(Replace(SQLRen, "@1", newname), "@2", categID)
 End Sub
 
 ' Adauga la cursul cursid categoria cu numele categname
 Sub AddCateg(cursid, categname, objCon)
  Const SQLAdd = "INSERT INTO TBTestsCategories(numecateg, id_curs) VALUES('@1', @2)"
  objCon.Execute Replace(Replace(SQLAdd, "@2", cursid), "@1", categname)
 End Sub
%>
<body>
<script language=vbscript>
  window.parent.ReloadTDC
  msgbox <%=Mesaj%>
</script>
</body>
