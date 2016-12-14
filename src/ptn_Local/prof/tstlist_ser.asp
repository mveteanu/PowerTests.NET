<%@ Language=VBScript %>
<%
 Const MsgDelOK = """Selected tests were deleted."", vbOkOnly+vbInformation"
 Const MsgRenOK = """Selected test was renamed successfully."", vbOkOnly+vbInformation"
 Const MsgError = """DB errors occured."", vbOkOnly+vbCritical"
 Const MsgAddOK = """The test was added successfully."", vbOkOnly+vbInformation"
 Const MsgEdtOK = """Test was edited successfully."", vbOkOnly+vbInformation"
 Const MsgPubOK = """Selected tests visibility changed."", vbOkOnly+vbInformation"
 Dim Mesaj, ReqAct
 
 on error resume next
 
 set cn = server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
  
 ReqAct = Request.Form("SelectAction")
 if ReqAct="del" then
   DeleteTests Request.Form("SelectList"), cn
   Mesaj = MsgDelOK
 elseif ReqAct="ren" then
   RenameTest Request.Form("SelectList"), Request.Form("SelectValue"), cn
   Mesaj = MsgRenOK
 elseif ReqAct="add" then
   AddEdtTest CLng(Session("CursID")), "", Request.Form("SelectValue"), cn
   Mesaj = MsgAddOK
 elseif ReqAct="edt" then
   AddEdtTest CLng(Session("CursID")), Request.Form("SelectList"), Request.Form("SelectValue"), cn
   Mesaj = MsgEdtOK
 elseif ReqAct="pub" then
   ChangeTstVisibility Request.Form("SelectList"), Request.Form("SelectValue"), cn
   Mesaj = MsgPubOK
 end if
 
 cn.Close
 set cn = nothing
 
 if Err.number <> 0 then
   Mesaj = MsgError
   Err.Clear 
 end if
 
 ' Sterge testele specificate
 Sub DeleteTests(tstids, objCon)
   Const SQLDel = "DELETE FROM TBTests WHERE id_test IN (@1)"
   objCon.Execute Replace(SQLDel, "@1", tstids)
 End Sub

 ' Schimba vizibilitatea testelor
 Sub ChangeTstVisibility(tstids, vis, objCon)
   Const SQLChVis = "UPDATE TBTests SET testpublic=@1 WHERE id_test IN (@2)"
   objCon.Execute Replace(Replace(SQLChVis, "@2", tstids), "@1", vis)
 End Sub

 ' Redenumeste testul specificat
 Sub RenameTest(tstid, newname, objCon)
   Const SQLRen = "UPDATE TBTests SET numetest = '@1' WHERE id_test = @2"
   objCon.Execute Replace(Replace(SQLRen, "@2", CStr(tstid)), "@1", newname)
 End Sub

 ' Adauga/Editeaza datele unui test
 Sub AddEdtTest(CursID, TestID, PackedData, objCon)
  Dim idt
  
  idt = AddMainTest(CursID, TestID, PackedData, objCon)
  AddFilterProblems idt, PackedData, objCon
  AddFilterCategs idt, PackedData, objCon
 End Sub

 ' Adauga datele principale in tabla TBTests
 Function AddMainTest(CursID, IDOpenTest, PackedData, objCon)
  Dim UnPackedData
  Dim serNumetest, serComentarii, serTimp, serSustineri, serRandom, serPublic 
  Dim IDT, rs

  UnPackedData  = Split(PackedData, Chr(1), -1, 1)
  serNumetest   = UnPackedData(0)
  serComentarii = UnPackedData(1)
  serTimp       = CInt(UnPackedData(2))
  serSustineri  = CInt(UnPackedData(3))
  serRandom     = CInt(UnPackedData(4))
  serPublic     = CBool(UnPackedData(5))
  
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open "TBTests", objCon, adOpenDynamic, adLockOptimistic, adCmdTable
  If IDOpenTest<>"" then rs.Filter = "id_test = "  & CStr(IDOpenTest) Else rs.AddNew 

  IDT = rs.Fields("id_test").Value 
  rs.Fields("id_curs").Value        = CursID
  rs.Fields("numetest").Value       = serNumetest
  rs.Fields("comentarii").Value     = serComentarii
  rs.Fields("timp").Value           = serTimp
  rs.Fields("nrmaxsustineri").Value = serSustineri
  rs.Fields("maxrandom").Value      = serRandom
  rs.Fields("testpublic").Value     = serPublic
  
  rs.Update 
  rs.Close 
  set rs = nothing
  AddMainTest = IDT
 End Function
 
 ' Adauga filtrele referitoare la probleme
 Sub AddFilterProblems(IDOpenTest, PackedData, objCon)
  Const SQLDelOld = "DELETE FROM TBTestsFilterProblems WHERE id_test=@1"
  Const SQLInsNew = "INSERT INTO TBTestsFilterProblems(id_test, id_problem, filter) VALUES(@1, @2, @3)"
  Dim UnPackedData, it
  Dim serIncPb, serExcPb
  
  UnPackedData  = Split(PackedData, Chr(1), -1, 1)
  serIncPb      = Split(UnPackedData(6),",",-1,1)
  serExcPb      = Split(UnPackedData(7),",",-1,1)
  
  objCon.Execute Replace(SQLDelOld, "@1", CStr(IDOpenTest))
  For Each it In serIncPb
    objCon.Execute Replace(Replace(Replace(SQLInsNew, "@3", "0"), "@2", it), "@1", CStr(IDOpenTest))
  Next
  For Each it In serExcPb
    objCon.Execute Replace(Replace(Replace(SQLInsNew, "@3", "1"), "@2", it), "@1", CStr(IDOpenTest))
  Next
 End Sub

 ' Adauga filtrele referitoare la categoriile de probleme
 Sub AddFilterCategs(IDOpenTest, PackedData, objCon)
  Const SQLDelOld = "DELETE FROM TBTestsFilterCateg WHERE id_test=@1"
  Const SQLInsNew = "INSERT INTO TBTestsFilterCateg(id_test, id_categpb, filter, nrproblems) VALUES(@4, @1, @2, @3)"
  Dim UnPackedData
  Dim serCatFilt, serCatFilt1, it
  
  UnPackedData  = Split(PackedData, Chr(1), -1, 1)
  serCatFilt    = Split(UnPackedData(8),"|",-1,1)
  
  objCon.Execute Replace(SQLDelOld, "@1", CStr(IDOpenTest))
  For Each it In serCatFilt
    serCatFilt1 = Split(it, ",", -1, 1)
    objCon.Execute Replace(Replace(Replace(Replace(SQLInsNew, "@3", serCatFilt1(2)), "@2", serCatFilt1(1)), "@1", serCatFilt1(0)), "@4", CStr(IDOpenTest))
  Next
 End Sub
%>
<body>
<script language=vbscript>
  window.parent.ReloadTDC
  msgbox <%=Mesaj%>
</script>
</body>
