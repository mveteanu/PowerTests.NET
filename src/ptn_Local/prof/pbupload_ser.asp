<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/clsUpload.asp" -->
<%
Dim StrMessageReturn, ReturnedError
ReturnedError = ""

On Error Resume Next
Set oUpload = New clsUpload
Set openpbid   = oUpload("idopenpb")
Set delimglist = oUpload("delimglist")
Set uploadlist = oUpload("uploadlist")
Set pbtext     = oUpload("pbtext")
Set pbprop = oUpload("pbprop"): PBPropertiesArray = Split(pbprop.value,Chr(3),-1,1)
Set pbansw = oUpload("pbansw"): PBAnswersLinArray = Split(pbansw.value,vbCrLf,-1,1)

Set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
cn.BeginTrans 

DeleteServerImages delimglist.value, cn
id_problem = EnterPB(openpbid.value, PBPropertiesArray, PBAnswersLinArray, cn)
imgenteredlist = EnterImages(uploadlist.value, id_problem, cn)
DeleteOldResponses id_problem, cn
EnterResponses PBAnswersLinArray, id_problem, cn
UpdatePBText ReplaceLocalImagesNames(pbtext.value, "../prof/getattach_img.asp?FileID=", imgenteredlist), id_problem, cn

If Err.number <> 0 then
	cn.RollbackTrans 
	ReturnedError = """Error (" & Err.number & "): " & Err.Description & """, vbOkOnly + vbCritical"
	Err.Clear 
Else
	cn.CommitTrans 
	StrMessageReturn = "Question was saved successfully."
End If

cn.Close 
Set cn = Nothing

Call PrintClientResponse


' Sterge vechile raspunsuri ale problemei si le introduce pe cele noi
Sub DeleteOldResponses(idopenpb, objCon)
	Const SQLDel = "DELETE FROM TBProblemsResponses WHERE id_problem=@1"
	objCon.Execute Replace(SQLDel, "@1", idopenpb)
End Sub


' Sterge imaginile ce au ID-urile in lista CSV: servimglist
Sub DeleteServerImages(servimglist, objCon)
	Const SQLDel = "DELETE FROM TBProblemsAttachements WHERE id_problemattachement IN (@1)"
	If servimglist<>"" then 
		objCon.Execute Replace(SQLDel, "@1", servimglist)
	End If  
End Sub


' Introduce informatiile generale despre o problema si returneaza
' id-ul problemei care tocmai s-a introdus
Function EnterPB(idopenpb, PBPropArr, PBAnsLinArr, objCon)
	Dim idpb
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "TBProblems", objCon, adOpenDynamic, adLockOptimistic, adCmdTable

	If idopenpb<>"" then
		rs.Filter = "id_problem = " & CStr(idopenpb)
	else
		rs.AddNew
	end if 
	
	idpb = rs.Fields("id_problem").Value 
	rs.Fields("id_curs").Value = Session("CursID")
	if PBPropArr(0)<>"" then
		rs.Fields("numeproblema").Value  = PBPropArr(0)
	else
		rs.Fields("numeproblema").Value  = "Question " & CStr(idpb)
	end if  
	rs.Fields("autorproblema").Value = PBPropArr(1)
	rs.Fields("acceptaraspunspartial").Value = (PBPropArr(2) = "t")
	rs.Fields("tipraspuns").Value = CLng(PBAnsLinArr(0))
	rs.Update 
	rs.Close 
	set rs = nothing
	EnterPB = idpb
End Function


' Se introduc imaginile in BD
Function EnterImages(imgupldlist, idproblem, objCon)
	Dim re
 
	re = ""
	If imgupldlist<>"" then
		Set rsa = Server.CreateObject("ADODB.Recordset")
		rsa.Open "TBProblemsAttachements", objCon, adOpenDynamic, adLockOptimistic, adCmdTable
		uploadfieldsarray = Split(imgupldlist,",",-1,1)
		for each uf in uploadfieldsarray
			set oFile = oUpload(uf)
			rsa.AddNew
			rsa.Fields("id_problem").Value = idproblem
			rsa.Fields("attachementtype").Value = oFile.ContentType
			rsa.Fields("attachement").AppendChunk = oFile.BinaryData & ChrB(0)
			re = re & oFile.FileName & "|" & rsa.Fields("id_problemattachement").Value & Chr(3)
			rsa.Update
			set oFile = nothing
		next  
		rsa.Close
		set rsa = Nothing
		re = Left(re, Len(re)-Len(Chr(3)))
	End If  
 
	EnterImages = re
End Function

' Introduce raspunsurile problemei in BD
Sub EnterResponses(PBAnsLinArr, idproblem, objCon)
	Set rsans = Server.CreateObject("ADODB.Recordset")
	rsans.Open "TBProblemsResponses", objCon, adOpenDynamic, adLockOptimistic, adCmdTable
	for i=1 to UBound(PBAnsLinArr)
		rsans.AddNew
		rsans.Fields("id_problem").Value = idproblem
		select case CInt(PBAnsLinArr(0))
			case 1,2  rsans.Fields("responsecorrect").Value = CStr(PBAnsLinArr(i))
			case else PBAnswersLine = Split(PBAnsLinArr(i), Chr(2), -1, 1)
				if UBound(PBAnswersLine) > 0 then 
					rsans.Fields("responsecorrect").Value = CStr(PBAnswersLine(0))
					rsans.Fields("responsedetails").Value = CStr(PBAnswersLine(1))
				end if  
		end select
		rsans.Update 
	next
	rsans.Close
	set rsans = nothing
End Sub

' Inlocuieste numele imaginilor locale cu nume de imagini din BD
Function ReplaceLocalImagesNames(textpb, asppagename, limglist)
	Dim re
	re = textpb
  
	If limglist <> "" then
		ilst = Split(limglist, Chr(3), -1, 1)
		for each ii in ilst
			ilstitem = Split(ii, "|", -1, 1)
			re = Replace(re, ilstitem(0), asppagename & ilstitem(1))
		next
	End If

	ReplaceLocalImagesNames = re
End Function

' Se updateaza o problema proaspat introdusa in vederea
' adaugarii si a textului
Sub UpdatePBText(textpb, idproblem, objCon)
	Const SQLSel = "SELECT * FROM TBProblems WHERE id_problem=@1"
	set rsu = Server.CreateObject("ADODB.Recordset")
	rsu.Open Replace(SQLSel,"@1",CStr(idproblem)), objCon, adOpenDynamic, adLockOptimistic, adCmdText
	rsu.Fields("textproblema").Value = textpb
	rsu.Update 
	rsu.Close 
	set rsu = nothing
End Sub

' Trimite spre client un raspuns ca rezultat al salvarii datelor
Sub PrintClientResponse
	With Response
		.Write "<script language=vbscript>" & vbCrLf
		If Len(ReturnedError) = 0 Then
			.Write "msgbox """ & StrMessageReturn & """, vbOkOnly + vbInformation" & vbCrLf
			.Write "window.parent.CloseSavedProblem"  & vbCrLf
		Else
			.Write "msgbox " & ReturnedError & vbCrLf
		End If
		.Write "</script>" & vbCrLf
	End With
End Sub 
%>
