<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/clsDatabase.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1

	Dim db
	Dim rs
	Dim w, idc
	
	w = CLng(Request.QueryString("w").Item)
	idc = Request.QueryString("idc").Item 
	Set db = New clsDatabase

	Select Case w
		Case 1
			Set rs = db.GetRs("SELECT id, name FROM TBTextsPages", Empty)
			With Response
				.Write Replace("id|name", "|", Chr(1)) & vbCrLf
				Do Until rs.EOF
					.Write rs.fields("id").Value & Chr(1)
					.Write rs.fields("name").Value & vbCrLf
					rs.MoveNext
				Loop
			End With
			Set rs = Nothing
		Case 2
			Set rs = db.GetRs("SELECT id, id_page, controlname FROM TBTextsControls", Empty)
			With Response
				.Write Replace("id|id_page|controlname", "|", Chr(1)) & vbCrLf
				Do Until rs.EOF
					.Write rs.fields("id").Value & Chr(1)
					.Write rs.fields("id_page").Value & Chr(1)
					.Write rs.fields("controlname").Value & vbCrLf
					rs.MoveNext
				Loop
			End With
			Set rs = Nothing
		Case 3
			Set rs = db.GetRs("SELECT id, id_lang, textvalue FROM TBTexts WHERE id_control = " & idc, Empty)
			With Response
				.Write Replace("id|id_lang|textvalue", "|", Chr(1)) & vbCrLf
				Do Until rs.EOF
					.Write rs.fields("id").Value & Chr(1)
					.Write rs.fields("id_lang").Value & Chr(1)
					.Write rs.fields("textvalue").Value & vbCrLf
					rs.MoveNext
				Loop
			End With
			Set rs = Nothing
	End Select


	Set db = Nothing
%>
