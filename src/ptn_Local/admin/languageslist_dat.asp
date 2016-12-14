<%@ Language=VBScript %>
<!-- #include file="../_serverscripts/clsDatabase.asp" -->
<!-- #include file="../_serverscripts/settings.asp" -->
<%
	Response.Buffer = True
	Response.Expires = -1

	Dim db
	Dim rs
	Dim id
	Dim preflang
	
	Set db = new clsDatabase
	preflang = GetPreferredLanguage(db.Connection)
	
	Set rs = db.GetRS("GetLanguages", Empty)
	With Response
		.Write "id|Language|Users|bpreferata|Default"& vbCrLf
		Do Until rs.EOF
			id = CLng(rs.fields("id").value)	
			.Write id & "|"
			.Write rs.fields("langname").value & "|"
			.Write rs.fields("users").value & "|"
			If id = preflang Then 
				.Write CInt(true) & "|Da" 
			Else 
				.Write CInt(false) & "|Nu"
			End If
			.Write vbCrLf
			rs.MoveNext
		Loop
	End With
	
	Set cn = Nothing
%>