<%@ Language=VBScript %>
<%
Const SQLSel = "SELECT * FROM TBProblemsAttachements WHERE id_problemattachement=@1"
Dim cn, rs, fisid

fisid = Request.QueryString("FileID")
If fisid<>"" then
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")
	cn.CursorLocation = adUseClient
	cn.Open Application("DSN")
	Set rs.ActiveConnection = cn
	rs.Open Replace(SQLSel, "@1", fisid)
    If not rs.EOF then
      Response.ContentType = rs.Fields("attachementtype")
      Response.BinaryWrite rs.Fields("attachement")
    else
      PrintEndError
    end if
	rs.Close
	cn.Close
	set cn = nothing
	set rs = nothing
Else
    PrintEndError 
End If    

Sub PrintEndError
    Response.Write "<b>Error:</b> Either FileID was not specified or the file was not found in DB."	
    Response.End
End Sub
%>