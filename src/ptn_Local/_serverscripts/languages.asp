<%
' Returns the language name based on its Id
Function GetLanguageByID(lngLangID, objCon)
	Dim db
	Dim re

	Set db = New clsDatabase
	Set db.Connection = objCon
	re = db.GetScalar("langname", "GetLanguageByID", CLng(lngLangID))
	Set db = Nothing
	 
	GetLanguageByID = re
End Function
%>
