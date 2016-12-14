<%
' Converteste o valoare de tip Data calendaristica intr-un string
' folosind formatul sFormat dat sub forma: "MM/DD/YYYY"
Private Function DToSR(ByVal dDate, sFormat)
	Dim sBuf
	
	If TypeName(dDate) = "Date" then
		sBuf = Replace(sFormat,"MM",Right("0" & Month(dDate),2))
		sBuf = Replace(sBuf, "DD", Right("0" & Day(dDate),2))
		sBuf = Replace(sBuf, "YYYY", Year(dDate))
	End If
	DToSR = sBuf
End Function

' Functia intoarce parametrul al doilea daca primul este Null
' sau pe primul in cazul in care acesta nu e null
' Functionarea este asemanatoare cu cea a functiei omonime din Access
Function Nz(ValueToTest, ValueIfNull)
	If IsNull(ValueToTest) Then
		Nz = ValueIfNull
	Else
		Nz = ValueToTest
	End If
End Function

' Converteste o valoare nula in numarul "0"
Function NullToZero(ByVal n)
	NullToZero = Nz(n, 0)
End Function

' Converteste un sir vid intr-un caracter "Space"
Function VidToSpace(ByVal n)
  if n="" then
    VidToSpace = " "
  else
    VidToSpace = n
  end if
End Function

' Converteste o valoare booleana intr-un sir de tipul "Da/Nu"
Function BooleanToAfirm(ByVal n)
  if n=true then
    BooleanToAfirm = "Yes"
  else
    BooleanToAfirm = "No"
  end if
End Function

' Converteste o valoare booleana intr-un sir de tipul "Da/Nu"
' formatat HTML cu atribute de culoare
Function BooleanToAfirmColor(ByVal n, strYesColor, strNoColor)
  if n=true then
    BooleanToAfirmColor = "<font color=" & strYesColor & ">Yes</font>"
  else
    BooleanToAfirmColor = "<font color=" & strNoColor & ">No</font>"
  end if
End Function



' Functia compara n1 cu n2 si il intoarce pe n1 in culoarea 
' specificata de strCul1 daca este adevarata relatia de 
' comparatie strComp intre n1 si n2, altfel il intoarce in
' culoarea strCul2. Daca culorile nu se specifica atunci
' numerele se intorc necolorate
Function IntToColor(ByVal n1, n2, strComp, strCul1, strCul2)
Dim ret1, ret2

if strCul1="" then 
  ret1 = CStr(n1)
else
  ret1 = "<font color=" & strCul1 & ">" & CStr(n1) & "</font>"
end if  

if strCul2="" then 
  ret2 = CStr(n1)
else
  ret2 = "<font color=" & strCul2 & ">" & CStr(n1) & "</font>"
end if  


Select case strComp
  case "="  if n1 = n2 then 
                 IntToColor = ret1
            else
                 IntToColor = ret2
            end if     
  case "<>" if n1 <> n2 then 
                 IntToColor = ret1
            else
                 IntToColor = ret2
            end if     
  case ">" if n1 > n2 then 
                 IntToColor = ret1
            else
                 IntToColor = ret2
            end if     
  case "<" if n1 < n2 then 
                 IntToColor = ret1
            else
                 IntToColor = ret2
            end if     
End Select
End Function



' Functia intoarce sirul str1 daca n1 si n2 sunt egale,
' iar in caz contrar intoarce sirul str2
Function CompareToString(n1, n2, str1, str2)
 if n1 = n2 then
   CompareToString = str1
 else
   CompareToString = str2
 end if    
End Function


' Functia intoarce sirul stri daca n1 si n2 sunt egale,
' iar in caz contrar il intoarce pe n1
Function CompareToString2(n1, n2, stri)
 If (IsNull(n1) and IsNull(n2)) then
   CompareToString2 = stri
 Else
   If n1 = n2 Then
	  CompareToString2 = stri
	Else
	  CompareToString2 = CStr(n1)
	End If
 End If
End Function


 ' Converteste codul permisiunii in string
 Function PermisionToString(ByVal perm)
   select case perm
     case 0 PermisionToString = "Automatic enrollment"
     case 1 PermisionToString = "Needs validation"
     case 2 PermisionToString = "Automatic deny"
   end select
 End Function

 ' Converteste nr. de studenti intr-un string
 Function NrStudToString(ByVal nrs, color)
   if nrs=0 then
     if color<>"" then
       NrStudToString = "<font color="& color &">Unlimited</font>"
     else
       NrStudToString = "Unlimited"
     end if  
   else   
     NrStudToString = CStr(nrs)
   end if  
 End Function

' Converteste textul text intr-unul de lungime nr.
' Daca e prea mare ia doar primele nr. caractere
' Daca e prea lung umple spatiul liber cu caractere fillstr
Function GetFixText(text, nr, fillstr)
 Dim re
 Dim lentext, i
 
 lentext = Len(text)
 If nr<lentext then 
  re = Left(text,nr)
 Else
  re = text
  for i = 1 to nr-lentext
   re = re & fillstr
  next
 End If 
 GetFixText = re
End Function


' Intoarce diferenta a 2 dictionare, adica acele elemente
' din objBigD a caror keie nu se gaseste in objLittleD
Function GetDictDifference(objBigD, objLittleD)
 Dim re, it
 
 Set re = CreateObject("Scripting.Dictionary")
 For each it in objBigD.keys
  If not objLittleD.Exists(it) then re.Add it, objBigD.item(it)
 Next
 
 Set GetDictDifference = re
End Function

' Intoarce suma itemilor dintr-un array de dictionare
' Daca itemii au aceeasi keye ei se introduc o singura data
' in dictionarul suma
Function GetDictSum(DictArr)
 Dim re, it1, it2
 Set re = CreateObject("Scripting.Dictionary")
 For each it1 in DictArr
  For each it2 in it1.keys
   If not re.Exists(it2) then re.Add it2, it1.item(it2)
  Next
 Next 
 Set GetDictSum = re
End Function

' Creaza un Dictionar din 2 arrayuri. 
' Elementele primului vor forma keyile 
' iar ale celui de al doilea item-ii propriuzisi
Function DictFromTwoArrays(Ar1, Ar2)
 Dim re, i
 
 Set re = CreateObject("Scripting.Dictionary")
 If UBound(Ar1) = UBound(Ar2) then
   For i=0 to UBound(Ar1)
    re.Add CStr(Ar1(i)), CStr(Ar2(i))
   Next 
 End If
 Set DictFromTwoArrays = re
End Function

 ' Obtine un string in format CSV din keyle unui dictionar
 Function DictkeysToCSVList(dict)
  Dim re, k
  re = ""
  For Each k in dict.Keys
   re = re & CStr(k) & ","
  Next
  If re <> "" then re = Left(re, Len(re)-Len(","))
  DictkeysToCSVList = re
 End Function

' Impacheteaza un dictionar intr-un string
Function PackDictToString(dict,separat1, separat2)
  Dim re, it
  re = ""
  For Each it in dict.Keys
   re = re & it & separat1 & dict.Item(it) & separat2
  Next
  If re <> "" then re = Left(re, Len(re)-Len(separat2))
  PackDictToString = re
End Function

' Despacheteaza intr-un dictionar un string
Function UnpackDictFromString(stri,separat1, separat2)
  Dim re, ar1, ar2, it
  Set re = CreateObject("Scripting.Dictionary")
  ar1 = Split(stri, separat2, -1, 1)
  For Each it in ar1
   ar2 = Split(it, separat1, -1, 1)
   re.Add ar2(0), ar2(1)
  Next
  Set UnpackDictFromString = re
End Function


' Obtine separatorul zecimal folosit de sistem
Function GetDecimalSeparator
 If CDbl("1.1") = 1.1 then GetDecimalSeparator = "." else GetDecimalSeparator = ","
End Function

' Intoarce un sir valid nr. real din sirul s prin inlocuirea
' caracterelor . si , cu caracterul separator zecimal
Function GetRegionalDouble(s)
 Dim DSep
 DSep = GetDecimalSeparator
 GetRegionalDouble = Replace(Replace(s, ".", DSep), ",", DSep)
End Function

' Intoarce un sir din care elimina toate caracterele nonnumerice
Function ExtractNumeric(s)
 Dim DSep
 Dim re, i, c, sw, foundvirgula

 DSep = GetDecimalSeparator
 foundvirgula = false
 re = ""
 sw = Replace(Replace(Trim(s), ".", DSep), ",", DSep)
 For i=1 to Len(sw)
  c = Mid(sw,i,1)
  If (c = "-") and (re = "") then
    re = re & c
  ElseIf (c = DSep) and (not foundvirgula) then
    re = re & c
    foundvirgula = true
  ElseIf IsNumeric(c) then
    re = re & c  
  End If
 Next
 ExtractNumeric = re
End Function

' Elimina tagurile HTML din string-uri si le
' truncheaza la marimea maxima specificata
Function RemoveTags(strString, intMaxLength)
	Dim s
	
	s = Server.HTMLEncode(strString)
	If intMaxLength > 0 Then
		If Len(s) > intMaxLength Then
			s = Left(s, intMaxLength)
		End If
	End If
	
	RemoveTags = s	
End Function
%>
