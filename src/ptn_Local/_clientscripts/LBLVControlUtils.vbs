'<script language="vbscript">

' Splits fullstr using splitstr and returns an array
' that DOESN'T contains empty eleemnts
Function NonvidSplit(fullstr, splitstr)
 Dim ar1, re()
 Dim l
 
 ar1 = Split(fullstr, splitstr, -1, 1)
 l = 0
 for each arit in ar1
  If arit<>"" then 
   Redim Preserve re(l)
   re(l) = arit
   l = l + 1
  End If
 next
 NonvidSplit = re
End Function


' Make the string text fixed size
' If it's bigger than only the first characters are used
' If it's shorter the text is padded using fillstr chars
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
