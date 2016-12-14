'<script language="vbscript">

' Convert a string of type nnpx to integer value nn
' Example: 27px -> 27
Function StyleSizeToInt(ssize)
  StyleSizeToInt = CInt(Left(ssize,Len(ssize)-2))
End Function

