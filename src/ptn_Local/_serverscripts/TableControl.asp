<!-- #include file="ControlUtils.asp" -->
<%
' *********************************************************************
' TTableGrid control server side support
' Date:  March 15, 2001
' *********************************************************************

' ***************************************************
' BEGIN
' PUBLIC
' SECTION
' ***************************************************

' Main sub for creating the grid plus the associated TDC
' Input:
'   iLeft   \ iLeft and iTop are used for absolute contol positioning
'   iTop    / if missing relative positioning will be used
'   iHeight = control height (width is automatically determined based on the width of all columns)
'   strTabNames = Array with column names (MUST be the same with the ones from TDC datasource !!!)
'   iTabSizes = Array with columns width. Should be the same size as above array !!!
'   iTipTable = Grid type (0=static, 1=single selection, 2=multiple selection)
'   tdcURL = TDC DataSource Url
'   bAllowSorting = If true - is allowed client side datagrid sorting
'   strComponentName = Component name
' TDC DataSource format: 
'     id|nume|prenume|email
'     1|Veteanu|Marian|mveteanu@yahoo.com
'     etc.
'   ID field is mandatory if iTipTable=1 or 2
Public Sub CreateTableControl(iLeft, iTop, iHeight, strTabNames, iTabSizes, iTipTable, tdcURL, bAllowSorting, strComponentName)
  Dim i

  ControlName = strComponentName
  ControlLeft = iLeft
  ControlTop = iTop
  ControlWidth = 0
  ControlHeight = iHeight
  ControlTabNames = strTabNames
  ControlTabSizes = iTabSizes
  ControlAllowSorting = bAllowSorting
  
  for each i in iTabSizes 
    ControlWidth = ControlWidth + i
  next
  ControlWidth = ControlWidth
  
  call AddTDC("tdc" & ControlName, "|", tdcURL)
  
  call OpenCanvas("table" & strComponentName, ControlLeft, ControlTop, ControlWidth, ControlHeight)
  for i = 0 to UBound(ControlTabNames)
    call CreateTableHeaderTab(i)
  next
  
  call CreateTableBody(iTipTable)
  call CloseCanvas
End Sub


'
' BEGIN PRIVATE SECTION
'

Dim ControlName
Dim ControlLeft
Dim ControlTop
Dim ControlWidth
Dim ControlHeight
Dim ControlTabNames
Dim ControlTabSizes
Dim ControlAllowSorting


' Creates the header of a column
Private Sub CreateTableHeaderTab(iTab)
  Dim DIVTemplate
  Dim itabLeft
  Dim i
  
  DIVTemplate = "<DIV id='@1' @6 unselectable='on' style='POSITION: absolute; OVERFLOW: hidden; CURSOR: @7; BORDER: outset thin; LEFT:@2px; TOP:0px; WIDTH:@3px; HEIGHT:20px; BACKGROUND-COLOR: buttonface; FONT-FAMILY:verdana; FONT-SIZE:8pt; FONT-WEIGHT:bolder; PADDING:1px; TEXT-ALIGN:left;'>@5@4</DIV>"

  iTabLeft = 0
  for i=0 to iTab-1
    iTabLeft = iTabLeft + ControlTabSizes(i)
  next
  
  DIVTemplate = Replace(DIVTemplate,"@1","tab" & CvtStrToID(ControlTabNames(iTab)))
  DIVTemplate = Replace(DIVTemplate,"@2",CStr(iTabLeft))
  DIVTemplate = Replace(DIVTemplate,"@3",CStr(ControlTabSizes(iTab)))
  if ControlAllowSorting then
   DIVTemplate = Replace(DIVTemplate,"@4","<span unselectable='on' style='position:absolute;top:2px;'>"& ControlTabNames(iTab)& "</span>")
   DIVTemplate = Replace(DIVTemplate,"@5","<span unselectable='on' style='display:none;font-family:webdings;'>6</span>")
   DIVTemplate = Replace(DIVTemplate,"@6","onclick='vbscript:HandleTableHeaderSort tdc"& ControlName &", "& iTab &", "& UBound(ControlTabNames) &"'")
   DIVTemplate = Replace(DIVTemplate,"@7","hand")
  else
   DIVTemplate = Replace(DIVTemplate,"@4",ControlTabNames(iTab))
   DIVTemplate = Replace(DIVTemplate,"@5","")
   DIVTemplate = Replace(DIVTemplate,"@6","")
   DIVTemplate = Replace(DIVTemplate,"@7","default")
  end if
  Response.Write DIVTemplate & vbCrLf
End Sub



' Creates the main grid part (autoscroll DIV and the TDC binded TABLE)
Private Sub CreateTableBody(tipcontrol)
  Dim DIVTemplate1
  Dim DIVTemplate2
  Dim TableTemplate1
  Dim TableTemplate2
  Dim TableTemplate3
  Dim TipControlStr
  Dim i
  
  DIVTemplate1 = "<DIV unselectable='on' style='BORDER: outset thin; POSITION: absolute; LEFT:0px; TOP:20px; WIDTH:@1px; HEIGHT:@2px; OVERFLOW:auto;'>"
  DIVTemplate2 = "</DIV>"
  DIVTemplate1 = Replace(DIVTemplate1,"@1",ControlWidth)
  DIVTemplate1 = Replace(DIVTemplate1,"@2",ControlHeight-20)
  
  select case tipcontrol
    case 0 TipControlStr = "style='cursor:default;'"
    case 1 TipControlStr = "style='cursor:hand;' onclick='vbscript:HandleTableClick(1)'"
    case 2 TipControlStr = "style='cursor:hand;' onclick='vbscript:HandleTableClick(2)'"
  end select
  
  TableTemplate1 = "<TABLE "& TipControlStr &" id='tbl"& ControlName &"' datasrc=#tdc"& ControlName &" class='TTableGrid' border=0 width=100% bgcolor=white cellspacing=0 cellpadding=0>" &vbcrlf &_
                   "<TBODY>" &vbcrlf &_
                   "<TR>" &vbcrlf
 
  TableTemplate2 = "</TR>" &vbcrlf &_
                   "</TBODY>" &vbcrlf &_
                   "</TABLE>" &vbcrlf
  
  TableTemplate3 = "<TD align=left valign=center width="& CStr(ControlTabSizes(0)-2) &" height=20 class='TTableRowUnSelected'><span datafld='id' style='display:none;'></span><span unselectable='on' DATAFORMATAS=HTML style='overflow:hidden;width:"& CStr(ControlTabSizes(i)-4) &"px;' datafld='"& ControlTabNames(0) &"'></span></TD>" &vbcrlf
  For i = 1 to UBound(ControlTabNames)
	 If i = UBound(ControlTabNames) Then
		rw = ""
		rw2 = ""
	 Else
		rw = "width=" & CStr(ControlTabSizes(i)-2)
		rw2 = "width:" & CStr(ControlTabSizes(i)-4) & "px;"
	 End If
	 TableTemplate3 = TableTemplate3 & "<TD align=left valign=center "& rw &" height=20 class='TTableRowUnSelected'><span unselectable='on' DATAFORMATAS=HTML style='overflow:hidden;" & rw2 & "' datafld='"& ControlTabNames(i) &"'></span></TD>" &vbcrlf  
  Next
  
  Response.Write DIVTemplate1 & vbCrLf
  Response.Write TableTemplate1
  Response.Write TableTemplate3
  Response.Write TableTemplate2
  Response.Write DIVTemplate2 & vbCrLf
End Sub


' Creates the main DIV where the other elements will be places
Private Sub OpenCanvas(strCanvasId, strLeft, strTop, strWidth, strHeight)
  Dim strStyleDef
  Dim strTag
  
  if not(IsNumeric(strLeft) and IsNumeric(strTop)) then
     strStyleDef = "DISPLAY:inline;position:relative;"
  else
     strStyleDef = "POSITION:absolute;LEFT:"+CStr(strLeft)+"px; TOP:"+CStr(strTop)+"px;"
  end if   
  strStyleDef = strStyleDef + " WIDTH:"+CStr(strWidth)+"px; HEIGHT:"+CStr(strHeight)+"px;"
  strStyleDef = strStyleDef +"Z-INDEX:0; background-color:buttonface; OVERFLOW:hidden;"
  
  strTag = "<DIV unselectable='on'"
  strTag = strTag + " id="+strCanvasId
  strTag = strTag + " name="+strCanvasId
  strTag = strTag + " style='"+strStyleDef+"'"
  strTag = strTag + ">" + vbCrLF+vbCrLF
  
  Response.Write strTag
End Sub


' Close main DIV
Private Sub CloseCanvas
  Response.Write (vbCrLf+"</DIV>"+vbCrLf)
End Sub


' Make an ID from a column name
Private Function CvtStrToID(strTabName)
  CvtStrToID = Replace(Replace(Replace(Replace(strTabName,",","_")," ","_"),"(","_"),")","_")
End Function
%>

