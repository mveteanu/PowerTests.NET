'<script language="vbscript">

' *******************************************************************
' Simple vertical menu. The implementation is based on the popup
' menu supported by IE5.5+
' Date  : March 05, 2001
' *******************************************************************

'dim menu

' Display menu helper
' x,y = display position
' w   = width of menu (pixels)
' handlername = the name of the event handler for item selected event
' menuitems   = array containing the texts of menu items
' Returns a popup type object
'
function ShowMenu( x, y, w, handlername, menuitems )
 set Menu = window.createPopup
 set Menubody = Menu.document.body
 dim i
 
 menutext="<table border=0 cellpadding=0 cellspacing=0 height=100% width=100% "&_
 "onmouseover='parent.coloreaza(event.srcElement,1)' onmouseout='parent.coloreaza(event.srcElement,2)' "&_
 "onclick='parent." & handlername & "(event.srcElement.innerHTML)' style='font-family:Tahoma;font-size:8pt;cursor:default;'>" & vbCrLf
 for each i in menuitems
   menutext = menutext + "<tr><td height=20 align=left valign=center style='padding-left:5px;padding-right:5px;'>" & i & "</td></tr>" & vbCrLf
 next
 menutext = menutext + "</table>"

 Menubody.style.backgroundColor = "buttonface"
 Menubody.style.border = "outset thin"
 Menubody.InnerHTML = menutext
 
 Menu.Show x,y,w,(1+UBound(menuitems))*20+3,window.document.body
 set ShowMenu = Menu
end function

sub coloreaza(obj,a)
   if a=1 then 
     cul  = "activecaption"
     cul2 = "captiontext"
   else 
     cul  = "buttonface"
     cul2 = "menutext"     
   end if
   if obj.innerhtml<>"<HR>" and obj.parentElement.innerhtml<>"<HR>" then
   obj.style.backgroundcolor=cul
   obj.style.color=cul2
   end if
 end sub
