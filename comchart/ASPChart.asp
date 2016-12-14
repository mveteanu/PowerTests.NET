<% @Language = "VBScript" %>
<% 
 Set ch = Server.CreateObject("VMAObjects.ASPChart") 
 ch.DefineCanvas "Site-uri web VMA soft", 250,180
 ch.AddBar "VMA soft",10
 ch.AddBar "Psihoteste", 23
 ch.AddBar "Windows", 0
 ch.GenerateChart 
%>

