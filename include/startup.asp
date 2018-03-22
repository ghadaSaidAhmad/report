<%
dim ObjConnectionSQL
 

 
Set ObjConnectionSQL = Server.CreateObject("ADODB.Connection")

ObjConnectionSQL.Open "Provider=SQLOLEDB; Data Source =.\SQLEXPRESS; Initial Catalog = sos; User Id =ghada; Password=ghada123"

 

dim msgNoAccess
msgNoAccess = "You are not allowed to access this information"

'on error resume next	'Cuando est� en Real, se deber�a activar 
						'para que no aparezcan errores raros.

dim bottomMessage
bottomMessage = ""

dim refreshOpener
refreshOpener = false

dim showMenu
showMenu = false
%>