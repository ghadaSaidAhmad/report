<%
Sub SendSQLMail(Recipients, CC, Subject, Body)
dim res

	'Quitamos espacios en blanco de las direcciones de correo.
	Recipients=Replace(Recipients," ","")
	
	set ObjConnection12 = CreateObject("ADODB.Connection")
	
	on error resume next
    Body = "<div style='font-size:12px;color:#000;font-family:Arial,Verdana;'>" & Body & "</div>"
	ObjConnection12.Open Application("ConnectToSQLMail")
	SQL = "EXECUTE master.dbo.sp_smtp_sendmail @recipients='" & Recipients & "', @copy_recipients='" & CC & "', @subject='" & replace(Subject,"'","''") & "', @message='" & replace(Body,"'","''") & "'"
	res = ObjConnection12.Execute(SQL)
	ObjConnection12.Close
	if Err<>0 then
		'Error connecting to SQLMail Server
		exit sub
	end if
	on error goto 0

	set ObjConnection12 = nothing

End Sub
%>