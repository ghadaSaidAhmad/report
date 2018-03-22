<%
Function PaginarGR (iRegsPorPag, iPag, vector, Fields, Accion, Tipo, vectorTipos, TableWidth, vectorWidth)
'Esta función realiza la paginación y la presentación de los datos en una tabla.
'iRegsPorPag: Numero de Registro por página.
'ipag: Página a presentar.
'Vector: array de datos seleccionados.
'Fields: Campos a presentar en la cabecera
'Accion: Opcion de realizar algo con el registro en cuestión.
'	Accion="100" --> Botón Edit
'	Accion="010" --> Botón Delete
'	Accion="001" --> Botón New
'	Accion="000" --> Nada
'	Y mezclados...
'	Accion="111" --> Todo
'TableWidth: Ancho de la tabla de resultados

'I, J se utilizan para recorrer el vector
Dim I, J 
'Total de páginas y la página que queremos mostrar
Dim iPaginas, iPagActual
'Total de registros, registro en que empezamos y registro en que terminamos
Dim iTotal, iComienzo, iFin
Dim rst, QUERY, QUERY_SEARCH

dim fecha


set rst = CreateObject("ADODB.RecordSet")


'Eliminamos el Order by de la cadena Query
IF INSTR(1,QUERY,"ORDER",1)<>0 THEN
	QUERY_SEARCH = MID(QUERY,1,INSTR(1,QUERY,"ORDER",1)-1)
ELSE
	QUERY_SEARCH = QUERY
END IF

if not IsArray(vectorWidth) then
	vectorWidth = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
end if
if not IsArray(vectorTipos) then
	vectorTipos = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
end if

'Hallo el total de registros devueltos
iTotal = UBound(aDatos,2)+1
'Calculo el numero de páginas que tenemos
iPaginas = (iTotal \ iRegsPorPag)
'Si daba decimales, añado una más
'para mostrar los últimos registros
if iTotal mod iRegsPorPag > 0 then
	iPaginas = iPaginas + 1
end if
'Si no es una página válida, comienzo en la primera
if iPag < 1 then
	iPag = 1
end if
'Si es una página mayor al nº de páginas, comienzo en la última
if iPag > iPaginas then
	iPag = iPaginas
end if

Response.Write "<table width=""" & TableWidth & """ ><tr><td>"
Response.Write("<font face=Verdana size=1>" & IDM_ListPage & " " & iPag & " " & IDM_ListPageOf & " " & iPaginas & " (" & iTotal & " " & IDM_ListPageRecords & ")</font><br>")
Response.Write "</td>"

'Imprimo enlaces, si son necesarios
if iPag > 1 then
	Response.Write "<TD width=40 align=center class=""listNavON"">"
	Response.Write "<A onclick=""_fireEvent('Primero','click',''); return false;"" href="""">" & IDM_ListFirst & "</A>"
	Response.Write "</TD>"

	Response.Write "<TD width=40 align=center class=""listNavON"">"
	Response.Write("<A onclick=""_fireEvent('Anterior','click','');return false;"" href="""">" & IDM_ListPrevious & "</A>")
	Response.Write "</TD>"
else
	Response.Write "<TD width=40 align=center class=""listNavOFF"">" & IDM_ListFirst & "</TD>"

	Response.Write "<TD width=40 align=center class=""listNavOFF"">" & IDM_ListPrevious & "</TD>"
end if
if iPag < iPaginas then
	Response.Write "<TD width=40 align=center class=""listNavON"">"
	Response.Write("<A onclick=""_fireEvent('Siguiente','click','');return false;"" href="""">" & IDM_ListNext & "</A>")
	Response.Write "</TD>"
	Response.Write "<TD width=40 align=center class=""listNavON"">"
	Response.Write "<A onclick=""_fireEvent('Ultimo','click',''); return false;"" href="""">" & IDM_ListLast & "</A>"
	Response.Write "</TD>"
else
	Response.Write "<TD width=40 align=center class=""listNavOFF"">" & IDM_ListNext & "</TD>"
	Response.Write "<TD width=40 align=center class=""listNavOFF"">" & IDM_ListLast & "</TD>"
end if

Response.Write "</tr></table>"

'Calculo el índice donde comienzo:
iComienzo = (iPag-1)*iRegsPorPag 
'y donde termino:
iFin = iComienzo + (iRegsPorPag-1)
'Si no tengo suficientes registros restantes,
'voy hasta el final
if iFin > UBound(vector, 2) then 
	iFin = UBound(vector, 2)
end if
'Pinto la tabla
Response.Write("<TABLE BORDER=""0"" width=""" & TableWidth & """ cellpadding=2 cellspacing=0>")
Response.Write("<tr class=""TDListTitle"" >")
if Accion<>"0000" then
	Response.Write("<TD align =""left"" width=50 valign=""top""><font face=""Verdana,Arial"" size=""1""><b>" & FLD_OpcionLista & "</b></font></TD>")
else
	Response.Write "<td></td>"
end if
dim p
FOR p=0 TO Fields.count -1
	if request("order")=Fields(p).name then 'and request("Srch_Order")="" then
		Response.Write("<td width=" & vectorWidth(p) & " align =""left"" valign=""top"">" & "<A onclick=""ChangeOrder('" & Fields(p).name & "');return false;"" href=""""> <B><font class=listColumn> " & Fields(p).name & "&nbsp;</font></B></A><img src=images/small_arrow_down.gif align=middle></td>")
	else
		Response.Write("<td width=" & vectorWidth(p) & " align =""left"" valign=""top"">" & "<A onclick=""ChangeOrder('" & Fields(p).name & "');return false;"" href=""""> <B><font class=listColumn> " & Fields(p).name & "</font></B></A></td>")
	end if
NEXT


Response.Write("<TR>")
strSql=""
dim permiso, strAcciones
permiso = accion
For I= iComienzo To iFin
	'Response.Write("<tr><TD>" & accion & "</td></tr>")
	'Response.Write("<tr><TD>" & accion & "</td></tr>")
	Response.Write("<TR>")
	Response.Write "<TD align=left width=60>"
	if Accion<>"0000" then
		strAcciones = ""
		if mid(Accion,1,1) = "1" then
			strAcciones = "<a href="""" onclick=""editElement('" & vector(0,I) & "', '" & Tipo & "');return false;""><img src=""images/edit.png"" border=0 alt=""" & FLD_OpcionEdit & """></a>&nbsp;&nbsp;&nbsp;"
		end if
		if mid(Accion,2,1) = "1" then
			strAcciones = strAcciones & "<a href="""" onclick=""DeleteElement('" & vector(0,I) & "', '" & Tipo & "');return false;""><img src=""images/delete.png"" border=0 alt=""" & FLD_OpcionBorrar & """></a>&nbsp;&nbsp;&nbsp;"
		end if
		if mid(Accion,3,1) = "1" then
			strAcciones = strAcciones & "<a href="""" onclick=""copyElement('" & vector(0,I) & "', '" & Tipo & "');return false;""><img src=""images/copy.png"" border=0 alt=""" & FLD_OpcionCopiar & """></a>&nbsp;&nbsp;"
		end if
		if mid(Accion,4,1) = "1" then
			strAcciones = strAcciones & "<a href="""" onclick=""editElement('" & vector(0,I) & "', '" & Tipo & "');return false;""><img src=""images/iEye.gif"" border=0 alt=""View Details""></a>&nbsp;"
		end if
		Response.Write strAcciones
	end if
	Response.Write "</TD>"
	For J=0 To UBound(vector,1)
		Response.Write "<TD><font face=Verdana size=1>"
		if vectorTipos(J)<>"" then
			if instr(1, vectorTipos(J),"ID") then		'le pasan el tipo ID(n) donde n es el 
				pos = instr(1,vectorTipos(J),"ID")+3	'número de dígitos (pone ceros a la izquierda)
				pos2 = instr(pos+1,vectorTipos(J),")")
				nchars = mid(vectorTipos(J),pos,pos2-pos)
				str0 = ""
				for x = 0 to nchars
					str0 = str0 & "0"
				next
				Response.Write right(str0 & vector(J,I),cint(nchars))
			elseif instr(1, vectorTipos(J),"fechaAS") then	'transforma una fecha de AS400 a fecha normal.
				fecha = right("0" & vector(J,I),8)
				fecha = mid(fecha,1,2) & "/" & mid(fecha,3,2) & "/" & mid(fecha,5,4)
				Response.Write fecha
			elseif instr(1, vectorTipos(J),"NUM") then	'le pasan el tipo NUM(n)
				pos = instr(1,vectorTipos(J),"NUM")+4	'donde n es el número de decimales
				pos2 = instr(pos+1,vectorTipos(J),")")
				nchars = mid(vectorTipos(J),pos,pos2-pos)
				Response.Write FormatNumber(vector(J,I),nchars)
			elseif instr(1, vectorTipos(J),"EmpleadoAS") then
				SQL = "SELECT ApellidosNombre FROM EmpleadosGlobal WHERE NEmpleado=" & vector(J,I)
				rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
				if not rst.EOF then
					Response.Write rst("ApellidosNombre")
				else
					Response.Write "Unknown"
				end if
				rst.Close
				
			else
				Response.Write vector(J,I)
			end if
		else
			Response.Write vector(J,I)
		end if
		'Response.Write vector(J,I)
		Response.Write "</font></TD>"
	Next
next
Response.Write("</TABLE>")
Response.Write "</font>"


PaginarGR = 0
End Function

'-------------------------------------------------------------

'--------------------------------------------------------------


'-------------------------------------------------------------
'-------------------------------------------------------------
'Sub Siguiente página
Sub Siguiente_click()
	if Pagina = "" then
		Pagina = "1"
	else
		Pagina = Pagina + 1
	end if
End Sub

'-------------------------------------------------------------
'-------------------------------------------------------------
'Sub Página Anterior
Sub Anterior_click()
	if Pagina = "" then
		Pagina = "1"
	else
		Pagina = Pagina - 1
	end if
End Sub

'--------------------------------------------------------
'--------------------------------------------------------
Sub Primero_click()
	Pagina = 1
End Sub

'--------------------------------------------------------
'--------------------------------------------------------
Sub Ultimo_click()
	Pagina = -1
End Sub

select case EventObject
	case "Siguiente"
		Siguiente_click()
	case "Anterior"
		Anterior_click()
	case "Primero"
		Primero_click()
	case "Ultimo"
		Ultimo_click()
end select

%>

<script language="JavaScript">
function ChangeOrder(Field){
	document.thisForm.Order.value = Field;
	document.thisForm.submit();
}
</script>
