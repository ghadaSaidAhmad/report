<%

Function puedeSuplantar()
    
    if UCASE(Session("IDUser"))="SQLMANAGER" OR UCASE(Session("IDUser"))="JJIMENEZ" OR UCASE(Session("IDUser"))="SLOPEZ" OR UCASE(Session("IDUser"))="VLARRIBA" then
        puedeSuplantar = true
    else
        puedeSuplantar = false
    end if
    
End Function

Function getReportRows(brand)
    
    dim iSubcat, nSubcats
    nSubcats = 0
    for iSubcat = 0 to 9
        if brand.arrNShops(iSubcat) <> "" then
            nSubcats = nSubcats + 1
        end if
    next
    
    dim NFilasACTIVITY: NFilasACTIVITY = 5 + nSubcats

    dim nFilas_GeneralTheme: nFilas_GeneralTheme = 0
    if Request.Form("FILTER_SHOWGENERALTHEME")<>"" then nFilas_GeneralTheme = 1
    
    dim nFilas_RealDataNShops: nFilas_RealDataNShops = 0
    if Request.Form("FILTER_SHOWREALDATA_NSHOPS")<>"" then nFilas_RealDataNShops = 1 + nSubcats
    
    dim nFilas_TotalNShops: nFilas_TotalNShops = 0
    if Request.Form("FILTER_TOTALSHOPS")<>"" then nFilas_TotalNShops = 1 + nSubcats

    dim nFilas_KPIQuality: nFilas_KPIQuality = 0
    if Request.Form("FILTER_SHOWKPIQUALITY")<>"" then nFilas_KPIQuality = 1

    dim nFilas_Quality: nFilas_Quality = 0
    if Request.Form("FILTER_SHOWQUALITY")<>"" AND IsInputQuality() then nFilas_Quality = 2

    dim nFilas_NR: nFilas_NR = 0
    if Request.Form("FILTER_SHOWNR")<>"" then nFilas_NR = 1

    dim nFilas_FC: nFilas_FC = 0
    if Request.Form("FILTER_SHOWFC")<>"" then nFilas_FC = 1

    dim nFilas_NRVSLY: nFilas_NRVSLY = 0
    if Request.Form("FILTER_SHOWNRVSLY")<>"" then nFilas_NRVSLY = 1

    dim NFilasExtra: NFilasExtra = nFilas_RealDataNShops + nFilas_TotalNShops + nFilas_KPIQuality + nFilas_Quality + nFilas_NRVSLY + nFilas_FC + nFilas_NR + nFilas_GeneralTheme     ' Número de filas de General Theme, NR, LY, etc...
    
    getReportRows = NFilasACTIVITY + NFilasExtra
    
End Function
    

Function HexToDec(strHex)
  dim lngResult
  dim intIndex
  dim strDigit
  dim intDigit
  dim intValue

  lngResult = 0
  for intIndex = len(strHex) to 1 step -1
    strDigit = mid(strHex, intIndex, 1)
    intDigit = instr("0123456789ABCDEF", ucase(strDigit))-1
    if intDigit >= 0 then
      intValue = intDigit * (16 ^ (len(strHex)-intIndex))
      lngResult = lngResult + intValue
    else
      lngResult = 0
      intIndex = 0 ' stop the loop
    end if
  next

  HexToDec = lngResult
End Function

Function isInArray(arr, val)
    dim i, encontrado
    encontrado = false
    for each i in arr
        if CInt(i) = CInt(val) then
            encontrado = true
        end if
    next
    
    isInArray = encontrado
End Function


Sub RecoverApplication()
dim rst
dim SQL

	if Application("ApplicationCharged") = "" then
		
		'on error resume next		
		
		SQL = "SELECT * FROM TableVarApplication WHERE LoadASPVariable<>0"
		set rst = ObjConnectionSQL.Execute(SQL)
'		if Err=0 then
			while not rst.EOF
				Application(rst("Name")) = rst("VarValue")
				rst.MoveNext
			wend
			Application("ApplicationCharged") = "1"
'		end if
		rst.Close
		
	end if
	
End Sub

Function RecoverSession(redirect)
dim rst, rst2
dim SQL

	'Recover Application Variables if the table exists.
	RecoverApplication()
	
	'''''' Session.CodePage = 65001 'Multilanguage UNICODE
	'session("UserFullName")=""
	if session("UserFullName")="" then
		Set rst = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT em.ApellidosNombre, em.IDEmpleado, us.Idioma, " & _
		" CASE WHEN ug.IDGroup IS NOT NULL THEN '1' ELSE '0' END AS [isAdmin], " & _
		" CASE WHEN ug2.IDGroup IS NOT NULL THEN '1' ELSE '0' END AS [isInputData], " & _
		" CASE WHEN ug3.IDGroup IS NOT NULL THEN '1' ELSE '0' END AS [isInputQuality] " & _
		" FROM EmpleadosGlobal em " & _
		" INNER JOIN Users us ON em.IDEmpleado=us.IDEmpleado " & _
		" LEFT JOIN UserGroup ug ON ug.IDEmpleado=us.IDEmpleado AND ug.IDGroup=0 " & _
		" LEFT JOIN UserGroup ug2 ON ug2.IDEmpleado=us.IDEmpleado AND ug2.IDGroup=1 " & _
		" LEFT JOIN UserGroup ug3 ON ug3.IDEmpleado=us.IDEmpleado AND ug3.IDGroup=2 " & _
		" WHERE em.NTUser='" & session("IDUser") & "' AND em.IndBaja=0"
		
		'Temp by Abdallah to bypass authentication
		session("UserFullName") = "Abdallah"
		session("IDEmpleado") = "242"
		session("Idioma") = "EN"
		session("isAdmin") = 1
		session("isInputData") = 1
		session("isInputQuality") = 1
		RecoverSession = true
		Response.End
		'end
		
		rst.open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
		if rst.EOF then
			session("UserFullName") = ""
			session("IDEmpleado") = ""
			session("Idioma") = ""
			session("isAdmin") = ""
			session("isInputData") = ""
			session("isInputQuality") = ""
			if redirect then
				Response.Write "<script language=JavaScript>try{window.location.href='NoAccess.asp';}catch(e){}</script>"
			else
				
			end if
			Response.End


		elseif rst.RecordCount>1 then
			session("UserFullName") = ""
			session("IDEmpleado") = ""
			session("Idioma") = ""
			session("isAdmin") = ""
			session("isInputData") = ""
			session("isInputQuality") = ""
			msgError "SOASRVRF-RecoverSession-20 " & MSG_ErrorNTUserDuplicado, false, true
			RecoverSession = false
			
		else
			session("UserFullName") = rst("ApellidosNombre")
			session("IDEmpleado") = rst("IDEmpleado")
			session("Idioma") = rst("Idioma")
			session("isAdmin") = rst("isAdmin")
			session("isInputData") = rst("isInputData")
			session("isInputQuality") = rst("isInputQuality")
			RecoverSession = true

		end if
		rst.close
	else
		RecoverSession = true
	end if

	set rst = nothing

End Function


'Conecta al SQL si hace falta
Function RecoverSQLConnection()

	'Si l'obre bé, si no, també, perquè vol dir que ja estava obert.
	on error resume next
	Session("SQLConnection").Open Application("ConnectToSQL")
	RecoverSQLConnection = true
	
	set ObjConnectionSQL = Session("SQLConnection")
    Err.Clear
    
End Function

Function IsAdmin()

	if session("isAdmin") = "1" then
		IsAdmin = true
	else
		IsAdmin = false
	end if
	
End Function

Function IsInputData()

	if session("isInputData") = "1" then
		IsInputData = true
	else
		IsInputData = false
	end if
	
End Function

Function IsInputQuality()
    
    if session("isInputQuality") = "1" then
        IsInputQuality = true
    else
        IsInputQuality = false
    end if
    
End Function


'Función para escribir chivatos durante el desarrollo.
Sub PrintDebug(Variable, Value)
	if Application("DebugPrint")<>"" then
		if Value<>"" then
			Response.Write Variable & ": [" & server.HTMLEncode(Value) & "]<br>"
		else
			Response.Write Variable & ": []<br>"
		end if
	end if
End Sub

Sub PrintSQL()
	Response.Write "SQL: [" & SQL & "]<br>"
End Sub

'Devielve un mensaje de error en Cliente y 
'puede volver atrás automáticamente
Sub MsgError(strError, Back, RespEnd)
	Response.Write "<script>alert('" & replace(replace(strError,"\","\\"),"'","\'") & "');"
	if Back then
		Response.Write "history.back()"
	end if
	Response.Write "</script>"
	if RespEnd then
		Response.End
	end if
End Sub

Function ValidarFecha(ByRef fecha)
	
	if fecha<>"" then
		if len(fecha)=10 then
			if mid(fecha,3,1)="/" AND mid(fecha,6,1)="/" then
				if isdate(fecha) then
					dia = mid(fecha,1,2)
					mes = mid(fecha, 4, 2)
					anho = mid(fecha, 7)
					fecha = dateserial(anho, mes, dia)
					ValidarFecha = true
				else
					ValidarFecha = false
				end if
			else
				ValidarFecha = false
			end if
		else
			ValidarFecha = false
		end if
	else
		fecha = null
		ValidarFecha = true
	end if
	
End Function

'Comprueba que la fecha sea con el formato dd/mm/yyyy
'Devuelve 0 si es correcta y 1 si es incorrecta.
Function CheckDateFormat(thedate)

	if Len(thedate) <> 10  then
		CheckDateFormat = 1
	elseif Mid(thedate,3,1) <> "/" then
		CheckDateFormat = 1
	elseif Mid(thedate,6,1) <> "/" then
		CheckDateFormat = 1
	else
		if isDate(thedate) then
			CheckDateFormat = 0
		else
			CheckDateFormat = 1
		end if
	end if
End function



'Comprueba que el usuario tenga permiso para ver la página
'con la posibilidad de redireccionar.
Function CheckProfile(sessionName, vCheck, Redirect)
	if ((session(sessionName) AND 7) = vCheck) then
		CheckProfile = true
	else
		CheckProfile = false
		if Redirect<>"" then
			Response.Redirect Redirect
		end if
	end if
End Function


'Criterio avanzado de búsqueda.
'Escribiendo +uno +dos separa los criterior con AND/OR campo LIKE '%uno%' etc...
Function TratarCriterioAvanzado(Crit, Campo, AndOr, TipoOp, TipoDato)
dim pos1, pos2
dim strAux, Valor

	Crit = trim(Crit)
	strAux=""
	pos1 = 0
	pos1 = instr(pos1+1, Crit, "+")
	'Si encuentra un '+' en la primera posición lo trata 
	'como una consulta avanzada. Si no está en la primera
	'posición, NO.
	if TipoOp = "LIKE" then
		strOperacion = " LIKE '%"
		strOperacion2 = "%' "
	elseif TipoOp = "=" then
		if TipoDato = "Texto" then
			strOperacion = " = '"
			strOperacion2 = "' "
		else
			strOperacion = " = "
			strOperacion2 = " "
		end if
	end if
	if pos1=0 or pos1>1 then
		TratarCriterioAvanzado = " " & AndOr & "  " & Campo & strOperacion & Crit & strOperacion2
	else
		while pos1 <> 0
			pos2 = instr(pos1+1, Crit, "+")
			if pos2=0 then pos2=len(Crit)+1 end if
			Valor = trim(mid(Crit, pos1+1, pos2-pos1-1))
			strAux = strAux & " " & AndOr & "  " & Campo & strOperacion & Valor & strOperacion2
			pos1 = instr(pos1+1, Crit, "+")
		wend
		TratarCriterioAvanzado = strAux
	end if
'	Response.Write TratarCriterioAvanzado
End Function

'Escribe los puntos decimales de los números
Function FormatoNum(Num, Miles, Ndecimales)
	dim pos1, Ent, auxEnt, x, Dec, i
	
	pos1 = instr(1,Num,".")
	if pos1>0 then
		Ent = mid(Num,1,pos1-1)
	else
		Ent = Num
		pos1=len(Num)
	end if

	if Miles then
		auxEnt = ""
		x=1
		for i=len(Ent) To 0 step -1
			auxEnt = right(Ent, 1) & auxEnt
			if len(Ent)>0 then
				Ent = left(Ent,len(Ent)-1)
			end if
			if (x mod 3 = 0) AND len(Ent)>0 then
				auxEnt = "." & auxEnt
			end if
			x=x+1
		next
		Ent = auxEnt
	end if
	
	Dec = mid(Num,pos1+1,len(Num)-pos1)
	Dec = left(Dec & "0000000000", Ndecimales)
	
	if NDecimales>0 then
		FormatoNum = Ent & "," & Dec
	else
		FormatoNum = Ent
	end if
	
End Function

Function locMonthName(month, lang)
    if lang = "EN" then
        select case month
            case 1: locMonthName = "January"
            case 2: locMonthName = "February"
            case 3: locMonthName = "March"
            case 4: locMonthName = "April"
            case 5: locMonthName = "May"
            case 6: locMonthName = "June"
            case 7: locMonthName = "July"
            case 8: locMonthName = "August"
            case 9: locMonthName = "September"
            case 10: locMonthName = "October"
            case 11: locMonthName = "November"
            case 12: locMonthName = "December"
        end select
    elseif lang = "ES" then
        select case month
            case 1: locMonthName = "Enero"
            case 2: locMonthName = "Febrero"
            case 3: locMonthName = "Marzo"
            case 4: locMonthName = "Abril"
            case 5: locMonthName = "Mayo"
            case 6: locMonthName = "Junio"
            case 7: locMonthName = "Julio"
            case 8: locMonthName = "Agosto"
            case 9: locMonthName = "Septiembre"
            case 10: locMonthName = "Octubre"
            case 11: locMonthName = "Noviembre"
            case 12: locMonthName = "Diciembre"
        end select
    end if 
    
End Function

'Convierte Pesetas a EUROS
Function EuroConvert(PTA)
	
	EURO = 166.386
	EuroConvert = Round(Clng(PTA) * 100 / EURO) / 100
	
End Function

'Convierte EUROS a Pesetas
Function PTAConvert(EUR)
	
	EURO = 166.386
	PTAConvert = Round(cdbl(EUR)*EURO)
	
End Function

Function FechaAS400(Fecha)
	
	FechaAS400 = right("0" & day(Fecha),2) & right("0" & month(Fecha),2) & right(year(Fecha),4)
	
End Function

'Devuelve una fecha con el formato dd/mm/yyyy
Function FormatoFecha(Fecha)
	
	if not isnull(Fecha) and Fecha<>"" then
		if isdate(Fecha) then
			FormatoFecha = right("0" & Day(Fecha),2) & "/" & right("0" & month(Fecha),2) & "/" & year(Fecha)
		else
			FormatoFecha = ""
		end if
	else
		FormatoFecha = ""
	end if
	
End Function

Function FormatoFechaHover(Fecha)		
	'Response.Write Fecha	
	if not isnull(Fecha) and Fecha<>"" then		
		if isdate(Fecha) then						
			Fecha=CDate(Fecha)
			if Application("USDateFormatHover")="YES" then					
				'Response.Write "entra"
				FormatoFechaHover = right("0" & Month(Fecha),2) & "/" & right("0" & Day(Fecha),2) & "/" & Year(Fecha)
			else
				'Response.Write "entraNO"
				FormatoFechaHover = right("0" & Day(Fecha),2) & "/" & right("0" & Month(Fecha),2) & "/" & Year(Fecha)
			end if
		else
			FormatoFechaHover = ""
		end if
	else
		FormatoFechaHover = ""
	end if
	'Response.Write FormatoFechaHover & "ff"
End Function


Function FormatNumberView(rArgs)
	'Declare local variables
	Dim p1,p2,p3,p4
	
	valor=rArgs(0)	
	
	if isnull(valor) then
		FormatNumberView = ""
	else

		'Initialize the local variables (assign
		'them all to empty strings
		p1 = "" : p2 = "" : p3 = "" : p4 = ""
	
		Select Case UBound(rArgs)
			Case 1
				p1 = rArgs(1)	
			Case 2
				p1 = rArgs(1)
				p2 = rArgs(2)
			Case 3
				p1 = rArgs(1)
				p2 = rArgs(2)
				p3 = rArgs(3)
			Case 4
				p1 = rArgs(1)
				p2 = rArgs(2)
				p3 = rArgs(3)
				p4 = rArgs(4)
		End Select

		locale_orig=getlocale()
		strlocale=request.servervariables("HTTP_ACCEPT_LANGUAGE")
		select case strlocale
			case "es"			
				setlocale("es")						
			case "en-gb"			
				setlocale("en-gb")
			case "en-us"			
				setlocale("en-us")
			case else
				setlocale("en-gb")
		end select	
		if p1="" then
			p1=-1
		end if
		if p2="" then
			p2=-2
		end if	
		if p3="" then
			p3=-2
		end if
		if p4="" then
			p4=-2
		end if				
		valorconvert=FormatNumber(valor,p1,p2,p3,p4)	
		SetLocale(locale_orig)	
		FormatNumberView=CStr(valorconvert)
	
	end if
	
End Function

Function FormatNumberSQL(valor)  	

	if isnull(valor) then 
		valor="0"
	end if

	if not isNumeric(valor) then
		valor = "0"
	end if
	
	if valor<>"" then
		strlocale=request.servervariables("HTTP_ACCEPT_LANGUAGE")
		select case strlocale
			case "es"
				valor=replace(valor,".","")
				valor=replace(valor,",",".")						
			case "en-gb"			
				valor=replace(valor,",","")						
			case "en-us"
				valor=replace(valor,",","")						
			case else
				valor=replace(valor,",","")						
		end select
	end if
	
	FormatNumberSQL=CStr(valor)
	
End Function


%>