<%@language=VBScript%>
<%
    Response.Expires=0
    Response.CharSet="iso-8859-1"   
    Response.Expires = -1
    Response.ExpiresAbsolute = #1/1/2000 00:01:00#
    Response.AddHeader "pragma", "no-cache"
    Response.AddHeader "cache-control","private"
    Response.CacheControl = "no-cache"
    Response.Buffer = TRUE

%>
<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->

<!-- #include file = "ClassInclude.asp" -->
<%

Dim rst, rst2, SQL, arrDat
set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")

RecoverSQLConnection()
RecoverSession(false)


dim TipoReq: TipoReq = Request("T")
dim Dat: Dat = Request("D")

Dim strOut
strOut = ""


Dim arrThemes, t


if TipoReq = "ListaTematicas" then
	
	'Dat viene con IDClient;IDTheme
	
	Dim IDClient, IDTheme
	
	if InStr(Dat, ";") then
	    arrDat = split(Dat, ";")
	    IDClient = arrDat(0)
	    IDTheme = arrDat(1)
	else
	    IDClient = Dat
	end if

	strOut = ""
	
	if IDTheme<>"" then
	    arrThemes = getThemesIncludeCurrent(CInt(IDClient), CInt(IDTheme))
	else
    	arrThemes = getThemes(CInt(IDClient))
    end if
    for each t in arrThemes
		strOut = strOut & ",{""id"":""" & t.ID & """,""Name"":""" & t.Name & """,""ImageFileName"":""" & t.ImageFileName & """}"
    next
	
	
	strOut = "{""Themes"":[" & mid(strOut,2) & "]}"

elseif TipoReq = "SaveCalExp" then
    
	if IsInputQuality() then
	
	    ' Dat viene con CE_IIIII___valor
	    '   donde IIIII = IDActivity
    	
    	on error resume next
	    ID = Mid(Dat, 4, 5)
	    Value = Mid(Dat, 12)
        
        SQL = "UPDATE Activity " & _
        " SET IDCalidadExp = " & Value & _
        " WHERE ID = " & ID
	    ObjConnectionSQL.Execute SQL
        
        if Err<>0 then
            strOut = Err.Number & ""
        else
            strOut = "OK"
        end if
        
    else
    	strOut = "OK"
	end if
	
    
elseif TipoReq = "SaveCalOf" then
    
	if IsInputQuality() then
	
	    ' Dat viene con CE_IIIII___valor
	    '   donde IIIII = IDActivity
    	
    	on error resume next
	    ID = Mid(Dat, 4, 5)
	    Value = Mid(Dat, 12)
        
        SQL = "UPDATE Activity " & _
        " SET IDCalidadOf = " & Value & _
        " WHERE ID = " & ID
	    ObjConnectionSQL.Execute SQL
        
        if Err<>0 then
            strOut = Err.Number & ""
        else
            strOut = "OK"
        end if
        
    else
    	strOut = "OK"
	end if
	

elseif TipoReq = "GuardarTitulo" then
	
	hayError = FALSE
	
	'Dat viene con IDProt;Titulo
	
	arrDat = split(Dat, ";")
	IDProt = arrDat(0)
	Titulo = arrDat(1)
	
	if isNumeric(IDProt) then
		on error resume next
		IDProt = CLng(IDProt)
		if Err<>0 then
			hayError = TRUE
		end if
		on error goto 0
		
		strTitulo = "NULL"
		if Titulo<>"" then
			strTitulo = "'" & replace(Titulo,"'","''") & "'"
		end if
		
		if Session("WI_DISTR_IDDistr")<>"" then
			
			if NOT hayError then
				
				SQL = "UPDATE PROTProtocolo " & _
				" SET Nombre=" & strTitulo & " " & _
				" WHERE IDProt= " & IDProt & " AND " & _
				" IDProt IN (SELECT IDProt FROM PROTProtocolo WHERE IDDistr='" & Session("WI_DISTR_IDDistr") & "') "
				conn.Execute SQL, n
				if n>0 then
					strOut = "OK"
				else
					strOut = "ERROR"
				end if
			else
				strOut = "ERROR"
			end if
		else
			strOut = "ERROR"
		end if
	else
		strOut = "ERROR"
	end if
	
	
end if

Response.Clear
Response.Write strOut

%>
