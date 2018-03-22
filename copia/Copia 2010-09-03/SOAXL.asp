<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
    Server.ScriptTimeout = 300
	Response.Buffer=true
    
%>
<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/Idioma.asp" -->

<!-- #include file = "RenderFunctionsXL.asp" -->
<!-- #include file = "ClassInclude.asp" -->
<%
Sub TerminarConError
    
    objExcel.Quit
    set objExcel = nothing
    
    Response.Write "<table width=600 align=center style=""border:2 solid red;""><tr><td align=center>ERROR<br>" & Err.Description & "</td></tr></table>"
    Response.End
    
End Sub


    dim rst: set rst = CreateObject("ADODB.RecordSet"): rst.CursorLocation = adUseClient
    dim rst2: set rst2 = CreateObject("ADODB.RecordSet"): rst2.CursorLocation = adUseClient
    dim nextRowStart, ReportRowStart, itemRowStart, reportNCols, clientSideFileName
    dim bra, b, cli, c
    dim sName, sTitle
    dim ReportNumRowsPerYear
    dim ViewMonths, StartYear, StartMonth

    RecoverSQLConnection()
    RecoverSession(true)
    
    if Request.Form("FILTER_VIEWMONTHS")<>"" then
        ViewMonths = CInt(Request.Form("FILTER_VIEWMONTHS"))
    else
        ViewMonths = 4
    end if
    if Request.Form("FILTER_YEAR")<>"" then
        StartYear = CInt(Request.Form("FILTER_YEAR"))
    else
        StartYear = Year(Date)
    end if
    if Request.Form("FILTER_STARTMONTH")<>"" then
        StartMonth = CInt(Request.Form("FILTER_STARTMONTH"))
    else
        StartMonth = Month(Date)
    end if


    ReportNumRowsPerYear = getReportRows()

    

    ' CONTROL DE ERRORES ************************************************************
    ' Esto no hay que quitarlo. De vez en cuando el código mirará si ha habido algún 
    ' error y parará la ejecución para cerrar el objeto Excel (objExcel) ya que,
    ' si no se cierra, queda en memoria en el servidor
    ' Cuando hay una instrucción de la que queremos controlar el error,
    ' justo antes se comprueba si había algún error para así terminar con la ejecución. 
    ' Luego se ejecuta la instrucción que queremos controlar, y si hay error, 
    ' se hace lo que toque y luego Err.Clear
    
     on error resume next
     
    ' *******************************************************************************
    
    
    
    ' OBJETOS NECESARIOS PARA GENERAR EL FICHERO EXCEL ******************************
    dim objExcel
    
    if isObject(Session("ExcelApp")) then
        set objExcel = Session("ExcelApp")
        if objExcel = "" then
            set objExcel = CreateObject("Excel.Application")
            set Session("ExcelApp") = objExcel
        end if
    else
        set objExcel = CreateObject("Excel.Application")
        set Session("ExcelApp") = objExcel
    end if
    objExcel.DisplayAlerts = False ' MUY IMPORTANTE PARA QUE NO SE QUEDE EL OBJETO EN MEMORIA
    dim fso: set fso = CreateObject("Scripting.FileSystemObject")
    dim objLibro: set objLibro = objExcel.Workbooks.Add
    dim sSheet: set sSheet = objLibro.Sheets(1)
    ' *******************************************************************************
    
    
    if Err<>0 then
        TerminarConError
    end if
    
    objLibro.Sheets(1).Name = "SOA"
    objExcel.ActiveWindow.DisplayGridlines = false

    ' Título del report
    PintarTitulo sSheet, IDM_MAINTITLE1, IDM_MAINTITLE2
    
    ' Pinta el calendario a partir de la celda D4
    PintarCalendarioXL sSheet, 7, StartYear, StartMonth, ViewMonths
    
    'Calcula el número de columnas del report
    reportNCols = 3
    reportNCols = 3 + (ViewMonths*2)

    ' Fila donde empieza el report
    nextRowStart = 9
    ReportRowStart = nextRowStart
    
    if Err<>0 then
        TerminarConError
    end if
    if Request.Form("FILTER_REPORTTYPE") = "0" then
        set cli = getClient(Request.Form("FILTER_CLIENT"))
        
        clientSideFileName = "SOA_" & cli.Name & ".xls"

        ' INSERTAR UNA IMAGEN
        if cli.ImageFileNameH<>"" then
            sSheet.Cells(3, 1).Select
            
            if Err<>0 then
                TerminarConError
            end if
            sSheet.Pictures.Insert(request.servervariables("APPL_PHYSICAL_PATH") & "images\Clients\" & cli.ImageFileNameH )
            if Err<>0 then
                sSheet.Cells(3, 1).Value = cli.Name
                Err.Clear
            end if
        else
            sSheet.Cells(3, 1).Value = cli.Name
        end if
        
        for each b in split(Request.Form("FILTER_MULTIBRAND"), ",")
            set bra = getBrand(CInt(b))
            
            itemRowStart = nextRowStart
            
            PintarReportClientBrandXL sSheet, StartYear, StartMonth, ViewMonths, cli, bra, ReportNumRowsPerYear, nextRowStart, request.servervariables("APPL_PHYSICAL_PATH") & "images\Brands\" & bra.ImageFileNameV, bra.Name
            
            if Request.Form("FILTER_LASTYEAR")<>"" then
                nextRowStart = nextRowStart + (ReportNumRowsPerYear*2)
            else
                nextRowStart = nextRowStart + ReportNumRowsPerYear
            end if

            sSheet.Range(sSheet.Cells(itemRowStart,3), sSheet.Cells(nextRowStart-1,reportNCols)).Borders.Color = RGB(0, 0, 0)
            
            'Deja una fila en blanco
            nextRowStart = nextRowStart + 2
            sSheet.Range(sSheet.Cells(nextRowStart-1, 1), sSheet.Cells(nextRowStart-1, reportNCols)).Borders(8).LineStyle = -4119

            if Err<>0 then
                TerminarConError
            end if
        next
        
    elseif Request.Form("FILTER_REPORTTYPE") = "1" then 
        set bra = getBrand(Request.Form("FILTER_BRAND"))

        clientSideFileName = "SOA_" & bra.Name & ".xls"

        ' INSERTAR UNA IMAGEN
        if bra.ImageFileNameH<>"" then
            sSheet.Cells(3, 1).Select
            
            if Err<>0 then
                TerminarConError
            end if
            sSheet.Pictures.Insert(request.servervariables("APPL_PHYSICAL_PATH") & "images\Brands\" & bra.ImageFileNameH )
            if Err<>0 then
                sSheet.Cells(3, 1).Value = bra.Name
                Err.Clear
            end if
        else
            sSheet.Cells(3, 1).Value = bra.Name
        end if
        
        for each c in split(Request.Form("FILTER_MULTICLIENT"), ",")
        
            set cli = getClient(CInt(c))

            itemRowStart = nextRowStart

            PintarReportClientBrandXL sSheet, StartYear, StartMonth, ViewMonths, cli, bra, ReportNumRowsPerYear, nextRowStart, request.servervariables("APPL_PHYSICAL_PATH") & "images\Clients\" & cli.ImageFileNameV, cli.Name
            
            if Request.Form("FILTER_LASTYEAR")<>"" then
                nextRowStart = nextRowStart + (ReportNumRowsPerYear*2)
            else
                nextRowStart = nextRowStart + ReportNumRowsPerYear
            end if

            sSheet.Range(sSheet.Cells(itemRowStart,3), sSheet.Cells(nextRowStart-1,reportNCols)).Borders.Color = RGB(0, 0, 0)

            'Deja una fila en blanco
            nextRowStart = nextRowStart + 2
            sSheet.Range(sSheet.Cells(nextRowStart-1, 1), sSheet.Cells(nextRowStart-1, reportNCols)).Borders(8).LineStyle = -4119
            
            if Err<>0 then
                TerminarConError
            end if
        next

    end if
    
    
    
    ' INMOBILIZA LOS PANELES SUPERIOR E IZQUIERDO
    sSheet.Select
    sSheet.Cells(ReportRowStart, 4).Select
    objExcel.ActiveWindow.FreezePanes = true
	
	
    ' GUARDA EL FICHERO *************************************************************
    dim fileName: fileName =  "SOA_" & session("IDUser") & ".xls"
    dim fullFileName: fullFileName = request.servervariables("APPL_PHYSICAL_PATH") & "XL\" & fileName
    if fso.FileExists(fullFileName) then
	    fso.DeleteFile fullFileName
    end if
    objLibro.SaveAs fullFileName
    if Err<>0 then
        TerminarConError
    end if
    objLibro.Close
    
    ' Esto a veces no funciona. Por este motivo hemos colocado el objeto en una session
    ' que intentará hacer el Quit cuando se cierre (global.asa Session_OnEnd)
    objExcel.Quit  
    
    set sSheet = nothing
    set objLibro = nothing
    set objExcel = nothing
    ' *******************************************************************************
    
    
    if Err<>0 then
        TerminarConError
    end if

    ' ENVÍA EL FICHERO QUE SE HA GENERADO
    ' *******************************************************************************
    Response.AddHeader "Content-Type", "application/vnd.ms-excel"
    Response.AddHeader "Content-disposition", "attachment; filename=" & clientSideFileName

    Const clChunkSize = 1048 ' 100KB
    Dim oStream : Set oStream = Server.CreateObject("ADODB.Stream")
    oStream.Type = 1 ' Binary
    oStream.open
    oStream.LoadFromFile(fullFileName)
    Dim i
    For i = 0 To oStream.Size \ clChunkSize
        Response.BinaryWrite oStream.Read(clChunkSize)
    Next
    oStream.close
    ' *******************************************************************************
    
    
    ' BORRA EL FICHERO **************************************************************
    if fso.FileExists(fullFileName) then
	    fso.DeleteFile fullFileName
    end if
    
%>
