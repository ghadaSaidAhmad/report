<%


Function XLTextClean(txt)
    if txt<>"" then
        XLTextClean = Replace(txt, chr(13), "")
    else
        XLTextClean = ""
    end if
End Function

Sub PintarTitulo(sSheet, sTitulo1, sTitulo2)
    
	sSheet.Columns("A").ColumnWidth = 15
	sSheet.Columns("B").ColumnWidth = 5
	sSheet.Columns("C").ColumnWidth = 15
	
    sSheet.Rows(3).RowHeight = 40
    sSheet.Rows("4:6").Hidden = True
    
    sSheet.Range("A3:CA500").Font.Size = 9
    
    sSheet.Range("A1:A2").Font.Bold = True
    sSheet.Range("A1:A2").Font.Size = 20
    sSheet.Range("A1:A2").Font.Color = RGB(0, 0, 0)
    
	sSheet.Cells(1, 1).Value = sTitulo1
	sSheet.Cells(2, 1).Value = sTitulo2
    
	sSheet.Cells(1, 6).Value = IDM_SOAUpdated & ": " & Right("0" & Day(Date), 2) & "/" & Right("0" & Month(Date), 2) & "/" & Year(Date) & " " & Right("0" & Hour(Time), 2) & ":" & Right("0" & Minute(Time), 2)
End Sub

Sub PintarCalendarioXL(sSheet, FromRow, StartYear, StartMonth, ViewMonths)
    dim i, iMonth, iYear, iLoop, iColMes
    

    iMonth = StartMonth
    iYear = StartYear
    iLoop = 1
    for i = StartMonth to (StartMonth + ViewMonths - 1)
        iMonth = i
        if iMonth > 12 then
            iMonth = iMonth - 12
            iYear = StartYear + 1
        end if
        
        iColMes = (iLoop*2) + 2
        
        ' Nombre del MES
	    sSheet.Cells(FromRow, iColMes).NumberFormat = "@"
	    sSheet.Cells(FromRow, iColMes).Value = locMonthName(iMonth, Idioma) & " " & iYear
	    sSheet.Range(sSheet.Cells(FromRow, iColMes), sSheet.Cells(FromRow, iColMes+1)).Merge

        ' Borders + Align
        setFormat sSheet.Range(sSheet.Cells(FromRow, iColMes), sSheet.Cells(FromRow, iColMes+1)), RGB(256, 256, 256), 2, 3, ""
        
        ' Q1 y Q2
	    sSheet.Cells(FromRow + 1, iColMes).Value = "Q1"
	    sSheet.Cells(FromRow + 1, iColMes+1).Value = "Q2"
	    sSheet.Cells(FromRow + 1, iColMes).ColumnWidth = 20
	    sSheet.Cells(FromRow + 1, iColMes+1).ColumnWidth = 20
        
        ' Borders + Align
        setFormat sSheet.Cells(FromRow + 1, iColMes), RGB(256, 256, 256), 2, 3, ""
        setFormat sSheet.Cells(FromRow + 1, iColMes+1), RGB(256, 256, 256), 2, 3, ""
        
        iLoop = iLoop + 1
    next



    
End Sub


Sub PintarCabecerasFilas(sSheet, RowStart, CurrentLast)
    dim iRow: iRow = RowStart
    dim RGBColor
    
    RGBColor = getCurrentLastBG(CurrentLast)
            
    if Request.Form("FILTER_SHOWGENERALTHEME")<>"" then
        sSheet.Cells(iRow, 3).Value = IDM_GeneralTheme
        sSheet.Cells(iRow, 3).RowHeight = 30
        iRow = iRow + 1
    end if
    
    sSheet.Cells(iRow, 3).Value = IDM_Oferta
    sSheet.Cells(iRow, 3).RowHeight = 30
    sSheet.Cells(iRow+1, 3).Value = IDM_Folleto
    sSheet.Cells(iRow+1, 3).RowHeight = 30
    sSheet.Cells(iRow+2, 3).Value = IDM_Cabecera
    sSheet.Cells(iRow+2, 3).RowHeight = 30
    sSheet.Cells(iRow+3, 3).Value = IDM_NTiendas
    sSheet.Cells(iRow+3, 3).RowHeight = 30
    sSheet.Cells(iRow+4, 3).Value = IDM_Adicional
    sSheet.Cells(iRow+4, 3).RowHeight = 30
    
    
    iRow = iRow + 5
    
    if Request.Form("FILTER_SHOWQUALITY")<>"" then
        sSheet.Cells(iRow, 3).Value = IDM_CalidadExp
        sSheet.Cells(iRow, 3).RowHeight = 30
        sSheet.Cells(iRow + 1, 3).Value = IDM_CalidadOf
        sSheet.Cells(iRow + 1, 3).RowHeight = 30
        iRow = iRow + 2
    end if

    if Request.Form("FILTER_SHOWREALDATA_NSHOPS")<>"" then
        sSheet.Cells(iRow, 3).Value = IDM_NTiendasReal
        sSheet.Cells(iRow, 3).RowHeight = 30
        iRow = iRow + 1
    end if
    
    if Request.Form("FILTER_SHOWREALDATA_PERCENTCOMPLAINT")<>"" then
        sSheet.Cells(iRow, 3).Value = IDM_PercentComplaint
        sSheet.Cells(iRow, 3).RowHeight = 30
        iRow = iRow + 1
    end if
    
    if Request.Form("FILTER_SHOWNR")<>"" then
        sSheet.Cells(iRow, 3).Value = "NR"
        sSheet.Cells(iRow, 3).RowHeight = 30
        iRow = iRow + 1
    end if

    if Request.Form("FILTER_SHOWFC")<>"" then
        sSheet.Cells(iRow, 3).Value = "FC"
        sSheet.Cells(iRow, 3).RowHeight = 30
        iRow = iRow + 1
    end if

    if Request.Form("FILTER_SHOWNRVSLY")<>"" then
        sSheet.Cells(iRow, 3).Value = "%NR vs LY"
        sSheet.Cells(iRow, 3).RowHeight = 30
        iRow = iRow + 1
    end if
    
    
    
    ' Borders + Align
    setFormat sSheet.Range(sSheet.Cells(RowStart, 3), sSheet.Cells(iRow-1, 3)), RGBColor, 2, 1, ""

End Sub


Sub PintarCabecerasAnyo(sSheet, sYear, RowStart, ReportNumRowsPerYear, CurrentLast)
    dim RGBColor    
    RGBColor = getCurrentLastBG(CurrentLast)
    
    sSheet.Range(sSheet.Cells(RowStart, 2), sSheet.Cells(RowStart + ReportNumRowsPerYear-1, 2)).Merge
    sSheet.Range(sSheet.Cells(RowStart, 2), sSheet.Cells(RowStart + ReportNumRowsPerYear-1, 2)).Value = sYear
    
    ' Borders + Align
    setFormat sSheet.Range(sSheet.Cells(RowStart, 2), sSheet.Cells(RowStart + ReportNumRowsPerYear-1, 2)), getCurrentLastBG(CurrentLast), 2, 3, ""

    
End Sub

Sub setFormat(sRange, sBGColor, sVAlign, sHAlign, NumberFormat)
    
    sRange.HorizontalAlignment = sHAlign
    sRange.VerticalAlignment = sVAlign
    sRange.WrapText = True
    sRange.Interior.Color = sBGColor
    sRange.Borders.Color = RGB(0, 0, 0)
    if NumberFormat<>"" then
        sRange.NumberFormat = NumberFormat
    end if
    
End Sub


' El color debe tener el formato #XXXXXX
Function getRGBColorFromHexa(strHexaColor)
    dim hx1, hx2, hx3, RGBColor
    
    if strHexaColor<>"" then

        if mid(strHexaColor,1,1) = "#" then
            
            hx1 = HexToDec(mid(strHexaColor, 2, 2))
            hx2 = HexToDec(mid(strHexaColor, 4, 2))
            hx3 = HexToDec(mid(strHexaColor, 6, 2))
            if isNumeric(hx1) and isnumeric(hx2) and isnumeric(hx3) then
                if hx1 <= 256 and hx2 <= 256 and hx3 <= 256 then
                    RGBColor = RGB(hx1, hx2, hx3)
                else
                    RGBColor = RGB(128, 128, 128)
                end if
            else
                RGBColor = RGB(128, 128, 128)
            end if
        else
            RGBColor = RGB(128, 128, 128)
        end if
    else
        RGBColor = RGB(128, 128, 128)
    end if
    getRGBColorFromHexa = RGBColor
End Function


Function getCurrentLastBG(CurrentLast)
    dim strHexaColor, hx1, hx2, hx3
    
    if CurrentLast = "CURRENT" then
        strHexaColor = Application("ColorCurrentYear")
    else
        strHexaColor = Application("ColorLastYear")
    end if
    
    getCurrentLastBG = getRGBColorFromHexa(strHexaColor)
    
End Function



Sub PintarTematicaGeneralXL(sSheet, StartYear, StartMonth, ViewMonths, cli, bra, RowStart, CurrentLast)
    dim rst, SQL, strWhere
    
    strWhere = " gt.IDClient = " & cli.IDClient
    
    if (StartMonth + ViewMonths) > 13 then
        'Hay meses del año siguiente
        strWhere = strWhere & " AND ((gt.WYear = " & StartYear & " AND gt.WMonth >= " & StartMonth & ")"
        strWhere = strWhere & " OR ( gt.WYear = " & StartYear + 1 & " AND gt.WMonth <= " & StartMonth + ViewMonths - 12 - 1 & ")) "
    else
        strWhere = strWhere & " AND (gt.WYear = " & StartYear & " AND gt.WMonth >= " & StartMonth & " AND gt.WMonth <= " & StartMonth + ViewMonths - 1 & ")"
    end if
    
    SQL = "SELECT gt.*, th.Name AS ThemeName, th.ImageFileName " & _
    " FROM GeneralTheme gt " & _
    " LEFT JOIN Theme th ON gt.IDTheme = th.ID " & _
    " WHERE " & strWhere
    set rst = ObjConnectionSQL.Execute(SQL)
    
    dim mesRel, X, Y, targetCol, sColor
    while NOT rst.EOF
    
        ' ***************************************************************
        ' **** CALCULA LA COLUMNA DONDE TIENE QUE PINTAR LA ACTIVIDAD
            mesRel = rst("WMonth")
            if rst("WYear") > StartYear then
                mesRel = mesRel + 12
            end if
            X = mesRel - StartMonth
            Y = X + 4
            targetCol = X + Y
            '' O lo que es lo mismo: --->>> '''targetCol = (2*mesRel) - (2*StartMonth) + 4
        ' ***************************************************************
        ' ***************************************************************
        

        ' Si es segunda quincena, suma 1
        if rst("WHalf") = 2 then
            targetCol = targetCol + 1
        end if
        
        
        sSheet.Cells(RowStart, targetCol).Value = XLTextClean(rst("ThemeName") & " " & rst("Name"))
        
        'Si hay una imagen en la temática, se incrusta
        if rst("ImageFileName")<>"" then
            
            sSheet.Cells(RowStart, targetCol).Select
            if Err<>0 then
                TerminarConError
            end if
            sSheet.Pictures.Insert(request.servervariables("APPL_PHYSICAL_PATH") & "images\Themes\" & rst("ImageFileName"))
            if Err<>0 then
                Err.Clear
            end if
        end if
        
        ' Color según Año actual o anterior
        if CurrentLast = "CURRENT" then
            sColor = getRGBColorFromHexa(Application("ColorBGGeneralThemeCurrentYear"))
        else
            sColor = getRGBColorFromHexa(Application("ColorBGGeneralThemeLastYear"))
        end if
        
        setFormat sSheet.Cells(RowStart, targetCol), sColor, 2, 3, ""
        
    
        rst.MoveNext
        
    wend
    
    rst.close
    set rst = nothing
    
End Sub


Sub PintarActivityXL(sSheet, StartYear, StartMonth, ViewMonths, cli, bra, RowStart, CurrentLast)
    dim rst, SQL, strWhere
    dim currentRow
    
    strWhere = " act.IDClient = " & cli.IDClient & " AND act.IDBrand = " & bra.IDBrand
    
    if (StartMonth + ViewMonths) > 13 then
        'Hay meses del año siguiente
        strWhere = strWhere & " AND ((act.WYear = " & StartYear & " AND act.WMonth >= " & StartMonth & ")"
        strWhere = strWhere & " OR ( act.WYear = " & StartYear + 1 & " AND act.WMonth <= " & StartMonth + ViewMonths - 12 - 1 & ")) "
    else
        strWhere = strWhere & " AND (act.WYear = " & StartYear & " AND act.WMonth >= " & StartMonth & " AND act.WMonth <= " & StartMonth + ViewMonths - 1 & ")"
    end if
    
    SQL = "SELECT act.*, rat.BGColor AS RatioColor, " & _
    " ce.Descripcion AS DCalidadExp, co.Descripcion AS DCalidadOf " & _
    " FROM Activity act " & _
    " LEFT JOIN ActivityRatio rat ON act.IDRatio = rat.ID " & _
    " LEFT JOIN CalidadExp ce ON act.IDCalidadExp = ce.ID " & _
    " LEFT JOIN CalidadOf co ON act.IDCalidadOf = co.ID " & _
    " WHERE " & strWhere
    set rst = ObjConnectionSQL.Execute(SQL)
    
    dim mesRel, X, Y, targetCol, sColor, sColorForeground
    while NOT rst.EOF
    
        ' ***************************************************************
        ' **** CALCULA LA COLUMNA DONDE TIENE QUE PINTAR LA ACTIVIDAD
            mesRel = rst("WMonth")
            if rst("WYear") > StartYear then
                mesRel = mesRel + 12
            end if
            X = mesRel - StartMonth
            Y = X + 4
            targetCol = X + Y
            '' O lo que es lo mismo: --->>> '''targetCol = (2*mesRel) - (2*StartMonth) + 4
        ' ***************************************************************
        ' ***************************************************************
        

        ' Si es segunda quincena, suma 1
        if rst("WHalf") = 2 then
            targetCol = targetCol + 1
        end if
        
        currentRow = RowStart
        if rst("Oferta")<>"" then
            sSheet.Cells(currentRow, targetCol).Value = XLTextClean(rst("Oferta"))
            ' Color según Año actual o anterior
            if CurrentLast = "CURRENT" then
                sColor = getRGBColorFromHexa(Application("ColorBGOferta"))
            else
                sColor = getRGBColorFromHexa(Application("ColorBGOfertaLY"))
            end if
            setFormat sSheet.Cells(currentRow, targetCol), sColor, 2, 3, ""

            sColorForeground = getRGBColorFromHexa(Application("ColorFGStatus" & rst("IDStatus")))
            sSheet.Cells(currentRow, targetCol).Font.Color = sColorForeground
        end if
        currentRow = currentRow + 1

        if rst("Folleto")<>"" then
            sSheet.Cells(currentRow, targetCol).Value = XLTextClean(rst("Folleto"))
            ' Color según Año actual o anterior
            if CurrentLast = "CURRENT" then
                sColor = getRGBColorFromHexa(Application("ColorBGFolleto"))
            else
                sColor = getRGBColorFromHexa(Application("ColorBGFolletoLY"))
            end if

            setFormat sSheet.Cells(currentRow, targetCol), sColor, 2, 3, ""

            sColorForeground = getRGBColorFromHexa(Application("ColorFGStatus" & rst("IDStatus")))
            sSheet.Cells(currentRow, targetCol).Font.Color = sColorForeground
        end if
        currentRow = currentRow + 1

        if rst("Cabecera")<>"" then
            sSheet.Cells(currentRow, targetCol).Value = XLTextClean(rst("Cabecera"))
            ' Color según Año actual o anterior
            sColor = getRGBColorFromHexa(rst("RatioColor"))

            setFormat sSheet.Cells(currentRow, targetCol), sColor, 2, 3, ""

            sColorForeground = getRGBColorFromHexa(Application("ColorFGStatus" & rst("IDStatus")))
            sSheet.Cells(currentRow, targetCol).Font.Color = sColorForeground
        end if
        currentRow = currentRow + 1
        
        if rst("NShops")<>"" then
            sSheet.Cells(currentRow, targetCol).Value = rst("NShops")
            ' Color según Año actual o anterior
            if CurrentLast = "CURRENT" then
                sColor = getRGBColorFromHexa(Application("ColorBGNShops"))
            else
                sColor = getRGBColorFromHexa(Application("ColorBGNShopsLY"))
            end if
            setFormat sSheet.Cells(currentRow, targetCol), sColor, 2, 3, ""
        end if
        currentRow = currentRow + 1
    
        if rst("Adicional")<>"" then
            sSheet.Cells(currentRow, targetCol).Value = XLTextClean(rst("Adicional"))
            ' Color según Año actual o anterior
            if CurrentLast = "CURRENT" then
                sColor = getRGBColorFromHexa(Application("ColorBGAdicional"))
            else
                sColor = getRGBColorFromHexa(Application("ColorBGAdicionalLY"))
            end if
            setFormat sSheet.Cells(currentRow, targetCol), sColor, 2, 3, ""
        end if
        currentRow = currentRow + 1


        if Request.Form("FILTER_SHOWQUALITY")<>"" then
            sSheet.Cells(currentRow, targetCol).Value = rst("DCalidadExp")
            sColor = getRGBColorFromHexa("#FFFFFF")
            setFormat sSheet.Cells(currentRow, targetCol), sColor, 2, 3, ""
            currentRow = currentRow + 1

            sSheet.Cells(currentRow, targetCol).Value = rst("DCalidadOf")
            sColor = getRGBColorFromHexa("#FFFFFF")
            setFormat sSheet.Cells(currentRow, targetCol), sColor, 2, 3, ""
            currentRow = currentRow + 1
        end if

        rst.MoveNext
        
    wend
    
    rst.close
    set rst = nothing

End Sub


Sub PintarSOARealDataXL(sSheet, StartYear, StartMonth, ViewMonths, cli, bra, RowStart, CurrentLast)
    dim rst, SQL, strWhere
    dim currentRow
    
    strWhere = " rd.IDClient = " & cli.IDClient & " AND rd.IDBrand = " & bra.IDBrand
    
    if (StartMonth + ViewMonths) > 13 then
        'Hay meses del año siguiente
        strWhere = strWhere & " AND ((rd.WYear = " & StartYear & " AND rd.WMonth >= " & StartMonth & ")"
        strWhere = strWhere & " OR ( rd.WYear = " & StartYear + 1 & " AND rd.WMonth <= " & StartMonth + ViewMonths - 12 - 1 & ")) "
    else
        strWhere = strWhere & " AND (rd.WYear = " & StartYear & " AND rd.WMonth >= " & StartMonth & " AND rd.WMonth <= " & StartMonth + ViewMonths - 1 & ")"
    end if
    
    SQL = "SELECT rd.* " & _
    " FROM RealData rd " & _
    " WHERE " & strWhere
    set rst = ObjConnectionSQL.Execute(SQL)
    
    dim mesRel, X, Y, targetCol, sColor, sColorForeground
    while NOT rst.EOF
    
        ' ***************************************************************
        ' **** CALCULA LA COLUMNA DONDE TIENE QUE PINTAR LA ACTIVIDAD
            mesRel = rst("WMonth")
            if rst("WYear") > StartYear then
                mesRel = mesRel + 12
            end if
            X = mesRel - StartMonth
            Y = X + 4
            targetCol = X + Y
            '' O lo que es lo mismo: --->>> '''targetCol = (2*mesRel) - (2*StartMonth) + 4
        ' ***************************************************************
        ' ***************************************************************
        

        ' Si es segunda quincena, suma 1
        if rst("WHalf") = 2 then
            targetCol = targetCol + 1
        end if
        
        currentRow = RowStart
        if Request.Form("FILTER_SHOWREALDATA_NSHOPS")<>"" then
            sSheet.Cells(currentRow, targetCol).Value = rst("NShops")
            sSheet.Cells(currentRow, targetCol).NumberFormat = "0"
            ' Color según Año actual o anterior
            sColor = getRGBColorFromHexa("#FFFFFF")
            setFormat sSheet.Cells(currentRow, targetCol), sColor, 2, 3, "0"
            currentRow = currentRow + 1
        end if

        if Request.Form("FILTER_SHOWREALDATA_PERCENTCOMPLAINT")<>"" then
            sSheet.Cells(currentRow, targetCol).Value = rst("PercentComplaint")
            sSheet.Cells(currentRow, targetCol).NumberFormat = "0"
            ' Color según Año actual o anterior
            sColor = getRGBColorFromHexa("#FFFFFF")
            setFormat sSheet.Cells(currentRow, targetCol), sColor, 2, 3, "0"
            currentRow = currentRow + 1
        end if

        rst.MoveNext
        
    wend
    
    rst.close
    set rst = nothing
    
End Sub


Sub PintarNRXL(sSheet, StartYear, StartMonth, ViewMonths, cli, bra, RowStart, CurrentLast)
    dim rst, SQL
    
    dim i, mesRel, X, Y, targetCol, sColor, iMonth, iYear
    for i = StartMonth to (StartMonth+ViewMonths-1)
    
        iMonth = i
        iYear = StartYear
        if iMonth>12 then 
            iMonth = iMonth - 12
            iYear = StartYear + 1
        end if

        SQL = "SELECT NR " & _
        " FROM NetRevenue " & _
        " WHERE IDClient = " & cli.IDClient & " AND IDBrand = " & bra.IDBrand & " AND WYear = " & iYear & " AND WMonth = " & iMonth
        set rst = ObjConnectionSQL.Execute(SQL)

        ' ***************************************************************
        ' **** CALCULA LA COLUMNA DONDE TIENE QUE PINTAR LA ACTIVIDAD
            mesRel = iMonth
            if iYear > StartYear then
                mesRel = mesRel + 12
            end if
            X = mesRel - StartMonth
            Y = X + 4
            targetCol = X + Y
            '' O lo que es lo mismo: --->>> '''targetCol = (2*mesRel) - (2*StartMonth) + 4
        ' ***************************************************************
        ' ***************************************************************
        
        sSheet.Range(sSheet.Cells(RowStart, targetCol), sSheet.Cells(RowStart, targetCol+1)).Merge
        
        if rst.EOF then
            sSheet.Cells(RowStart, targetCol).Value = 0
        else
            sSheet.Cells(RowStart, targetCol).Value = rst("NR")
        end if
        '''''sSheet.Cells(RowStart, targetCol).NumberFormat = "#,###0.00"
        
        
        ' Color según Año actual o anterior
        if CurrentLast = "CURRENT" then
            sColor = getRGBColorFromHexa(Application("ColorNR_CY"))
        else
            sColor = getRGBColorFromHexa(Application("ColorNR_LY"))
        end if
        
        setFormat sSheet.Cells(RowStart, targetCol), sColor, 2, 3, "#,###0.00"
        
        rst.close
        set rst = nothing
    next
    
    
End Sub


Sub PintarFCXL(sSheet, StartYear, StartMonth, ViewMonths, cli, bra, RowStart, CurrentLast)
    dim rst, SQL
    
    dim i, mesRel, X, Y, targetCol, sColor, iMonth, iYear
    for i = StartMonth to (StartMonth+ViewMonths-1)
    
        iMonth = i
        iYear = StartYear
        if iMonth>12 then 
            iMonth = iMonth - 12
            iYear = StartYear + 1
        end if

        SQL = "SELECT FC " & _
        " FROM Forecast " & _
        " WHERE IDClient = " & cli.IDClient & " AND IDBrand = " & bra.IDBrand & " AND WYear = " & iYear & " AND WMonth = " & iMonth
        set rst = ObjConnectionSQL.Execute(SQL)

        ' ***************************************************************
        ' **** CALCULA LA COLUMNA DONDE TIENE QUE PINTAR LA ACTIVIDAD
            mesRel = iMonth
            if iYear > StartYear then
                mesRel = mesRel + 12
            end if
            X = mesRel - StartMonth
            Y = X + 4
            targetCol = X + Y
            '' O lo que es lo mismo: --->>> '''targetCol = (2*mesRel) - (2*StartMonth) + 4
        ' ***************************************************************
        ' ***************************************************************
        
        sSheet.Range(sSheet.Cells(RowStart, targetCol), sSheet.Cells(RowStart, targetCol+1)).Merge
        
        if rst.EOF then
            sSheet.Cells(RowStart, targetCol).Value = 0
        else
            sSheet.Cells(RowStart, targetCol).Value = rst("FC")
        end if
        sSheet.Cells(RowStart, targetCol).NumberFormat = "#,###0.00"
        
        
        ' Color según Año actual o anterior
        if CurrentLast = "CURRENT" then
            sColor = getRGBColorFromHexa(Application("ColorFC_CY"))
        else
            sColor = getRGBColorFromHexa(Application("ColorFC_LY"))
        end if
        
        setFormat sSheet.Cells(RowStart, targetCol), sColor, 2, 3, "#,###0.00"
        
        rst.close
        set rst = nothing
    next
    
    
End Sub

Sub PintarNRVSLYXL(sSheet, StartYear, StartMonth, ViewMonths, cli, bra, RowStart, CurrentLast)
    dim rst, SQL, NR, NRLY, Percent
    dim i, mesRel, X, Y, targetCol, sColor, iMonth, iYear
    
    for i = StartMonth to (StartMonth+ViewMonths-1)
    
        iMonth = i
        iYear = StartYear
        if iMonth>12 then 
            iMonth = iMonth - 12
            iYear = StartYear + 1
        end if

        SQL = "SELECT NR " & _
        " FROM NetRevenue " & _
        " WHERE IDClient = " & cli.IDClient & " AND IDBrand = " & bra.IDBrand & " AND WYear = " & iYear & " AND WMonth = " & iMonth
        set rst = ObjConnectionSQL.Execute(SQL)
        if not rst.EOF then
            NR = rst("NR")
        else
            NR = 0
        end if
        rst.close
        set rst = nothing

        SQL = "SELECT NR " & _
        " FROM NetRevenue " & _
        " WHERE IDClient = " & cli.IDClient & " AND IDBrand = " & bra.IDBrand & " AND WYear = " & iYear-1 & " AND WMonth = " & iMonth
        set rst = ObjConnectionSQL.Execute(SQL)
        if not rst.EOF then
            NRLY = rst("NR")
        else
            NRLY = 0
        end if
        rst.close
        set rst = nothing
        
        Percent = 0
        if NRLY>0 then
            Percent = NR * 100 / NRLY
        end if
        if Percent>0 AND Percent<100 then
            Percent = - 100 + Percent
        elseif Percent>100 then
            Percent = Percent - 100
        end if
        
        
        ' ***************************************************************
        ' **** CALCULA LA COLUMNA DONDE TIENE QUE PINTAR LA ACTIVIDAD
            mesRel = iMonth
            if iYear > StartYear then
                mesRel = mesRel + 12
            end if
            X = mesRel - StartMonth
            Y = X + 4
            targetCol = X + Y
            '' O lo que es lo mismo: --->>> '''targetCol = (2*mesRel) - (2*StartMonth) + 4
        ' ***************************************************************
        ' ***************************************************************
        
        sSheet.Range(sSheet.Cells(RowStart, targetCol), sSheet.Cells(RowStart, targetCol+1)).Merge
        
        sSheet.Cells(RowStart, targetCol).NumberFormat = "0.00%"
        sSheet.Cells(RowStart, targetCol).Value = Percent
        
        
        ' Color según Año actual o anterior
        if CurrentLast = "CURRENT" then
            sColor = getRGBColorFromHexa(Application("ColorNRvsLY_CY"))
        else
            sColor = getRGBColorFromHexa(Application("ColorNRvsLY_LY"))
        end if
        
        setFormat sSheet.Cells(RowStart, targetCol), sColor, 2, 3, "0.00%"
        
    next
    
    
End Sub


Sub PintarReportClientBrandXL(sSheet, StartYear, StartMonth, ViewMonths, cli, bra, ReportNumRowsPerYear, RowStart, ImageFileName, TitleName)
    dim showLY, currentRowStart, CurrentLast, currentRow
    dim RGBColor
    
    showLY = (Request.Form("FILTER_LASTYEAR")<>"")
    currentRowStart = RowStart
    
    dim iYearLoop, lastYearLoop
    if showLY then
        lastYearLoop = StartYear - 1
    else
        lastYearLoop = StartYear
    end if
    
    
    ' Imagen o Nombre
    sSheet.Cells(currentRowStart, 1).Select
    if ImageFileName<>"" then
        dim fso: set fso = CreateObject("Scripting.FileSystemObject")
        if fso.FileExists(ImageFileName) then

            if Err<>0 then
                TerminarConError
            end if
            sSheet.Pictures.Insert(ImageFileName)
            if Err<>0 then
                sSheet.Cells(currentRowStart, 1).Value = TitleName
                Err.Clear
            end if
        else
            sSheet.Cells(currentRowStart, 1).Value = TitleName
        end if
    else
        sSheet.Cells(currentRowStart, 1).Value = TitleName
    end if
    
    CurrentLast = "CURRENT"
    for iYearLoop = StartYear To lastYearLoop Step -1
        
        
        ' Escribe el AÑO en la columna "B"
        PintarCabecerasAnyo sSheet, iYearLoop, currentRowStart, ReportNumRowsPerYear, CurrentLast
        
        ' Escribe las cabeceras de fila en la columna "C"
        PintarCabecerasFilas sSheet, currentRowStart, CurrentLast
        
        currentRow = currentRowStart
        
        if Request.Form("FILTER_SHOWGENERALTHEME")<>"" then
            PintarTematicaGeneralXL sSheet, iYearLoop, StartMonth, ViewMonths, cli, bra, currentRow, CurrentLast
            currentRow = currentRow + 1
        end if
        
        PintarActivityXL sSheet, iYearLoop, StartMonth, ViewMonths, cli, bra, currentRow, CurrentLast
        currentRow = currentRow + 5
        if Request.Form("FILTER_SHOWQUALITY")<>"" then
            currentRow = currentRow + 2
        end if
        
        if Request.Form("FILTER_SHOWREALDATA_NSHOPS")<>"" OR Request.Form("FILTER_SHOWREALDATA_PERCENTCOMPLAINT")<>"" then
            PintarSOARealDataXL sSheet, iYearLoop, StartMonth, ViewMonths, cli, bra, currentRow, CurrentLast
            if Request.Form("FILTER_SHOWREALDATA_NSHOPS")<>"" then
                currentRow = currentRow + 1
            end if
            if Request.Form("FILTER_SHOWREALDATA_PERCENTCOMPLAINT")<>"" then
                currentRow = currentRow + 1
            end if
        end if
        
        
        if Request.Form("FILTER_SHOWNR")<>"" then
            PintarNRXL sSheet, iYearLoop, StartMonth, ViewMonths, cli, bra, currentRow, CurrentLast
            
            currentRow = currentRow + 1
        end if
        
        if Request.Form("FILTER_SHOWFC")<>"" then
            PintarFCXL sSheet, iYearLoop, StartMonth, ViewMonths, cli, bra, currentRow, CurrentLast
            
            currentRow = currentRow + 1
        end if

        if Request.Form("FILTER_SHOWNRVSLY")<>"" then
            PintarNRVSLYXL sSheet, iYearLoop, StartMonth, ViewMonths, cli, bra, currentRow, CurrentLast
            
            currentRow = currentRow + 1
        end if

        ' Incrementa el currentRowStart por si pinta el año anterior
        ' Si no, no pasa nada
        currentRowStart = currentRowStart + ReportNumRowsPerYear
        CurrentLast = "LAST"
    Next
    
    
End Sub

%>