<script runat=server language="vbscript">
    
    Function PrepararColumnasCalendario(ViewMonths)
        dim strOut
        strOut = ""
        dim i
        
        ' Primero pinta 10 columnas (año + título + 8 quincenas) en blanco
        '    (para evitar problemas con los colspan)
        strOut = strOut & "<tr>"
            strOut = strOut & "<td width='60px'></td>"
            strOut = strOut & "<td></td>"
            for i = 1 to ViewMonths*2
                strOut = strOut & "<td width='" & Application("ReportHalfWidth") & "px'></td>"
            next
        strOut = strOut & "</tr>"
        
        PrepararColumnasCalendario = strOut
    End Function
    
    
    ' Show CALENDAR
    Function PintarCalendario (StartYear, StartMonth, ViewMonths, Title)
        dim strOut
        strOut = ""
        
        dim i, iMonth, sYear, sQuincena, sClass

        
        ' Cabecera título
        strOut = strOut & "<tr>"
            strOut = strOut & "<td colspan=3 class=""gridlefttitle"">" & Title & "</td>"
            ' Nombre de los meses
            for i = StartMonth to (StartMonth+ViewMonths-1)
                iMonth = i
                sYear = ""
                sYear = "<font class=font10><br />" & StartYear & "</font>"
                if iMonth>12 then 
                    iMonth = iMonth - 12
                    sYear = "<font class=font10><br />" & StartYear + 1 & "</font>"
                end if
                
                strOut = strOut & "<td colspan=2 class=gridmonthtitle>" & locMonthName(iMonth, Idioma) & sYear & "</td>"
            next
        strOut = strOut & "</tr>"
        
        ' Cabecera quincenas
        strOut = strOut & "<tr>"
            strOut = strOut & "<td class=""gridlefthalfrow""><font class=font9>&nbsp;</font></td>"
            strOut = strOut & "<td></td>"
            strOut = strOut & "<td></td>"
            ' Nombre de los meses
            for i = 1 to ViewMonths*2
                if i mod 2 = 0 then
                    sQuincena = IDM_2aQuincena
                    sClass = "gridhalf2title"
                else
                    sQuincena = IDM_1aQuincena
                    sClass = "gridhalf1title"
                end if
                strOut = strOut & "<td class=" & sClass & ">" & sQuincena & "</td>"
            next
        strOut = strOut & "</tr>"
        
        
        PintarCalendario = strOut
    End Function
    
    
    Function PintarCalendarioNavegacionActividad (IDClient, IDBrand, StartYear, StartMonth, ViewMonths, currYear, currMonth, currHalf)
        dim strOut, strOutMeses, strOutQuincenas
        strOut = ""
        strOutMeses = "<tr>"
        strOutQuincenas = "<tr>"
        
        dim i, iMonth, sYear, iYear, sQuincena, sClass, q1Style, q2Style
        dim act1, act2
        
        strOut = strOut & "<table width=""100%"" cellpadding=0 cellspacing=0 style=""border-right:1px solid gray;border-bottom:1px solid gray;"">"
            ' Nombre de los meses
            for i = StartMonth to (StartMonth+ViewMonths-1)
                iMonth = i
                sYear = ""
                sYear = "<font class=font10><br />" & StartYear & "</font>"
                iYear = StartYear
                if iMonth>12 then 
                    iMonth = iMonth - 12
                    sYear = "<font class=font10><br />" & StartYear + 1 & "</font>"
                    iYear = StartYear + 1
                end if
                
                q1Style = ""
                q2Style = ""

                set act1 = getActivityFromDate(IDClient, IDBrand, iYear, iMonth, 1)
                set act2 = getActivityFromDate(IDClient, IDBrand, iYear, iMonth, 2)
                if act1.ID <> -1 then
                    q1Style = q1Style & "border:2px solid silver;background-color:gray;color:white;font-weight:bold;"
                else
                    q1Style = q1Style & "border:2px solid silver;background-color:white;color:black;font-weight:normal;"
                end if
                if act2.ID <> -1 then
                    q2Style = q2Style & "border:2px solid silver;background-color:gray;color:white;font-weight:bold;"
                else
                    q2Style = q2Style & "border:2px solid silver;background-color:white;color:black;font-weight:normal;"
                end if
                
                if CInt(currYear) = iYear AND CInt(currMonth) = iMonth AND CInt(currHalf) = 1 then
                    q1Style = q1Style & "border:2px solid red;"
                elseif CInt(currYear) = iYear AND CInt(currMonth) = iMonth AND CInt(currHalf) = 2 then
                    q2Style = q2Style & "border:2px solid red;"
                end if
                
                strOutMeses = strOutMeses & "<td colspan=2 class=gridmonthtitlesmall>" & Left(locMonthName(iMonth, Idioma), 3) & sYear & "</td>"
                
                strOutQuincenas = strOutQuincenas & "<td class=gridhalf1title style=""border-top:1px solid gray;cursor:pointer;" & q1Style & """ onClick=""navigateTo(" & iYear & ", " & iMonth & ", 1);"">1</td>"
                strOutQuincenas = strOutQuincenas & "<td class=gridhalf1title style=""border-top:1px solid gray;cursor:pointer;" & q2Style & """ onClick=""navigateTo(" & iYear & ", " & iMonth & ", 2);"">2</td>"
                
            next
            
        strOutMeses = strOutMeses & "</tr>"
        strOutQuincenas = strOutQuincenas & "</tr>"
        
        strOut = strOut & strOutMeses & strOutQuincenas
        
        strOut = strOut & "</table>"
        
        PintarCalendarioNavegacionActividad = strOut
    End Function
    

    Function PintarCalendarioNavegacionGeneralTheme (IDClient, StartYear, StartMonth, ViewMonths, currYear, currMonth, currHalf)
        dim strOut, strOutMeses, strOutQuincenas
        strOut = ""
        strOutMeses = "<tr>"
        strOutQuincenas = "<tr>"
        
        dim i, iMonth, sYear, iYear, sQuincena, sClass, q1Style, q2Style
        dim gthm1, gthm2
        
        strOut = strOut & "<table width=""100%"" cellpadding=0 cellspacing=0 style=""border-right:1px solid gray;border-bottom:1px solid gray;"">"
            ' Nombre de los meses
            for i = StartMonth to (StartMonth+ViewMonths-1)
                iMonth = i
                sYear = ""
                sYear = "<font class=font10><br />" & StartYear & "</font>"
                iYear = StartYear
                if iMonth>12 then 
                    iMonth = iMonth - 12
                    sYear = "<font class=font10><br />" & StartYear + 1 & "</font>"
                    iYear = StartYear + 1
                end if
                
                q1Style = ""
                q2Style = ""

                set gthm1 = getGeneralThemeFromDate(IDClient, iYear, iMonth, 1)
                set gthm2 = getGeneralThemeFromDate(IDClient, iYear, iMonth, 2)
                if gthm1.ID <> -1 then
                    q1Style = q1Style & "border:2px solid silver;background-color:gray;color:white;font-weight:bold;"
                else
                    q1Style = q1Style & "border:2px solid silver;background-color:white;color:black;font-weight:normal;"
                end if
                if gthm2.ID <> -1 then
                    q2Style = q2Style & "border:2px solid silver;background-color:gray;color:white;font-weight:bold;"
                else
                    q2Style = q2Style & "border:2px solid silver;background-color:white;color:black;font-weight:normal;"
                end if
                
                if CInt(currYear) = iYear AND CInt(currMonth) = iMonth AND CInt(currHalf) = 1 then
                    q1Style = q1Style & "border:2px solid red;"
                elseif CInt(currYear) = iYear AND CInt(currMonth) = iMonth AND CInt(currHalf) = 2 then
                    q2Style = q2Style & "border:2px solid red;"
                end if
                
                strOutMeses = strOutMeses & "<td colspan=2 class=gridmonthtitlesmall>" & Left(locMonthName(iMonth, Idioma), 3) & sYear & "</td>"
                
                strOutQuincenas = strOutQuincenas & "<td class=gridhalf1title style=""border-top:1px solid gray;cursor:pointer;" & q1Style & """ onClick=""navigateTo(" & iYear & ", " & iMonth & ", 1);"">1</td>"
                strOutQuincenas = strOutQuincenas & "<td class=gridhalf1title style=""border-top:1px solid gray;cursor:pointer;" & q2Style & """ onClick=""navigateTo(" & iYear & ", " & iMonth & ", 2);"">2</td>"
                
            next
            
        strOutMeses = strOutMeses & "</tr>"
        strOutQuincenas = strOutQuincenas & "</tr>"
        
        strOut = strOut & strOutMeses & strOutQuincenas
        
        strOut = strOut & "</table>"
        
        PintarCalendarioNavegacionGeneralTheme = strOut
    End Function

    
    Function PintarGrupo0(rowspan, Title, Anchor)
        dim strOut
        strOut = ""
        dim sTitle
        
        strOut = strOut & "<tr><td valign=top class=gridgroup0title rowspan=" & rowspan & "><a name=""" & Anchor & """>" & Title & "</td></tr>"
        
        PintarGrupo0 = strOut
    End Function
    

    Function PintarGrupo05(rowspan, Title, RowTitleBgColor)
        dim strOut
        strOut = ""
        dim sTitle
        
        strOut = strOut & "<tr><td class=gridtypetitle rowspan=" & rowspan & " bgcolor=""" & RowTitleBgColor & """>" & Title & "</td></tr>"
        
        PintarGrupo05 = strOut
    End Function

    Function PintarPrintPageBreak()
        dim strOut
        strOut = ""
        
        strOut = strOut & "<tr style=""page-break-after:always;""><td></td></tr>"
        
        PintarPrintPageBreak = strOut
    End Function

    

    
    
    ' Pinta las actividades de un tipo en el calendario
    Function PintarActivity(StartYear, StartMonth, ViewMonths, IDClient, IDBrand, CurrentLast)
        dim arrActivity
        dim strOut, strOut1, strOut2, strOut3, strOut4, strOut6, strOutKPI, strCalExp, strCalOf, strOutRD_NShops, strOutTOTALNShops
        dim sClassOferta1, sClassOferta2, FGColorOferta1, FGColorOferta2
        dim sClassFolleto1, sClassFolleto2, FGColorFolleto1, FGColorFolleto2
        dim sClassCabecera1, sClassCabecera2, BGColorCabecera1, FGColorCabecera1, BGColorCabecera2, FGColorCabecera2
        dim sClassCentros1, sClassCentros2, FGColorCentros1, FGColorCentros2
        dim sClassAdicional1, sClassAdicional2, FGColorAdicional1, FGColorAdicional2
        dim sClassKpiQuality1, sClassKpiQuality2, FGColorKpiQuality1, FGColorKpiQuality2
        dim sClass, sOnClick1, sOnClick2
        dim QExposicion, q, QOferta
        dim rst, SQL
        dim bra, iSubcat
        dim strOutSubcatNShops(9)
        dim strOutSubcatRD_NShops(9), strOutSubcatTOTALNShops(9)
        dim i, iMonth, iYear
        
        ' Carga los datos de la marca para usar las subcategorías
        set bra = getBrand(IDBrand)
        
        dim RowTitleBgColor
        if CurrentLast = "CURRENT" then
            RowTitleBgColor = Application("ColorCurrentYear")
        else
            RowTitleBgColor = Application("ColorLastYear")
        end if

        strOut1 = strOut1 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & " height=""30px;"">" & IDM_Oferta & "</td>"
        strOut2 = strOut2 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & " height=""30px;"">" & IDM_Folleto & "</td>"
        strOut3 = strOut3 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & " height=""30px;"">" & IDM_Cabecera & "</td>"
        strOut4 = strOut4 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & " height=""30px;"">" & IDM_NTiendas & "</td>"
        strOut6 = strOut6 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & " height=""30px;"">" & IDM_Adicional & "</td>"
        strOutKPI = strOutKPI & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & " height=""30px;"">" & IDM_KPIQuality & "</td>"
        strCalExp = strCalExp & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & " height=""30px;"">" & IDM_CalidadExp & "</td>"
        strCalOf = strCalOf & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & " height=""30px;"">" & IDM_CalidadOf & "</td>"
        
        strOutRD_NShops = strOutRD_NShops & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & IDM_NTiendasReal & "</td>"
        strOutTOTALNShops = strOutTOTALNShops & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & IDM_NTiendasTOTAL & "</td>"

        for iSubcat = 0 to 9
            strOutSubcatNShops(iSubcat) = ""
            strOutSubcatRD_NShops(iSubcat) = ""
            strOutSubcatTOTALNShops(iSubcat) = ""

            if bra.arrNShops(iSubcat) <> "" then
                strOutSubcatNShops(iSubcat) = strOutSubcatNShops(iSubcat) & "<tr><td class=gridtypetitleNoBold bgcolor=" & RowTitleBgColor & " height=""30px;"">" & bra.arrNShops(iSubcat) & "</td>"
                strOutSubcatRD_NShops(iSubcat) = strOutSubcatRD_NShops(iSubcat) & "<tr><td class=gridtypetitleNoBold bgcolor=" & RowTitleBgColor & " height=""30px;"">" & bra.arrNShops(iSubcat) & "</td>"
                strOutSubcatTOTALNShops(iSubcat) = strOutSubcatTOTALNShops(iSubcat) & "<tr><td class=gridtypetitleNoBold bgcolor=" & RowTitleBgColor & " height=""30px;"">" & bra.arrNShops(iSubcat) & "</td>"
            end if
        next
        
        ' Precarga las calidades de exposición
        set QExposicion = new QualityExp
        set QOferta = new QualityOf
        
        ' Para cada mes, busca info de cada quincena
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            arrActivity = getActivities(iYear, iMonth, IDClient, IDBrand)
            
            if CurrentLast = "CURRENT" then
                sClass = "Clk"
            else
                sClass = "cell"
            end if

            ' COLORES DE OFERTA ***************************************
            sClassOferta1 = sClass
            sClassOferta2 = sClass
            if arrActivity(0).Oferta<>"" then
                if CurrentLast = "CURRENT" then
                    sClassOferta1 = "OFCY"
                else
                    sClassOferta1 = "OFLY"
                end if
            end if
            if arrActivity(1).Oferta<>"" then
                if CurrentLast = "CURRENT" then
                    sClassOferta2 = "OFCY"
                else
                    sClassOferta2 = "OFLY"
                end if
            end if



            ' COLORES DE FOLLETO ***************************************
            sClassFolleto1 = sClass
            sClassFolleto2 = sClass
            if arrActivity(0).Folleto<>"" then
                if CurrentLast = "CURRENT" then
                    sClassFolleto1 = "FOCY"
                else
                    sClassFolleto1 = "FOLY"
                end if
            end if
            if arrActivity(1).Folleto<>"" then
                if CurrentLast = "CURRENT" then
                    sClassFolleto2 = "FOCY"
                else
                    sClassFolleto2 = "FOLY"
                end if
            end if


            ' COLORES DE CABECERA ***************************************
            sClassCabecera1 = sClass
            sClassCabecera2 = sClass
            BGColorCabecera1 = ""
            BGColorCabecera2 = ""
            if arrActivity(0).Cabecera<>"" then
                BGColorCabecera1 = arrActivity(0).RatioBackground
                if CurrentLast = "CURRENT" then
                    sClassCabecera1 = "CACY"
                else
                    sClassCabecera1 = "CALY"
                end if
            end if
            if arrActivity(1).Cabecera<>"" then
                BGColorCabecera2 = arrActivity(1).RatioBackground
                if CurrentLast = "CURRENT" then
                    sClassCabecera2 = "CACY"
                else
                    sClassCabecera2 = "CALY"
                end if
            end if

            
            ' COLORES DE NCENTROS ***************************************
            sClassCentros1 = sClass
            sClassCentros2 = sClass
            if arrActivity(0).NShops<>"" then
                if CurrentLast = "CURRENT" then
                    sClassCentros1 = "NSCY"
                else
                    sClassCentros1 = "NSLY"
                end if
            end if
            if arrActivity(1).NShops<>"" then
                if CurrentLast = "CURRENT" then
                    sClassCentros2 = "NSCY"
                else
                    sClassCentros2 = "NSLY"
                end if
            end if
            
            ' COLORES DE ADICIONAL ***************************************
            sClassAdicional1 = sClass
            sClassAdicional2 = sClass
            if arrActivity(0).Adicional<>"" then
                if CurrentLast = "CURRENT" then
                    sClassAdicional1 = "ADCY"
                else
                    sClassAdicional1 = "ADLY"
                end if
            end if
            if arrActivity(1).Adicional<>"" then
                if CurrentLast = "CURRENT" then
                    sClassAdicional2 = "ADCY"
                else
                    sClassAdicional2 = "ADLY"
                end if
            end if
            
            ' COLORES DE KPI QUALITY ***************************************
            sClassKpiQuality1 = sClass
            sClassKpiQuality2 = sClass
            if NOT isNull(arrActivity(0).KPIQuality) AND arrActivity(0).KPIQuality <> -1 then
                if CurrentLast = "CURRENT" then
                    sClassKpiQuality1 = "KPICY"
                else
                    sClassKpiQuality1 = "KPILY"
                end if
            end if
            if NOT isNull(arrActivity(1).KPIQuality) AND arrActivity(1).KPIQuality <> -1 then
                if CurrentLast = "CURRENT" then
                    sClassKpiQuality2 = "KPICY"
                else
                    sClassKpiQuality2 = "KPILY"
                end if
            end if

            
            FGColorOferta1 = Application("ColorFGStatus" & arrActivity(0).IDStatus)
            FGColorOferta2 = Application("ColorFGStatus" & arrActivity(1).IDStatus)
            FGColorFolleto1 = Application("ColorFGStatus" & arrActivity(0).IDStatus)
            FGColorFolleto2 = Application("ColorFGStatus" & arrActivity(1).IDStatus)
            FGColorCabecera1 = Application("ColorFGStatus" & arrActivity(0).IDStatus)
            FGColorCabecera2 = Application("ColorFGStatus" & arrActivity(1).IDStatus)
            FGColorCentros1 = "Black"
            FGColorCentros2 = "Black"
            

            sOnClick1 = ""
            sOnClick2 = ""
            if CurrentLast = "CURRENT" then
                sOnClick1 = "editAct('" & arrActivity(0).ID & "','" & IDClient & "','" & IDBrand & "','" & iYear & "','" & iMonth & "',1)"
                sOnClick2 = "editAct('" & arrActivity(1).ID & "','" & IDClient & "','" & IDBrand & "','" & iYear & "','" & iMonth & "',2)"
            end if


            ' *******************************
            ' Fila Oferta
            strOut1 = strOut1 & RenderCell(false, "", sClassOferta1, "", "", "", FGColorOferta1, "", arrActivity(0).Oferta, sOnClick1, "")
            strOut1 = strOut1 & RenderCell(false, "", sClassOferta2, "", "", "", FGColorOferta2, "", arrActivity(1).Oferta, sOnClick2, "")

            ' *******************************
            ' Fila Folleto
            strOut2 = strOut2 & RenderCell(false, "", sClassFolleto1, "", "", "", FGColorFolleto1, "", arrActivity(0).Folleto, sOnClick1, "")
            strOut2 = strOut2 & RenderCell(false, "", sClassFolleto2, "", "", "", FGColorFolleto2, "", arrActivity(1).Folleto, sOnClick2, "")

            ' *******************************
            ' Fila Cabecera
            strOut3 = strOut3 & RenderCell(false, "", sClassCabecera1, "", "", "center", FGColorCabecera1, BGColorCabecera1, arrActivity(0).Cabecera, sOnClick1, "")
            strOut3 = strOut3 & RenderCell(false, "", sClassCabecera2, "", "", "center", FGColorCabecera2, BGColorCabecera2, arrActivity(1).Cabecera, sOnClick2, "")

            ' *******************************
            ' Fila NShops
            strOut4 = strOut4 & RenderCell(false, "", sClassCentros1, "", "", "", FGColorCentros1, "", arrActivity(0).NShops, sOnClick1, "")
            strOut4 = strOut4 & RenderCell(false, "", sClassCentros2, "", "", "", FGColorCentros2, "", arrActivity(1).NShops, sOnClick2, "")
            
            
            ' *******************************
            ' Filas NShops, RD_NShops y TOTALNShops para cada Subcategoría
            for iSubcat = 0 to 9
                if bra.arrNShops(iSubcat) <> "" then
                    ' NShops
                    strOutSubcatNShops(iSubcat) = strOutSubcatNShops(iSubcat) & RenderCell(false, "", sClassCentros1, "", "", "", FGColorCentros1, "", arrActivity(0).arrNShops(iSubcat), sOnClick1, "")
                    strOutSubcatNShops(iSubcat) = strOutSubcatNShops(iSubcat) & RenderCell(false, "", sClassCentros2, "", "", "", FGColorCentros2, "", arrActivity(1).arrNShops(iSubcat), sOnClick2, "")
                    
                    ' RD_NShops
                    strOutSubcatRD_NShops(iSubcat) = strOutSubcatRD_NShops(iSubcat) & RenderCell(false, "", sClassCentros1, "", "", "", FGColorCentros1, "", arrActivity(0).arrRD_NShops(iSubcat), sOnClick1, "")
                    strOutSubcatRD_NShops(iSubcat) = strOutSubcatRD_NShops(iSubcat) & RenderCell(false, "", sClassCentros2, "", "", "", FGColorCentros2, "", arrActivity(1).arrRD_NShops(iSubcat), sOnClick2, "")

                    ' TOTAL NShops
                    strOutSubcatTOTALNShops(iSubcat) = strOutSubcatTOTALNShops(iSubcat) & RenderCell(false, "", sClassCentros1, "", "", "", FGColorCentros1, "", arrActivity(0).arrTOTALNShops(iSubcat), sOnClick1, "")
                    strOutSubcatTOTALNShops(iSubcat) = strOutSubcatTOTALNShops(iSubcat) & RenderCell(false, "", sClassCentros2, "", "", "", FGColorCentros2, "", arrActivity(1).arrTOTALNShops(iSubcat), sOnClick2, "")

                end if
            next


            ' *******************************
            ' Fila Adicional
            strOut6 = strOut6 & RenderCell(false, "", sClassAdicional1, "", "", "", FGColorAdicional1, "", arrActivity(0).Adicional, sOnClick1, "")
            strOut6 = strOut6 & RenderCell(false, "", sClassAdicional2, "", "", "", FGColorAdicional2, "", arrActivity(1).Adicional, sOnClick2, "")
            
            ' *******************************
            ' KPI Calidad
             dim sKPI1: sKPI1 = ""
             dim sKPI2: sKPI2 = ""
            FGColorKpiQuality1 = ""
            FGColorKpiQuality2 = ""
            if NOT isNull(arrActivity(0).KPIQuality) AND arrActivity(0).KPIQuality > -1 then
                sKPI1 = arrActivity(0).KPIQuality
                FGColorKpiQuality1 = "Black"
            else
                sClassKpiQuality1 = sClass
            end if
            if NOT isNull(arrActivity(1).KPIQuality) AND arrActivity(1).KPIQuality > -1 then
                sKPI2 = arrActivity(1).KPIQuality
                FGColorKpiQuality2 = "Black"
            else
                sClassKpiQuality2 = sClass
            end if
            strOutKPI = strOutKPI & RenderCell(false, "", sClassKpiQuality1, "", "", "center", FGColorKpiQuality1, "", sKPI1, sOnClick1, "")
            strOutKPI = strOutKPI & RenderCell(false, "", sClassKpiQuality2, "", "", "center", FGColorKpiQuality2, "", sKPI2, sOnClick2, "")
            
            
            dim sTipoSelector, sDescCalExp, sDescCalOf
            sTipoSelector = "SELECT"  ' RADIO / SELECT / DIV
            
            ' *******************************
            ' Calidad Exposición
            if CurrentLast = "CURRENT" then
                if arrActivity(0).ID > -1 then
                    if sTipoSelector = "RADIO" then
                        strCalExp = strCalExp & "<td class=""cell"">&nbsp;"
                        for each q in QExposicion.arrQuality
                            strCalExp = strCalExp & "<br><input onmousedown=""saveCalExp(this);"" type=radio alt=""" & q.ID & """ name=""CE_" & Right("0000" & arrActivity(0).ID, 5) & """ value=""" & q.ID & """"
                            if arrActivity(0).IDCalidadExp = q.ID then
                                strCalExp = strCalExp & " checked"
                            end if
                            strCalExp = strCalExp & ">" & q.Descripcion
                        next
                        strCalExp = strCalExp & "</td>"
                    elseif sTipoSelector = "DIV" then
                        strCalExp = strCalExp & "<td class=""cell"">"
                        sDescCalExp = arrActivity(0).DesCalidadExp
                        if arrActivity(0).DesCalidadExp = "" then
                            sDescCalExp = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        end if
                        strCalExp = strCalExp & "<label ID=""LCE_" & Right("0000" & arrActivity(0).ID, 5) & """ onclick=""valorarCalExp(CE_" & Right("0000" & arrActivity(0).ID, 5) & ");"" style=""cursor:pointer;"">" & sDescCalExp & "</label>"
                        strCalExp = strCalExp & "<br><select style=""display:none;"" alt=""" & arrActivity(0).IDCalidadExp & """ onchange=""saveCalExp(this, LCE_" & Right("0000" & arrActivity(0).ID, 5) & ");"" name=""CE_" & Right("0000" & arrActivity(0).ID, 5) & """ >"
                        strCalExp = strCalExp & "<option value="""">Cerrar</option>"
                        for each q in QExposicion.arrQuality
                            strCalExp = strCalExp & "<option value=""" & q.ID & """"
                            if arrActivity(0).IDCalidadExp = q.ID then
                                strCalExp = strCalExp & " selected"
                            end if
                            strCalExp = strCalExp & ">" & q.Descripcion
                            strCalExp = strCalExp & "</option>"
                        next
                        strCalExp = strCalExp & "</select></td>"
                    else
                        strCalExp = strCalExp & "<td class=cell>&nbsp;"
                        strCalExp = strCalExp & "<select onchange=""saveCalExp(this);"" name=""CE_" & Right("0000" & arrActivity(0).ID, 5) & """ >"
                        for each q in QExposicion.arrQuality
                            strCalExp = strCalExp & "<option value=""" & q.ID & """"
                            if arrActivity(0).IDCalidadExp = q.ID then
                                strCalExp = strCalExp & " selected"
                            end if
                            strCalExp = strCalExp & ">" & q.Descripcion
                            strCalExp = strCalExp & "</option>"
                        next
                        strCalExp = strCalExp & "</select></td>"
                    end if
                else
                    strCalExp = strCalExp & "<td class=""cell"">&nbsp;"
                    strCalExp = strCalExp & "</td>"
                end if

                if arrActivity(1).ID > -1 then
                    if sTipoSelector = "RADIO" then
                        strCalExp = strCalExp & "<td class=""cell"">&nbsp;"
                        for each q in QExposicion.arrQuality
                            strCalExp = strCalExp & "<br><input onmousedown=""saveCalExp(this);"" type=radio alt=""" & q.ID & """ name=""CE_" & Right("0000" & arrActivity(1).ID, 5) & """ value=""" & q.ID & """"
                            if arrActivity(1).IDCalidadExp = q.ID then
                                strCalExp = strCalExp & " checked"
                            end if
                            strCalExp = strCalExp & ">" & q.Descripcion
                        next
                        strCalExp = strCalExp & "</td>"
                    elseif sTipoSelector = "DIV" then
                        strCalExp = strCalExp & "<td class=""cell"">"
                        sDescCalExp = arrActivity(1).DesCalidadExp
                        if arrActivity(1).DesCalidadExp = "" then
                            sDescCalExp = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        end if
                        strCalExp = strCalExp & "<label ID=""LCE_" & Right("0000" & arrActivity(1).ID, 5) & """ onclick=""valorarCalExp(CE_" & Right("0000" & arrActivity(1).ID, 5) & ");"" style=""cursor:pointer;"">" & sDescCalExp & "</label>"
                        strCalExp = strCalExp & "<br><select style=""display:none;"" alt=""" & arrActivity(1).IDCalidadExp & """ onchange=""saveCalExp(this, LCE_" & Right("0000" & arrActivity(1).ID, 5) & ");"" name=""CE_" & Right("0000" & arrActivity(1).ID, 5) & """ >"
                        strCalExp = strCalExp & "<option value="""">Cerrar</option>"
                        for each q in QExposicion.arrQuality
                            strCalExp = strCalExp & "<option value=""" & q.ID & """"
                            if arrActivity(1).IDCalidadExp = q.ID then
                                strCalExp = strCalExp & " selected"
                            end if
                            strCalExp = strCalExp & ">" & q.Descripcion
                            strCalExp = strCalExp & "</option>"
                        next
                        strCalExp = strCalExp & "</select></td>"
                    else
                        strCalExp = strCalExp & "<td class=""cell"">&nbsp;"
                        strCalExp = strCalExp & "<select onchange=""saveCalExp(this);"" name=""CE_" & Right("0000" & arrActivity(1).ID, 5) & """ >"
                        for each q in QExposicion.arrQuality
                            strCalExp = strCalExp & "<option value=""" & q.ID & """"
                            if arrActivity(1).IDCalidadExp = q.ID then
                                strCalExp = strCalExp & " selected"
                            end if
                            strCalExp = strCalExp & ">" & q.Descripcion
                            strCalExp = strCalExp & "</option>"
                        next
                        strCalExp = strCalExp & "</select></td>"
                    end if
                else
                    strCalExp = strCalExp & "<td class=""cell"">&nbsp;"
                    strCalExp = strCalExp & "</td>"
                end if
            else
                strCalExp = strCalExp & "<td class=""cell"">" & arrActivity(0).DesCalidadExp & "&nbsp;</td>"
                strCalExp = strCalExp & "<td class=""cell"">" & arrActivity(1).DesCalidadExp & "&nbsp;</td>"
            end if
            
            
            ' *******************************
            ' Calidad Oferta
            
            if CurrentLast = "CURRENT" then
                if arrActivity(0).ID > -1 then
                    if sTipoSelector = "RADIO" then
                        strCalOf = strCalOf & "<td class=""cell"">&nbsp;"
                        for each q in QOferta.arrQuality
                            strCalOf = strCalOf & "<br><input onmousedown=""saveCalOf(this);"" type=radio alt=""" & q.ID & """ name=""CO_" & Right("0000" & arrActivity(0).ID, 5) & """ value=""" & q.ID & """"
                            if arrActivity(0).IDCalidadOf = q.ID then
                                strCalOf = strCalOf & " checked"
                            end if
                            strCalOf = strCalOf & ">" & q.Descripcion
                        next
                        strCalOf = strCalOf & "</td>"
                    elseif sTipoSelector = "DIV" then
                        strCalOf = strCalOf & "<td class=""cell"">"
                        sDescCalOf = arrActivity(0).DesCalidadOf
                        if arrActivity(0).DesCalidadOf = "" then
                            sDescCalOf = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        end if
                        strCalOf = strCalOf & "<label ID=""LCO_" & Right("0000" & arrActivity(0).ID, 5) & """ onclick=""valorarCalOf(CO_" & Right("0000" & arrActivity(0).ID, 5) & ");"" style=""cursor:pointer;"">" & sDescCalOf & "</label>"
                        strCalOf = strCalOf & "<br><select style=""display:none;"" alt=""" & arrActivity(0).IDCalidadOf & """ onchange=""saveCalOf(this, LCO_" & Right("0000" & arrActivity(0).ID, 5) & ");"" name=""CO_" & Right("0000" & arrActivity(0).ID, 5) & """ >"
                        strCalOf = strCalOf & "<option value="""">Cerrar</option>"
                        for each q in QOferta.arrQuality
                            strCalOf = strCalOf & "<option value=""" & q.ID & """"
                            if arrActivity(0).IDCalidadOf = q.ID then
                                strCalOf = strCalOf & " selected"
                            end if
                            strCalOf = strCalOf & ">" & q.Descripcion
                            strCalOf = strCalOf & "</option>"
                        next
                        strCalOf = strCalOf & "</select></td>"
                    else
                        strCalOf = strCalOf & "<td class=""cell"">&nbsp;"
                        strCalOf = strCalOf & "<select onchange=""saveCalOf(this);"" name=""CO_" & Right("0000" & arrActivity(0).ID, 5) & """ >"
                        for each q in QOferta.arrQuality
                            strCalOf = strCalOf & "<option value=""" & q.ID & """"
                            if arrActivity(0).IDCalidadOf = q.ID then
                                strCalOf = strCalOf & " selected"
                            end if
                            strCalOf = strCalOf & ">" & q.Descripcion
                            strCalOf = strCalOf & "</option>"
                        next
                        strCalOf = strCalOf & "</select></td>"
                    end if
                else
                    strCalOf = strCalOf & "<td class=""cell"">&nbsp;"
                    strCalOf = strCalOf & "</td>"
                end if

                if arrActivity(1).ID > -1 then
                    if sTipoSelector = "RADIO" then
                        strCalOf = strCalOf & "<td class=""cell"">&nbsp;"
                        for each q in QOferta.arrQuality
                            strCalOf = strCalOf & "<br><input onmousedown=""saveCalOf(this);"" type=radio alt=""" & q.ID & """ name=""CO_" & Right("0000" & arrActivity(1).ID, 5) & """ value=""" & q.ID & """"
                            if arrActivity(1).IDCalidadOf = q.ID then
                                strCalOf = strCalOf & " checked"
                            end if
                            strCalOf = strCalOf & ">" & q.Descripcion
                        next
                        strCalOf = strCalOf & "</td>"
                    elseif sTipoSelector = "DIV" then
                        strCalOf = strCalOf & "<td class=""cell"">"
                        sDescCalOf = arrActivity(1).DesCalidadOf
                        if arrActivity(1).DesCalidadOf = "" then
                            sDescCalOf = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        end if
                        strCalOf = strCalOf & "<label ID=""LCO_" & Right("0000" & arrActivity(1).ID, 5) & """ onclick=""valorarCalOf(CO_" & Right("0000" & arrActivity(1).ID, 5) & ");"" style=""cursor:pointer;"">" & sDescCalOf & "</label>"
                        strCalOf = strCalOf & "<br><select style=""display:none;"" alt=""" & arrActivity(1).IDCalidadOf & """ onchange=""saveCalOf(this, LCO_" & Right("0000" & arrActivity(1).ID, 5) & ");"" name=""CO_" & Right("0000" & arrActivity(1).ID, 5) & """ >"
                        strCalOf = strCalOf & "<option value="""">Cerrar</option>"
                        for each q in QOferta.arrQuality
                            strCalOf = strCalOf & "<option value=""" & q.ID & """"
                            if arrActivity(1).IDCalidadOf = q.ID then
                                strCalOf = strCalOf & " selected"
                            end if
                            strCalOf = strCalOf & ">" & q.Descripcion
                            strCalOf = strCalOf & "</option>"
                        next
                        strCalOf = strCalOf & "</select></td>"
                    else
                        strCalOf = strCalOf & "<td class=""cell"">&nbsp;"
                        strCalOf = strCalOf & "<select onchange=""saveCalOf(this);"" name=""CO_" & Right("0000" & arrActivity(1).ID, 5) & """ >"
                        for each q in QOferta.arrQuality
                            strCalOf = strCalOf & "<option value=""" & q.ID & """"
                            if arrActivity(1).IDCalidadOf = q.ID then
                                strCalOf = strCalOf & " selected"
                            end if
                            strCalOf = strCalOf & ">" & q.Descripcion
                            strCalOf = strCalOf & "</option>"
                        next
                        strCalOf = strCalOf & "</select></td>"
                    end if
                else
                    strCalOf = strCalOf & "<td class=""cell"">&nbsp;"
                    strCalOf = strCalOf & "</td>"
                end if
            else
                strCalOf = strCalOf & "<td class=""cell"">" & arrActivity(0).DesCalidadOf & "&nbsp;</td>"
                strCalOf = strCalOf & "<td class=""cell"">" & arrActivity(1).DesCalidadOf & "&nbsp;</td>"
            end if


            ' *******************************
            ' Fila NShops Real
            strOutRD_NShops = strOutRD_NShops & RenderCell(false, "", sClassCentros1, "", "", "", FGColorCentros1, "", arrActivity(0).RD_NShops, sOnClick1, "")
            strOutRD_NShops = strOutRD_NShops & RenderCell(false, "", sClassCentros2, "", "", "", FGColorCentros2, "", arrActivity(1).RD_NShops, sOnClick2, "")
            
            ' *******************************
            ' Fila TOTAL NShops
            strOutTOTALNShops = strOutTOTALNShops & RenderCell(false, "", sClassCentros1, "", "", "", FGColorCentros1, "", arrActivity(0).TOTALNShops, sOnClick1, "")
            strOutTOTALNShops = strOutTOTALNShops & RenderCell(false, "", sClassCentros2, "", "", "", FGColorCentros2, "", arrActivity(1).TOTALNShops, sOnClick2, "")
            
        next
        
        
        
        ' CERRAMOS FILAS
        strOut1 = strOut1 & "</tr>"
        strOut2 = strOut2 & "</tr>"
        strOut3 = strOut3 & "</tr>"
        strOut4 = strOut4 & "</tr>"
        for iSubcat = 0 to 9
            if bra.arrNShops(iSubcat) <> "" then
                strOutSubcatNShops(iSubcat) = strOutSubcatNShops(iSubcat) & "</tr>"
            end if
        next
        strOut6 = strOut6 & "</tr>"
        strOutKPI = strOutKPI & "</tr>"
        strCalExp = strCalExp & "</tr>"
        strCalOf = strCalOf & "</tr>"
        strOutRD_NShops = strOutRD_NShops & "</tr>"
        strOutTOTALNShops = strOutTOTALNShops & "</tr>"
        
        
        ' Monta la salida total
        strOut = strOut1 & strOut2 & strOut3 & strOut4 
        for iSubcat = 0 to 9
            if bra.arrNShops(iSubcat) <> "" then
                strOut = strOut & strOutSubcatNShops(iSubcat)
            end if
        next
        strOut = strOut & strOut6
        
        if Request.Form("FILTER_SHOWKPIQUALITY") <> "" then
            strOut = strOut & strOutKPI
        end if
        
        
        if Request.Form("FILTER_SHOWQUALITY") <> "" AND IsInputQuality() then
            strOut = strOut & strCalExp & strCalOf
        end if

        if Request.Form("FILTER_SHOWREALDATA_NSHOPS")<>"" then
            strOut = strOut & strOutRD_NShops
            for iSubcat = 0 to 9
                if bra.arrNShops(iSubcat) <> "" then
                    strOut = strOut & strOutSubcatRD_NShops(iSubcat)
                end if
            next
        end if

        if Request.Form("FILTER_TOTALSHOPS")<>"" then
            strOut = strOut & strOutTOTALNShops
            for iSubcat = 0 to 9
                if bra.arrNShops(iSubcat) <> "" then
                    strOut = strOut & strOutSubcatTOTALNShops(iSubcat)
                end if
            next
        end if
        
        
        
        
        PintarActivity = strOut
    End Function
    
    

    
    Function PintarFilaBlanco(StartYear, StartMonth, ViewMonths)
        dim strOut
        strOut = ""
        
        dim i, iMonth, iYear


        strOut = strOut & "<tr height=1px><td style=""border-top:1px solid black;border-bottom:1px solid black;""><font class=font8>&nbsp;</font></td><td style=""border-top:1px solid black;border-bottom:1px solid black;""><font class=font8>&nbsp;</font></td><td style=""border-top:1px solid black;border-bottom:1px solid black;""><font class=font8>&nbsp;</font></td>"

        ' Para cada mes, busca info de cada quincena
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if

            strOut = strOut & "<td style=""border-top:1px solid black;border-bottom:1px solid black;""><font class=font8>&nbsp;</font></td>"
            strOut = strOut & "<td style=""border-top:1px solid black;border-bottom:1px solid black;""><font class=font8>&nbsp;</font></td>"
            
            
        next
        
        strOut = strOut & "</tr>"
        
        PintarFilaBlanco = strOut
    End Function
    
    
    Function PintarReportClientBrand(StartYear, StartMonth, ViewMonths, client, brand, YearRowSpan, ReportNumRowsPerYear, RowTitle)
        dim strOut
        strOut = ""
        
        strOut = strOut & PintarGrupo0(2+YearRowSpan, RowTitle, client.IDClient & "_" & brand.IDBrand)
        strOut = strOut & PintarGrupo05(1+ReportNumRowsPerYear, StartYear, Application("ColorCurrentYear"))
        if Request.Form("FILTER_SHOWGENERALTHEME")<>"" then
            strOut = strOut & PintarGeneralTheme(client.IDClient, StartYear, StartMonth, ViewMonths, Application("ColorCurrentYear"), "CURRENT")
        end if
        
        strOut = strOut & PintarActivity(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, "CURRENT")
        
        if Request.Form("FILTER_SHOWNR")<>"" then
            strOut = strOut & PintarNetRevenue(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorCurrentYear"), Application("ColorNR_CY"))
        end if
        if Request.Form("FILTER_SHOWFC")<>"" then
            strOut = strOut & PintarForecast(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorCurrentYear"), Application("ColorFC_CY"))
        end if
        if Request.Form("FILTER_SHOWNRVSLY")<>"" then
            strOut = strOut & PintarNetRevenueVsLY(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorCurrentYear"), Application("ColorNRvsLY_CY"))
        end if
        
        if Request.Form("FILTER_LASTYEAR")<>"" then
            strOut = strOut & PintarGrupo05(1+ReportNumRowsPerYear, StartYear-1, Application("ColorLastYear"))
            if Request.Form("FILTER_SHOWGENERALTHEME")<>"" then
                strOut = strOut & PintarGeneralTheme(client.IDClient, StartYear-1, StartMonth, ViewMonths, Application("ColorLastYear"), "LAST")
            end if
            
            strOut = strOut & PintarActivity(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, "LAST")

            if Request.Form("FILTER_SHOWNR")<>"" then
                strOut = strOut & PintarNetRevenue(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorLastYear"), Application("ColorNR_LY"))
            end if
            if Request.Form("FILTER_SHOWFC")<>"" then
                strOut = strOut & PintarForecast(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorLastYear"), Application("ColorFC_LY"))
            end if
            if Request.Form("FILTER_SHOWNRVSLY")<>"" then
                strOut = strOut & PintarNetRevenueVsLY(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorLastYear"), Application("ColorNRvsLY_LY"))
            end if
            
        end if
        
        
        
        PintarReportClientBrand = strOut
    End Function
    
    
    
        
    Function PintarNetRevenue(StartYear, StartMonth, ViewMonths, IDClient, IDBrand, RowTitleBgColor, CellBgColor)
        dim strOut
        strOut = ""
        
        dim NR, i, iMonth, iYear
        
        strOut = strOut & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">NR</td>"

        ' Para cada mes, busca info
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            NR = getNR(iYear, iMonth, IDClient, IDBrand)
            
            strOut = strOut & RenderCell(false, "2", "cell", "", "", "center", "", CellBgColor, FormatoNum(NR, true, 2), "", "")
            
            
        next
        
        strOut = strOut & "</tr>"
        
        PintarNetRevenue = strOut
    End Function
    
    
    
    Function PintarNetRevenueVsLY(StartYear, StartMonth, ViewMonths, IDClient, IDBrand, RowTitleBgColor, CellBgColor)
        dim strOut
        strOut = ""
        
        dim Percent, NR, NRLY, i, iMonth, iYear, sFGColor
        
        strOut = strOut & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">%NR vs LY</td>"

        ' Para cada mes, busca info
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            NR = getNR(iYear, iMonth, IDClient, IDBrand)
            NRLY = getNR(iYear-1, iMonth, IDClient, IDBrand)
            
            Percent = 0
            sFGColor = "black"
            if NRLY>0 then
                Percent = NR * 100 / NRLY
            end if
            if Percent>0 AND Percent<100 then
                Percent = - 100 + Percent
                sFGColor = "red"
            elseif Percent>100 then
                Percent = Percent - 100
            end if
            
            strOut = strOut & RenderCell(false, "2", "cell", "", "", "center", sFGColor, CellBgColor, FormatoNum(Percent, false, 2) & "%", "", "")
            
            
        next
        
        strOut = strOut & "</tr>"
        
        PintarNetRevenueVsLY = strOut
    End Function
    
    Function PintarForecast(StartYear, StartMonth, ViewMonths, IDClient, IDBrand, RowTitleBgColor, CellBgColor)
        dim strOut
        strOut = ""
        
        dim FC, i, iMonth, iYear
        
        strOut = strOut & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">FC</td>"

        ' Para cada mes, busca info
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            FC = getForecast(iYear, iMonth, IDClient, IDBrand)
            
            strOut = strOut & RenderCell(false, "2", "cell", "", "", "center", "", CellBgColor, FormatoNum(FC, true, 2), "", "")
            
            
        next
        
        strOut = strOut & "</tr>"
        
        PintarForecast = strOut
    End Function
    
    
    ' Pinta las actividades de un tipo en el calendario
    Function PintarGeneralTheme(IDClient, StartYear, StartMonth, ViewMonths, RowTitleBgColor, CurrentLast)
        dim arrGeneralThemes
        dim strOut
        dim ClassName
        dim sClass1, sClass2, sOnClick1, sOnClick2
        dim i, iMonth, iYear
        
        strOut = strOut & "<tr><td class=gridtypetitle bgcolor=""" & RowTitleBgColor & """ height=""30px;"">" & IDM_GeneralTheme & "</td>"
        
        ' Para cada mes, busca info de cada quincena
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            arrGeneralThemes = getGeneralThemes(IDClient, iYear, iMonth)
            
            dim sImageTag1, sImageTag2
            if arrGeneralThemes(0).ThemeImageFileName <> "" then
                sImageTag1 = "<img alt=""" & arrGeneralThemes(0).GridText & """ src=""images/Themes/" & arrGeneralThemes(0).ThemeImageFileName & """ width=""" & Application("ThemeImageWidth") & "px"" />"
            else
                sImageTag1 = ""
            end if
            if arrGeneralThemes(1).ThemeImageFileName <> "" then
                sImageTag2 = "<img alt=""" & arrGeneralThemes(1).GridText & """ src=""images/Themes/" & arrGeneralThemes(1).ThemeImageFileName & """ width=""" & Application("ThemeImageWidth") & "px"" />"
            else
                sImageTag2 = ""
            end if
            
            if CurrentLast = "CURRENT" then
                sClass1 = "Clk"
                sClass2 = "Clk"
                ClassName = "TGCY"
            else
                sClass1 = "cell"
                sClass2 = "cell"
                ClassName = "TGLY"
            end if
            if Trim(arrGeneralThemes(0).GridText)<>"" then
                sClass1 = ClassName
            end if
            if Trim(arrGeneralThemes(1).GridText)<>"" then
                sClass2 = ClassName
            end if
            
            sOnClick1 = ""
            sOnClick2 = ""
            if CurrentLast = "CURRENT" then
                sOnClick1 = "editGenThem('SOAGeneralTheme.asp', '" & arrGeneralThemes(0).ID & "', '" & IDClient & "', '" & iYear & "', '" & iMonth & "', 1)"
                sOnClick2 = "editGenThem('SOAGeneralTheme.asp', '" & arrGeneralThemes(0).ID & "', '" & IDClient & "', '" & iYear & "', '" & iMonth & "', 2)"
            end if
            strOut = strOut & RenderCell(false, "", sClass1, "", "", "", "", "", arrGeneralThemes(0).GridText, sOnClick1, sImageTag1)
            strOut = strOut & RenderCell(false, "", sClass2, "", "", "", "", "", arrGeneralThemes(1).GridText, sOnClick2, sImageTag2)

        next
        
        strOut = strOut & "</tr>"

        
        PintarGeneralTheme = strOut
    End Function    
    
    
    ' Pinta una casilla del report
    Function RenderCell(Bold, Colspan, ClassName, Width, Height, Align, ForeColor, BackColor, Text, jsOnClick, ImageTag)
        dim strOut, strStyle
        strOut = ""
        strStyle = ""
        
        strOut = strOut & "<TD"
        if Colspan<>"" then
            strOut = strOut & " colspan=""" & Colspan & """"
        end if
        if ClassName<>"" then
            strOut = strOut & " class=""" & ClassName & """"
        end if
        if Align<>"" then
            strOut = strOut & " align=""" & Align & """"
        end if
        
        
        if ForeColor<>"" then
            strStyle = strStyle & "color:" & ForeColor & ";"
        end if
        if Width<>"" then
            strStyle = strStyle & "width:" & Width & ";"
        end if
        if Height<>"" then
            strStyle = strStyle & "height:" & Height & ";"
        end if
        if Bold then
            strStyle = strStyle & "font-weight:bold;"
        end if
        if BackColor<>"" then
            strStyle = strStyle & "background-color:" & BackColor & ";"
        end if
        
        if strStyle<>"" then
            strOut = strOut & " style=""" & strStyle & """"
        end if
        
        
        if jsOnClick<>"" then
            strOut = strOut & " onclick=""" & jsOnClick & """"
            'strOut = strOut & " onmouseover=""this.className='cellhover';"" onmouseout=""this.className='" & ClassName & "'"" "
        end if
        
        strOut = strOut & ">"
        if Trim(ImageTag) <> "" then
            strOut = strOut & ImageTag
        else
            if Trim(Text)<>"" then
                if Request("XL")<>"" then
                    if isNumeric(Text) then
                        strOut = strOut & "=" & server.HTMLEncode(Text)
                    else
                        strOut = strOut & server.HTMLEncode(Text)
                    end if
                else
                    strOut = strOut & server.HTMLEncode(Text)
                end if
            else
                strOut = strOut & "&nbsp;"
            end if
        end if
        strOut = strOut & "</TD>"
        
        RenderCell = strOut
    End Function
    
    Function PintarColumnasRealData(startYear, StartMonth, ViewMonths)
        dim strOut
        
        dim i, iMonth, iYear
        
        strOut = strOut & "<tr><td class=gridtypetitle>&nbsp;</td><td style=""border-top:1px solid gray;"" >&nbsp;</td><td style=""border-top:1px solid gray;"">&nbsp;</td>"

        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            'Half1
            strOut = strOut & "<td align=center class=gridhalf1title>"
            
            if Request.Form("ViewNShopsActivity")<>"" then
                strOut = strOut & "<input type=text style=""width:60px;border:0px;text-align:center;"" readonly value=""" & IDM_NTiendasShort & """ />"
            end if
            if Request.Form("ViewNShops")<>"" then
                strOut = strOut & "<input type=text style=""width:60px;border:0px;text-align:center;"" readonly value=""" & IDM_NTiendasRealShort & """ />"
            end if
            if Request.Form("ViewPercentComplaint")<>"" then
                strOut = strOut & "<input type=text style=""width:60px;border:0px;text-align:center;"" readonly value=""" & IDM_NTiendasTotal & """ />"
            end if
            strOut = strOut & "</td>"

            'Half2
            strOut = strOut & "<td align=center class=gridhalf2title>"
            if Request.Form("ViewNShopsActivity")<>"" then
                strOut = strOut & "<input type=text style=""width:60px;border:0px;text-align:center;"" readonly value=""" & IDM_NTiendasShort & """ />"
            end if
            if Request.Form("ViewNShops")<>"" then
                strOut = strOut & "<input type=text style=""width:60px;border:0px;text-align:center;"" readonly value=""" & IDM_NTiendasRealShort & """ />"
            end if
            if Request.Form("ViewPercentComplaint")<>"" then
                strOut = strOut & "<input type=text style=""width:60px;border:0px;text-align:center;"" readonly value=""" & IDM_NTiendasTotal & """ />"
            end if
            strOut = strOut & "</td>"
            
        next
        
        strOut = strOut & "</tr>"

        PintarColumnasRealData = strOut
    End Function
    
    
    Function PintarRealData(client, brand, StartYear, StartMonth, ViewMonths, brandIteration, nSubCategory, DescSubCategory)
        dim strOut
        dim act1, act2
        dim SQL, rst
        dim i, iMonth, iYear, iter
        dim BrandName
        dim NShops1, NShops2
        dim NShopsActivity1, NShopsActivity2
        dim NShopsTotal1, NShopsTotal2
        dim jsAutofill


        if DescSubCategory <> "" then
            BrandName = "&nbsp;&nbsp;&nbsp;" & DescSubCategory
        else
            BrandName = bra.Name
        end if
        
        strOut = strOut & "<tr><td class=gridtypetitle>" & BrandName & "</td><td style=""border-top:1px solid gray;"">&nbsp;</td><td style=""border-top:1px solid gray;"">&nbsp;</td>"

        iMonth = StartMonth
        iYear = StartYear
        iter = 10
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            
            ' Lanzamos un insert por cada valor por si no existe el registro
            ' Luego el SAVE hace sólo un UPDATE
            on error resume next
            SQL = "INSERT INTO RealData (WYear, IDClient, IDBrand, WMonth, WHalf, NShops" & nSubCategory & ") " & _
            " VALUES (" & iYear & ", " & client.IDClient & ", " & brand.IDBrand & ", " & iMonth & ", 1, NULL)"
            ObjConnectionSQL.Execute(SQL)
            
            SQL = "INSERT INTO RealData (WYear, IDClient, IDBrand, WMonth, WHalf, NShops" & nSubCategory & ") " & _
            " VALUES (" & iYear & ", " & client.IDClient & ", " & brand.IDBrand & ", " & iMonth & ", 2, NULL)"
            ObjConnectionSQL.Execute(SQL)
            on error goto 0
            
            
            ' Carga las dos actividades del mes. Si no hay actividad, igualmente cargará NShops real
            set act1 = getActivityFromDate(client.IDClient, brand.IDBrand, iYear, iMonth, 1)
            set act2 = getActivityFromDate(client.IDClient, brand.IDBrand, iYear, iMonth, 2)

            NShops1 = act1.RD_NShops
            if nSubCategory <> "" then
                NShops1 = act1.arrRD_NShops(CInt(nSubCategory))
            end if
            NShops2 = act2.RD_NShops
            if nSubCategory <> "" then
                NShops2 = act2.arrRD_NShops(CInt(nSubCategory))
            end if
            
            NShopsActivity1 = act1.NShops
            if nSubCategory <> "" then
                NShopsActivity1 = act1.arrNShops(CInt(nSubCategory))
            end if
            NShopsActivity2 = act2.NShops
            if nSubCategory <> "" then
                NShopsActivity2 = act2.arrNShops(CInt(nSubCategory))
            end if
            
            NShopsTotal1 = act1.TotalNShops
            if nSubCategory <> "" then
                NShopsTotal1 = act1.arrTOTALNShops(CInt(nSubCategory))
            end if
            NShopsTotal2 = act2.TotalNShops
            if nSubCategory <> "" then
                NShopsTotal2 = act2.arrTOTALNShops(CInt(nSubCategory))
            end if
            
            'autoFill --> para que se copien los valores a las Subcategorías
            jsAutofill = ""
            if nSubCategory = "" then
                jsAutofill = "autoFill(this);"
            end if
            
            'Half1
            strOut = strOut & "<td align=center class=Clk>"
            if Request.Form("ViewNShopsActivity")<>"" then
                strOut = strOut & "<input type=text readonly style=""width:60px;"" value=""" & NShopsActivity1 & """ class=""textfieldreadonly"" />"
            end if
            if Request.Form("ViewNShops")<>"" then
                strOut = strOut & "<input onchange=""modif();" & jsAutofill & """ tabindex=""" & iter & Right("0" & brandIteration, 2) & """ type=text name=""NS_" & Right("000" & client.IDClient, 4) & "_" & Right("000" & brand.IDBrand, 4) & "_" & iYear & "_" & Right("0" & iMonth, 2) & "_1_" & nSubCategory & """ class=""textfieldNShops"" style=""width:60px;"" value=""" & NShops1 & """  />"
            end if
            if Request.Form("ViewPercentComplaint")<>"" then
                strOut = strOut & "<input readonly tabindex=""" & iter+2 & Right("0" & brandIteration, 2) & """ type=text class=""textfieldPercentCompl"" style=""width:60px;"" value=""" & NShopsTotal1 & """ />"
            end if
            strOut = strOut & "</td>"
            
            'Half2
            strOut = strOut & "<td align=center class=Clk>"
            if Request.Form("ViewNShopsActivity")<>"" then
                strOut = strOut & "<input type=text readonly style=""width:60px;"" value=""" & NShopsActivity2 & """ class=""textfieldreadonly"" />"
            end if
            if Request.Form("ViewNShops")<>"" then
                strOut = strOut & "<input onchange=""modif();" & jsAutofill & """ tabindex=""" & iter+1 & Right("0" & brandIteration, 2) & """ type=text name=""NS_" & Right("000" & client.IDClient, 4) & "_" & Right("000" & brand.IDBrand, 4) & "_" & iYear & "_" & Right("0" & iMonth, 2) & "_2_" & nSubCategory & """ class=""textfieldNShops"" style=""width:60px;"" value=""" & NShops2 & """ />"
            end if
            if Request.Form("ViewPercentComplaint")<>"" then
                strOut = strOut & "<input readonly tabindex=""" & iter+3 & Right("0" & brandIteration, 2) & """ type=text class=""textfieldPercentCompl"" style=""width:60px;"" value=""" & NShopsTotal2 & """/>"
            end if
            strOut = strOut & "</td>"
            
            
            iter = iter + 4
        next
        
        strOut = strOut & "</tr>"
        
        PintarRealData = strOut
    End Function
    
    
</script>