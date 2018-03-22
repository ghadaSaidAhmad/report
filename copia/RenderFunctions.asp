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
        
        strOut = strOut & "<table width=""100%"" cellpadding=0 cellspacing=0 style=""border-right:1 solid gray;border-bottom:1 solid gray;"">"
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
                    q1Style = q1Style & "border:2 solid silver;background-color:gray;color:white;font-weight:bold;"
                else
                    q1Style = q1Style & "border:2 solid silver;background-color:white;color:black;font-weight:normal;"
                end if
                if act2.ID <> -1 then
                    q2Style = q2Style & "border:2 solid silver;background-color:gray;color:white;font-weight:bold;"
                else
                    q2Style = q2Style & "border:2 solid silver;background-color:white;color:black;font-weight:normal;"
                end if
                
                if CInt(currYear) = iYear AND CInt(currMonth) = iMonth AND CInt(currHalf) = 1 then
                    q1Style = q1Style & "border:2 solid red;"
                elseif CInt(currYear) = iYear AND CInt(currMonth) = iMonth AND CInt(currHalf) = 2 then
                    q2Style = q2Style & "border:2 solid red;"
                end if
                
                strOutMeses = strOutMeses & "<td colspan=2 class=gridmonthtitlesmall>" & Left(locMonthName(iMonth, Idioma), 3) & sYear & "</td>"
                
                strOutQuincenas = strOutQuincenas & "<td class=gridhalf1title style=""border-top:1 solid gray;cursor:pointer;" & q1Style & """ onClick=""navigateTo(" & iYear & ", " & iMonth & ", 1);"">1</td>"
                strOutQuincenas = strOutQuincenas & "<td class=gridhalf1title style=""border-top:1 solid gray;cursor:pointer;" & q2Style & """ onClick=""navigateTo(" & iYear & ", " & iMonth & ", 2);"">2</td>"
                
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
        
        strOut = strOut & "<table width=""100%"" cellpadding=0 cellspacing=0 style=""border-right:1 solid gray;border-bottom:1 solid gray;"">"
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
                    q1Style = q1Style & "border:2 solid silver;background-color:gray;color:white;font-weight:bold;"
                else
                    q1Style = q1Style & "border:2 solid silver;background-color:white;color:black;font-weight:normal;"
                end if
                if gthm2.ID <> -1 then
                    q2Style = q2Style & "border:2 solid silver;background-color:gray;color:white;font-weight:bold;"
                else
                    q2Style = q2Style & "border:2 solid silver;background-color:white;color:black;font-weight:normal;"
                end if
                
                if CInt(currYear) = iYear AND CInt(currMonth) = iMonth AND CInt(currHalf) = 1 then
                    q1Style = q1Style & "border:2 solid red;"
                elseif CInt(currYear) = iYear AND CInt(currMonth) = iMonth AND CInt(currHalf) = 2 then
                    q2Style = q2Style & "border:2 solid red;"
                end if
                
                strOutMeses = strOutMeses & "<td colspan=2 class=gridmonthtitlesmall>" & Left(locMonthName(iMonth, Idioma), 3) & sYear & "</td>"
                
                strOutQuincenas = strOutQuincenas & "<td class=gridhalf1title style=""border-top:1 solid gray;cursor:pointer;" & q1Style & """ onClick=""navigateTo(" & iYear & ", " & iMonth & ", 1);"">1</td>"
                strOutQuincenas = strOutQuincenas & "<td class=gridhalf1title style=""border-top:1 solid gray;cursor:pointer;" & q2Style & """ onClick=""navigateTo(" & iYear & ", " & iMonth & ", 2);"">2</td>"
                
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
    Function PintarActivityTipo0(StartYear, StartMonth, ViewMonths, IDClient, IDBrand, RowSpan, RowTitleBgColor)
        dim arrActivity
        dim IDType
        IDType = 0
        dim strOut
        
        dim actType
        set actType = getActivityType(IDType, Idioma)
        
        dim i, iMonth, iYear
        
        'strOut = strOut & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & " rowspan=" & RowSpan & ">" & StartYear & "</td>"
        strOut = strOut & "<td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & actType.Name & "</td>"
        
        ' Para cada mes, busca info de cada quincena
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            arrActivity = getActivities00(iYear, iMonth, IDClient, IDBrand, IDType)
            
            dim sImageTag1, sImageTag2
            if arrActivity(0).ThemeImageFileName <> "" then
                sImageTag1 = "<img alt=""" & arrActivity(0).GridText & """ src=""images/Themes/" & arrActivity(0).ThemeImageFileName & """ width=""" & Application("ThemeImageWidth") & """ />"
            else
                sImageTag1 = ""
            end if
            if arrActivity(1).ThemeImageFileName <> "" then
                sImageTag2 = "<img alt=""" & arrActivity(1).GridText & """ src=""images/Themes/" & arrActivity(1).ThemeImageFileName & """ width=""" & Application("ThemeImageWidth") & """ />"
            else
                sImageTag2 = ""
            end if
            
            ' Primera fila
            strOut = strOut & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "30px;", "center", arrActivity(0).FGColor, arrActivity(0).BGColor, arrActivity(0).GridText, "editAct('SOAActivity00.asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1, '" & IDType & "')", sImageTag1)
            strOut = strOut & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", arrActivity(1).FGColor, arrActivity(1).BGColor, arrActivity(1).GridText, "editAct('SOAActivity00.asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2, '" & IDType & "')", sImageTag2)
            
        next
        
        strOut = strOut & "</tr>"
        
        PintarActivityTipo0 = strOut
    End Function
    
    
    
    ' Pinta las actividades de un tipo en el calendario
    Function PintarActivity(StartYear, StartMonth, ViewMonths, IDClient, IDBrand, CurrentLast)
        dim arrActivity
        dim strOut, strOut1, strOut2, strOut3, strOut4, strOut5, strOut6
        dim BGColorOferta1, BGColorOferta2, FGColorOferta1, FGColorOferta2
        dim BGColorFolleto1, FGColorFolleto1, BGColorFolleto2, FGColorFolleto2
        dim BGColorCabecera1, FGColorCabecera1, BGColorCabecera2, FGColorCabecera2
        dim BGColorCentros1, FGColorCentros1, BGColorCentros2, FGColorCentros2
        dim BGColorAdicional1, FGColorAdicional1, BGColorAdicional2, FGColorAdicional2
        
        dim i, iMonth, iYear
        
        dim RowTitleBgColor
        if CurrentLast = "CURRENT" then
            RowTitleBgColor = Application("ColorCurrentYear")
        else
            RowTitleBgColor = Application("ColorLastYear")
        end if

        strOut1 = strOut1 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & IDM_Oferta & "</td>"
        strOut2 = strOut2 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & IDM_Folleto & "</td>"
        strOut3 = strOut3 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & IDM_Cabecera & "</td>"
        strOut4 = strOut4 & "<tr><td class=gridtypetitleNoBold bgcolor=" & RowTitleBgColor & ">" & IDM_NTiendas & "</td>"
        strOut5 = strOut5 & "<tr><td class=gridtypetitleNoBold bgcolor=" & RowTitleBgColor & ">" & IDM_PercentComplaint & "</td>"
        strOut6 = strOut6 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & IDM_Adicional & "</td>"
        
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
            
            ' COLORES DE OFERTA ***************************************
            BGColorOferta1 = ""
            BGColorOferta2 = ""
            if arrActivity(0).Oferta<>"" then
                if CurrentLast = "CURRENT" then
                    BGColorOferta1 = Application("ColorBGOferta")
                else
                    BGColorOferta1 = Application("ColorBGOfertaLY")
                end if
            end if
            if arrActivity(1).Oferta<>"" then
                if CurrentLast = "CURRENT" then
                    BGColorOferta2 = Application("ColorBGOferta")
                else
                    BGColorOferta2 = Application("ColorBGOfertaLY")
                end if
            end if



            ' COLORES DE FOLLETO ***************************************
            BGColorFolleto1 = ""
            BGColorFolleto2 = ""
            if arrActivity(0).Folleto<>"" then
                if CurrentLast = "CURRENT" then
                    BGColorFolleto1 = Application("ColorBGFolleto")
                else
                    BGColorFolleto1 = Application("ColorBGFolletoLY")
                end if
            end if
            if arrActivity(1).Folleto<>"" then
                if CurrentLast = "CURRENT" then
                    BGColorFolleto2 = Application("ColorBGFolleto")
                else
                    BGColorFolleto2 = Application("ColorBGFolletoLY")
                end if
            end if


            ' COLORES DE CABECERA ***************************************
            BGColorCabecera1 = ""
            BGColorCabecera2 = ""
            if arrActivity(0).Cabecera<>"" then
                BGColorCabecera1 = arrActivity(0).RatioBackground
            end if
            if arrActivity(1).Cabecera<>"" then
                BGColorCabecera2 = arrActivity(1).RatioBackground
            end if

            
            ' COLORES DE NCENTROS ***************************************
            BGColorCentros1 = ""
            BGColorCentros2 = ""
            if arrActivity(0).NShops<>"" then
                if CurrentLast = "CURRENT" then
                    BGColorCentros1 = Application("ColorBGNShops")
                else
                    BGColorCentros1 = Application("ColorBGNShopsLY")
                end if
            end if
            if arrActivity(1).NShops<>"" then
                if CurrentLast = "CURRENT" then
                    BGColorCentros2 = Application("ColorBGNShops")
                else
                    BGColorCentros2 = Application("ColorBGNShopsLY")
                end if
            end if
            
            ' COLORES DE ADICIONAL ***************************************
            BGColorAdicional1 = ""
            BGColorAdicional2 = ""
            if arrActivity(0).Adicional<>"" then
                if CurrentLast = "CURRENT" then
                    BGColorAdicional1 = Application("ColorBGAdicional")
                else
                    BGColorAdicional1 = Application("ColorBGAdicionalLY")
                end if
            end if
            if arrActivity(1).Adicional<>"" then
                if CurrentLast = "CURRENT" then
                    BGColorAdicional2 = Application("ColorBGAdicional")
                else
                    BGColorAdicional2 = Application("ColorBGAdicionalLY")
                end if
            end if
            
            
            FGColorOferta1 = Application("ColorFGStatus" & arrActivity(0).IDStatus)
            FGColorOferta2 = Application("ColorFGStatus" & arrActivity(1).IDStatus)
            FGColorFolleto1 = Application("ColorFGStatus" & arrActivity(0).IDStatus)
            FGColorFolleto2 = Application("ColorFGStatus" & arrActivity(1).IDStatus)
            FGColorCabecera1 = Application("ColorFGStatus" & arrActivity(0).IDStatus)
            FGColorCabecera2 = Application("ColorFGStatus" & arrActivity(1).IDStatus)


            ' *******************************
            ' Fila Oferta
            strOut1 = strOut1 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "30px;", "center", FGColorOferta1, BGColorOferta1, arrActivity(0).Oferta, "editAct('SOAActivity.asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1)", "")
            strOut1 = strOut1 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", FGColorOferta2, BGColorOferta2, arrActivity(1).Oferta, "editAct('SOAActivity.asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2)", "")

            ' *******************************
            ' Fila Folleto
            strOut2 = strOut2 & RenderCell(true, "", "cell", Application("ReportHalfWidth") & "px", "30px;", "center", FGColorFolleto1, BGColorFolleto1, arrActivity(0).Folleto, "editAct('SOAActivity.asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1)", "")
            strOut2 = strOut2 & RenderCell(true, "", "cell", Application("ReportHalfWidth") & "px", "", "center", FGColorFolleto2, BGColorFolleto2, arrActivity(1).Folleto, "editAct('SOAActivity.asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2)", "")

            ' *******************************
            ' Fila Cabecera
            strOut3 = strOut3 & RenderCell(true, "", "cell", Application("ReportHalfWidth") & "px", "30px;", "center", FGColorCabecera1, BGColorCabecera1, arrActivity(0).Cabecera, "editAct('SOAActivity.asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1)", "")
            strOut3 = strOut3 & RenderCell(true, "", "cell", Application("ReportHalfWidth") & "px", "", "center", FGColorCabecera2, BGColorCabecera2, arrActivity(1).Cabecera, "editAct('SOAActivity.asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2)", "")

            ' *******************************
            ' Fila NShops
            strOut4 = strOut4 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "30px;", "center", FGColorCentros1, BGColorCentros1, arrActivity(0).NShops, "editAct('SOAActivity.asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1)", "")
            strOut4 = strOut4 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", FGColorCentros2, BGColorCentros2, arrActivity(1).NShops, "editAct('SOAActivity.asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2)", "")

            ' *******************************
            ' Fila Percent Complaint
            strOut5 = strOut5 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "30px;", "center", FGColorCabecera1, BGColorCabecera1, arrActivity(0).RD_PercentComplaint, "editAct('SOAActivity.asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1)", "")
            strOut5 = strOut5 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", FGColorCabecera2, BGColorCabecera2, arrActivity(1).RD_PercentComplaint, "editAct('SOAActivity.asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2)", "")

            ' *******************************
            ' Fila Adicional
            strOut6 = strOut6 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "30px;", "center", FGColorAdicional1, BGColorAdicional1, arrActivity(0).Adicional, "editAct('SOAActivity.asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1)", "")
            strOut6 = strOut6 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", FGColorAdicional2, BGColorAdicional2, arrActivity(1).Adicional, "editAct('SOAActivity.asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2)", "")

        next
        
        strOut1 = strOut1 & "</tr>"
        strOut2 = strOut2 & "</tr>"
        strOut3 = strOut3 & "</tr>"
        strOut4 = strOut4 & "</tr>"
        strOut5 = strOut5 & "</tr>"
        strOut6 = strOut6 & "</tr>"

        strOut = strOut1 & strOut2 & strOut3 & strOut4 
        '''''''' strOut = strOut & strOut5
        strOut = strOut & strOut6

        
        PintarActivity = strOut
    End Function
    
    
    ' Pinta las actividades de un tipo en el calendario
    Function PintarActivityTipo2(StartYear, StartMonth, ViewMonths, IDClient, IDBrand, RowTitleBgColor)
        dim arrActivity
        dim IDType
        IDType = 2
        dim strOut, strOut1, strOut2, strOut3
        strOut1 = ""
        strOut2 = ""
        strOut3 = ""
        
        dim actType
        set actType = getActivityType(IDType, Idioma)
        
        dim i, iMonth, iYear
        
        strOut1 = strOut1 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & actType.Name & "</td>"
        strOut2 = strOut2 & "<tr><td class=gridtypetitleNoBold bgcolor=" & RowTitleBgColor & ">&nbsp;&nbsp;&nbsp;" & IDM_NTiendas & "</td>"
        strOut3 = strOut3 & "<tr><td class=gridtypetitleNoBold bgcolor=" & RowTitleBgColor & ">&nbsp;&nbsp;&nbsp;" & IDM_PercentComplaint & "</td>"
        
        ' Para cada mes, busca info de cada quincena
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            arrActivity = getActivities02(iYear, iMonth, IDClient, IDBrand, IDType)
            
            ' Primera fila
            strOut1 = strOut1 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "30px", "center", arrActivity(0).TextFGColor, arrActivity(0).TextBGColor, arrActivity(0).GridText, "editAct('SOAActivity" & Right("0" & IDType, 2) & ".asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1, '" & IDType & "')", "")
            strOut1 = strOut1 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", arrActivity(1).TextFGColor, arrActivity(1).TextBGColor, arrActivity(1).GridText, "editAct('SOAActivity" & Right("0" & IDType, 2) & ".asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2, '" & IDType & "')", "")
            
            ' Segunda fila (Otro campo de la actividad)
            strOut2 = strOut2 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", "", "", arrActivity(0).NShops, "editAct('SOAActivity" & Right("0" & IDType, 2) & ".asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1, '" & IDType & "')", "")
            strOut2 = strOut2 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", "", "", arrActivity(1).NShops, "editAct('SOAActivity" & Right("0" & IDType, 2) & ".asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2, '" & IDType & "')", "")
            
            ' Tercera fila (Otro campo de la actividad)
            strOut3 = strOut3 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", "", "", arrActivity(0).PercentComplaint, "editAct('SOAActivity" & Right("0" & IDType, 2) & ".asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1, '" & IDType & "')", "")
            strOut3 = strOut3 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", "", "", arrActivity(1).PercentComplaint, "editAct('SOAActivity" & Right("0" & IDType, 2) & ".asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2, '" & IDType & "')", "")

            ' Cuarta fila (Otro campo de la actividad)
            ' strOut4 = strOut4 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", arrActivity(0).StatusFGColor, arrActivity(0).StatusBGColor, arrActivity(0).Status, "editAct('SOAActivity" & Right("0" & IDType, 2) & ".asp', '" & arrActivity(0).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 1, '" & IDType & "')", "")
            ' strOut4 = strOut4 & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", arrActivity(1).StatusFGColor, arrActivity(1).StatusBGColor, arrActivity(1).Status, "editAct('SOAActivity" & Right("0" & IDType, 2) & ".asp', '" & arrActivity(1).ID & "', '" & IDClient & "', '" & IDBrand & "', '" & iYear & "', '" & iMonth & "', 2, '" & IDType & "')", "")

        next
        
        strOut1 = strOut1 & "</tr>"
        strOut2 = strOut2 & "</tr>"
        strOut3 = strOut3 & "</tr>"

        strOut = strOut1 & strOut2 & strOut3
        
        PintarActivityTipo2 = strOut
    End Function
    
    
    Function PintarFilaBlanco(StartYear, StartMonth, ViewMonths)
        dim strOut
        strOut = ""
        
        dim i, iMonth, iYear


        strOut = strOut & "<tr height=1px><td style=""border-top:1 solid black;border-bottom:1 solid black;""><font class=font8>&nbsp;</font></td><td style=""border-top:1 solid black;border-bottom:1 solid black;""><font class=font8>&nbsp;</font></td><td style=""border-top:1 solid black;border-bottom:1 solid black;""><font class=font8>&nbsp;</font></td>"

        ' Para cada mes, busca info de cada quincena
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if

            strOut = strOut & "<td style=""border-top:1 solid black;border-bottom:1 solid black;""><font class=font8>&nbsp;</font></td>"
            strOut = strOut & "<td style=""border-top:1 solid black;border-bottom:1 solid black;""><font class=font8>&nbsp;</font></td>"
            
            
        next
        
        strOut = strOut & "</tr>"
        
        PintarFilaBlanco = strOut
    End Function
    
    
    Function PintarReportClientBrand(StartYear, StartMonth, ViewMonths, client, brand, YearRowSpan, ReportNumRowsPerYear, RowTitle)
        dim strOut
        strOut = ""
        
        strOut = strOut & PintarGrupo0(YearRowSpan, RowTitle, client.IDClient & "_" & brand.IDBrand)
        strOut = strOut & PintarGrupo05(ReportNumRowsPerYear, StartYear, Application("ColorCurrentYear"))
        if Request.Form("FILTER_SHOWGENERALTHEME")<>"" then
            strOut = strOut & PintarGeneralTheme(client.IDClient, StartYear, StartMonth, ViewMonths, Application("ColorCurrentYear"), Application("ColorBGGeneralThemeCurrentYear"), "Black")
        end if
        
        if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "1") then
            strOut = strOut & PintarActivity(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, "CURRENT")
        end if

        if FALSE then
            if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "0") then
                strOut = strOut & PintarActivityTipo0(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, ReportNumRowsPerYear, Application("ColorCurrentYear"))
            end if
            if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "1") then
                strOut = strOut & PintarActivity(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, 1, Application("ColorCurrentYear"))
            end if
            if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "2") then
                strOut = strOut & PintarActivityTipo2(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorCurrentYear"))
            end if
            if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "3") then
                strOut = strOut & PintarActivity(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, 3, Application("ColorCurrentYear"))
            end if
        end if

        strOut = strOut & PintarSOARealData(StartYear, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorCurrentYear"), "")
        
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
            strOut = strOut & PintarGrupo05(ReportNumRowsPerYear, StartYear-1, Application("ColorLastYear"))
            if Request.Form("FILTER_SHOWGENERALTHEME")<>"" then
                strOut = strOut & PintarGeneralTheme(client.IDClient, StartYear-1, StartMonth, ViewMonths, Application("ColorLastYear"), Application("ColorBGGeneralThemeLastYear"), "Black")
            end if
            
            if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "1") then
                strOut = strOut & PintarActivity(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, "LAST")
            end if

            if FALSE then
                if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "0") then
                    strOut = strOut & PintarActivityTipo0(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, ReportNumRowsPerYear, Application("ColorLastYear"))
                end if
                if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "1") then
                    strOut = strOut & PintarActivity(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, 1, Application("ColorLastYear"))
                end if
                if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "2") then
                    strOut = strOut & PintarActivityTipo2(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorLastYear"))
                end if
                if IsInArray(Request.Form("FILTER_MULTIACTIVITYTYPE"), "3") then
                    strOut = strOut & PintarActivity(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, 3, Application("ColorLastYear"))
                end if
            end if

            strOut = strOut & PintarSOARealData(StartYear-1, StartMonth, ViewMonths, client.IDClient, brand.IDBrand, Application("ColorLastYear"), "")
            
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
    
    
    
    Function PintarSOARealData(StartYear, StartMonth, ViewMonths, IDClient, IDBrand, RowTitleBgColor, CellBgColor)
        dim strOut, strOut1, strOut2
        strOut = ""
        strOut1 = ""
        strOut2 = ""
        
        dim arrRealData, NShops, PercentComplaint, i, iMonth, iYear
        
        strOut1 = strOut1 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & IDM_NTiendasReal & "</td>"
        strOut2 = strOut2 & "<tr><td class=gridtypetitle bgcolor=" & RowTitleBgColor & ">" & IDM_PercentComplaint & "</td>"

        ' Para cada mes, busca info
        iMonth = StartMonth
        iYear = StartYear
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            arrRealData = getRealDatas(IDClient, IDBrand, iYear, iMonth)
            
            ' NShops
            strOut1 = strOut1 & RenderCell(false, "", "cell", "", "", "center", "", CellBgColor, arrRealData(0).NShops, "", "")
            strOut1 = strOut1 & RenderCell(false, "", "cell", "", "", "center", "", CellBgColor, arrRealData(1).NShops, "", "")
            
            ' PercentComplaint
            strOut2 = strOut2 & RenderCell(false, "", "cell", "", "", "center", "", CellBgColor, arrRealData(0).PercentComplaint, "", "")
            strOut2 = strOut2 & RenderCell(false, "", "cell", "", "", "center", "", CellBgColor, arrRealData(1).PercentComplaint, "", "")
            
            
        next
        
        strOut1 = strOut1 & "</tr>"
        strOut2 = strOut2 & "</tr>"
        
        if Request.Form("FILTER_SHOWREALDATA_NSHOPS")<>"" then
            strOut = strOut & strOut1
        end if
        if Request.Form("FILTER_SHOWREALDATA_PERCENTCOMPLAINT")<>"" then
            strOut = strOut & strOut2
        end if
        
        PintarSOARealData = strOut
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
    Function PintarGeneralTheme(IDClient, StartYear, StartMonth, ViewMonths, RowTitleBgColor, BGColor, FGColor)
        dim arrGeneralThemes
        dim strOut
        
        dim i, iMonth, iYear
        
        strOut = strOut & "<tr><td class=gridtypetitle bgcolor=""" & RowTitleBgColor & """>" & IDM_GeneralTheme & "</td>"
        
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
                sImageTag1 = "<img alt=""" & arrGeneralThemes(0).GridText & """ src=""images/Themes/" & arrGeneralThemes(0).ThemeImageFileName & """ width=""" & Application("ThemeImageWidth") & """ />"
            else
                sImageTag1 = ""
            end if
            if arrGeneralThemes(1).ThemeImageFileName <> "" then
                sImageTag2 = "<img alt=""" & arrGeneralThemes(1).GridText & """ src=""images/Themes/" & arrGeneralThemes(1).ThemeImageFileName & """ width=""" & Application("ThemeImageWidth") & """ />"
            else
                sImageTag2 = ""
            end if
            
            dim sBGColor1, sBGColor2, sFGColor1, sFGColor2
            sBGColor1 = ""
            sFGColor1 = ""
            if Trim(arrGeneralThemes(0).GridText)<>"" then
                sBGColor1 = BGColor
                sFGColor1 = FGColor
            end if
            sBGColor2 = ""
            sFGColor2 = ""
            if Trim(arrGeneralThemes(1).GridText)<>"" then
                sBGColor2 = BGColor
                sFGColor2 = FGColor
            end if

            strOut = strOut & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "30px;", "center", sFGColor1, sBGColor1, arrGeneralThemes(0).GridText, "editGenThem('SOAGeneralTheme.asp', '" & arrGeneralThemes(0).ID & "', '" & IDClient & "', '" & iYear & "', '" & iMonth & "', 1)", sImageTag1)
            strOut = strOut & RenderCell(false, "", "cell", Application("ReportHalfWidth") & "px", "", "center", sFGColor2, sBGColor2, arrGeneralThemes(1).GridText, "editGenThem('SOAGeneralTheme.asp', '" & arrGeneralThemes(1).ID & "', '" & IDClient & "', '" & iYear & "', '" & iMonth & "', 2)", sImageTag2)

        next
        
        strOut = strOut & "</tr>"

        
        PintarGeneralTheme = strOut
    End Function    
    
    
    ' Pinta una casilla del report
    Function RenderCell(Bold, Colspan, ClassName, Width, Height, Align, ForeColor, BackColor, Text, jsOnClick, ImageTag)
        dim strOut
        strOut = ""
        
        strOut = strOut & "<TD "
        if Colspan<>"" then
            strOut = strOut & " colspan=""" & Colspan & """ "
        end if
        if ClassName<>"" then
            strOut = strOut & " class=""" & ClassName & """ "
        end if
        if Align<>"" then
            strOut = strOut & " align=""" & Align & """ "
        end if
        strOut = strOut & " style="""
        if ForeColor<>"" then
            strOut = strOut & "color:" & ForeColor & ";"
        end if
        if Width<>"" then
            strOut = strOut & "width:" & Width & ";"
        end if
        if Height<>"" then
            strOut = strOut & "height:" & Height & ";"
        end if
        if Bold then
            strOut = strOut & "font-weight:bold;"
        end if
        if BackColor<>"" then
            strOut = strOut & "background-color:" & BackColor & ";"
        end if
        strOut = strOut & """  "
        if jsOnClick<>"" then
            strOut = strOut & " onclick=""" & jsOnClick & """ "
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
        
        strOut = strOut & "<tr><td class=gridtypetitle>&nbsp;</td><td style=""border-top:1 solid gray;"" >&nbsp;</td><td style=""border-top:1 solid gray;"">&nbsp;</td>"

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
                strOut = strOut & "<input type=text style=""width:60px;border:0px;text-align:center;"" readonly value=""" & IDM_PercentComplaintShort & """ />"
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
                strOut = strOut & "<input type=text style=""width:60px;border:0px;text-align:center;"" readonly value=""" & IDM_PercentComplaintShort & """ />"
            end if
            strOut = strOut & "</td>"
            
        next
        
        strOut = strOut & "</tr>"

        PintarColumnasRealData = strOut
    End Function
    
    
    Function PintarRealData(client, brand, StartYear, StartMonth, ViewMonths, brandIteration)
        dim strOut
        dim act1, act2
        dim SQL, rst
        dim i, iMonth, iYear, iter
        dim NShops1, PercentCompl1, NShops2, PercentCompl2
        dim NShopsActivity1, NShopsActivity2

        strOut = strOut & "<tr><td class=gridtypetitle>" & bra.Name & "</td><td style=""border-top:1 solid gray;"">&nbsp;</td><td style=""border-top:1 solid gray;"">&nbsp;</td>"

        iMonth = StartMonth
        iYear = StartYear
        iter = 10
        for i = StartMonth to (StartMonth + ViewMonths - 1)
            iMonth = i
            if iMonth > 12 then
                iMonth = iMonth - 12
                iYear = StartYear + 1
            end if
            
            NShops1 = ""
            PercentCompl1 = ""
            SQL = "SELECT NShops, PercentComplaint " & _
            " FROM RealData " & _
            " WHERE WYear = " & iYear & " AND IDClient = " & client.IDClient & " AND IDBrand = " & brand.IDBrand & _
                " AND WMonth = " & iMonth & " AND WHalf = 1 "
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NShops1 = rst("NShops")
                PercentCompl1 = rst("PercentComplaint")
            else
                SQL = "INSERT INTO RealData (WYear, IDClient, IDBrand, WMonth, WHalf, NShops, PercentComplaint) " & _
                " VALUES (" & iYear & ", " & client.IDClient & ", " & brand.IDBrand & ", " & iMonth & ", 1, NULL, NULL)"
                ObjConnectionSQL.Execute(SQL)
            end if
            set rst = nothing

            NShops2 = ""
            PercentCompl2 = ""
            SQL = "SELECT NShops, PercentComplaint " & _
            " FROM RealData " & _
            " WHERE WYear = " & iYear & " AND IDClient = " & client.IDClient & " AND IDBrand = " & brand.IDBrand & _
                " AND WMonth = " & iMonth & " AND WHalf = 2 "
            set rst = ObjConnectionSQL.Execute(SQL)
            if NOT rst.EOF then
                NShops2 = rst("NShops")
                PercentCompl2 = rst("PercentComplaint")
            else
                SQL = "INSERT INTO RealData (WYear, IDClient, IDBrand, WMonth, WHalf, NShops, PercentComplaint) " & _
                " VALUES (" & iYear & ", " & client.IDClient & ", " & brand.IDBrand & ", " & iMonth & ", 2, NULL, NULL)"
                ObjConnectionSQL.Execute(SQL)
            end if
            set rst = nothing
            
            
            NShopsActivity1 = ""
            set act1 = getActivityFromDate(client.IDClient, brand.IDBrand, iYear, iMonth, 1)
            NShopsActivity2 = ""
            set act2 = getActivityFromDate(client.IDClient, brand.IDBrand, iYear, iMonth, 2)
            
            
            'Half1
            strOut = strOut & "<td align=center class=cell>"
            if Request.Form("ViewNShopsActivity")<>"" then
                strOut = strOut & "<input type=text readonly style=""width:60px;"" value=""" & act1.NShops & """ class=""textfieldreadonly"" />"
            end if
            if Request.Form("ViewNShops")<>"" then
                strOut = strOut & "<input onchange=""modif();"" tabindex=""" & iter & Right("0" & brandIteration, 2) & """ type=text name=""NS_" & Right("000" & client.IDClient, 4) & "_" & Right("000" & brand.IDBrand, 4) & "_" & iYear & "_" & Right("0" & iMonth, 2) & "_1"" class=""textfieldNShops"" style=""width:60px;"" value=""" & NShops1 & """ />"
            end if
            if Request.Form("ViewPercentComplaint")<>"" then
                strOut = strOut & "<input onchange=""modif();"" tabindex=""" & iter+2 & Right("0" & brandIteration, 2) & """ type=text name=""PC_" & Right("000" & client.IDClient, 4) & "_" & Right("000" & brand.IDBrand, 4) & "_" & iYear & "_" & Right("0" & iMonth, 2) & "_1"" class=""textfieldPercentCompl"" style=""width:60px;"" value=""" & PercentCompl1 & """ />"
            end if
            strOut = strOut & "</td>"
            
            'Half2
            strOut = strOut & "<td align=center class=cell>"
            if Request.Form("ViewNShopsActivity")<>"" then
                strOut = strOut & "<input type=text readonly style=""width:60px;"" value=""" & act2.NShops & """ class=""textfieldreadonly"" />"
            end if
            if Request.Form("ViewNShops")<>"" then
                strOut = strOut & "<input onchange=""modif();"" tabindex=""" & iter+1 & Right("0" & brandIteration, 2) & """ type=text name=""NS_" & Right("000" & client.IDClient, 4) & "_" & Right("000" & brand.IDBrand, 4) & "_" & iYear & "_" & Right("0" & iMonth, 2) & "_2"" class=""textfieldNShops"" style=""width:60px;"" value=""" & NShops2 & """ />"
            end if
            if Request.Form("ViewPercentComplaint")<>"" then
                strOut = strOut & "<input onchange=""modif();"" tabindex=""" & iter+3 & Right("0" & brandIteration, 2) & """ type=text name=""PC_" & Right("000" & client.IDClient, 4) & "_" & Right("000" & brand.IDBrand, 4) & "_" & iYear & "_" & Right("0" & iMonth, 2) & "_2"" class=""textfieldPercentCompl"" style=""width:60px;"" value=""" & PercentCompl2 & """/>"
            end if
            strOut = strOut & "</td>"
            
            
            iter = iter + 4
        next
        
        strOut = strOut & "</tr>"
        
        PintarRealData = strOut
    End Function
    
    
</script>