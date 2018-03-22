<%
Function getClientNavigation()
    dim strOut
    strOut = ""
    
    strOut = strOut & "<DIV onclick=""DIVClient_MouseOut();"" id=""DIV_ClientNavigator"" style=""background-color:white;position:absolute;left:0px;top:0px;width:100%;height:100%;display:none;"">"
    strOut = strOut & "<table width=""100%""><tr><td align=center background='images/grad5.gif'><font class=font15><font color=white>" & IDM_ActivitySelectClient & "</font></font></td><td width=100 align=center><input type=button class=button value=""" & IDM_ActivityNoChange & """></td></tr></table>"
    strOut = strOut & "<br><br>"
    strOut = strOut & "<table align=center width=""80%"" height=""100px;""><tr height=80>"
    
    dim doShow, iCli
    dim arrClients, arrSelClients, c
    arrClients = getClients("NOMBRE")
    arrSelClients = split(FILTER_MULTICLIENT, ",")
    iCli = 1
    for each c in arrClients
        if isInArray(arrSelClients, c.IDClient) then
            doShow = true
            if IDClient<>"" then
                if CInt(IDClient) = c.IDClient then
                    'doShow = false
                end if
            end if
            if doShow then
                iCli = iCli + 1
                if iCli mod 2 = 0 then
                    strOut = strOut & "</tr><tr height=80>"
                end if

                strOut = strOut & "<td valign=middle align=center><a href='' onclick=""navigateToClient('" & c.IDClient & "');return false;"">"
                if c.ImageFileNameH<>"" then
                    strOut = strOut & "<img border=0 src='images/Clients/" & c.ImageFileNameH & "'>"
                else
                    strOut = strOut & "<font class=font15>" & c.Name & "</font>"
                end if
                strOut = strOut & "</a></td>"
            end if
            
        end if
    next

    strOut = strOut & "</tr></table>"
    strOut = strOut & "</DIV>"
    
    
    getClientNavigation = strOut
End Function


Function getBrandNavigation()
    dim strOut
    strOut = ""
    
    strOut = strOut & "<DIV onclick=""DIVBrand_MouseOut();"" id=""DIV_BrandNavigator"" style=""background-color:white;position:absolute;left:0px;top:0px;width:100%;height:100%;display:none;"">"
    strOut = strOut & "<table width=""100%""><tr><td align=center background='images/grad5.gif'><font class=font15><font color=white>" & IDM_ActivitySelectBrand & "</font></font></td><td width=100 align=center><input type=button class=button value=""" & IDM_ActivityNoChange & """></td></tr></table>"
    strOut = strOut & "<br><br>"
    strOut = strOut & "<table align=center width=""80%"" height=""100px;""><tr height=80>"
    
    dim doShow, iBra
    dim arrBrands, arrSelBrands, b
    arrBrands = getBrands("NOMBRE")
    arrSelBrands = split(FILTER_MULTIBRAND, ",")
    iBra = 2
    for each b in arrBrands
        if isInArray(arrSelBrands, b.IDBrand) then
            doShow = true
            if IDBrand<>"" then
                if CInt(IDBrand) = b.IDBrand then
                    'doShow = false
                end if
            end if
            if doShow then
                iBra = iBra + 1
                if iBra mod 3 = 0 then
                    strOut = strOut & "</tr><tr>"
                end if

                strOut = strOut & "<td valign=middle align=center><a href='' onclick=""navigateToBrand('" & b.IDBrand & "');return false;"">"
                if b.ImageFileNameH<>"" then
                    strOut = strOut & "<img border=0 src='images/Brands/" & b.ImageFileNameH & "'>"
                else
                    strOut = strOut & "<font class=font15>" & b.Name & "</font>"
                end if
                strOut = strOut & "</a></td>"
            end if
            
        end if
    next

    strOut = strOut & "</tr></table>"
    strOut = strOut & "</DIV>"
    
    
    getBrandNavigation = strOut
End Function
%>
