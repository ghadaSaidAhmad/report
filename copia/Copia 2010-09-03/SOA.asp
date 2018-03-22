<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
	Response.Buffer=true
%>

<%if request("XL")<>"" then%>
	<%Response.ContentType = "application/vnd.ms-excel"%>
<%else %>
<!-- #include file = "include/EventFunctions1.asp" -->
<%end if%>

<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->

<!-- #include file = "RenderFunctions.asp" -->
<!-- #include file = "ClassInclude.asp" -->


<!-- #include file = "ClassSuplantar.asp" -->
<%

showMenu = TRUE  ' Muestra el menú de la aplicación


dim rst, rst2
set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")

rst.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient

RecoverSQLConnection()

RecoverSession(true) 


' Si es el primer acceso y no es administrador, no hace falta el menú
if Request.Form("FILTER_REPORTTYPE") = "" AND NOT IsAdmin() then
    showMenu = FALSE
end if


Sub ApplyFilter_click()
    
    
End Sub
%>

<!-- #include file = "include/Idioma.asp" -->

<% 
Select Case EventObject
	case "ApplyFilter" call ApplyFilter_click()
	
End Select
%>



<HTML>
<HEAD>
    <TITLE>Activity</TITLE>
    <%if Request("XL")="" then %>
        <LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
        <LINK REL=StyleSheet HREF="ClassCellStyles.asp" TYPE="text/css">
        <script>
            function editAct(id, idclient, idbrand, wyear, wmonth, whalf)
            {
                window.open("SOAActivity.asp?ID=" + id + "&IDClient=" + idclient + "&IDBrand=" + idbrand + "&WYear=" + wyear + "&WMonth=" + wmonth + "&WHalf=" + whalf + '&FILTER_YEAR=<%=Request.Form("FILTER_YEAR") %>&FILTER_STARTMONTH=<%=Request.Form("FILTER_STARTMONTH") %>&FILTER_VIEWMONTHS=<%=Request.Form("FILTER_VIEWMONTHS") %>&FILTER_MULTIBRAND=<%=Request.Form("FILTER_MULTIBRAND") %>&FILTER_MULTICLIENT=<%=Request.Form("FILTER_MULTICLIENT") %>', 'ACT', 'width=600, height=600, top=50, left=200, scrollbars, status');
            }
            function editThm(id, idclient, idbrand, wyear, wmonth, whalf)
            {
                window.open("SOAThemeCB.asp?ID=" + id + "&IDClient=" + idclient + "&IDBrand=" + idbrand + "&WYear=" + wyear + "&WMonth=" + wmonth + "&WHalf=" + whalf, 'THM', 'width=600, height=350, top=150, left=200, scrollbars');
            }
            function editGenThem(page, id, idclient, wyear, wmonth, whalf)
            {
                window.open(page + "?ID=" + id + "&IDClient=" + idclient + "&WYear=" + wyear + "&WMonth=" + wmonth + "&WHalf=" + whalf + '&FILTER_YEAR=<%=Request.Form("FILTER_YEAR") %>&FILTER_STARTMONTH=<%=Request.Form("FILTER_STARTMONTH") %>&FILTER_VIEWMONTHS=<%=Request.Form("FILTER_VIEWMONTHS") %>', 'THM', 'width=600, height=350, top=150, left=200, scrollbars, status');
            }
            function valorarCalExp(obj)
            {
                obj.style.display = '';
            }
            function valorarCalOf(obj)
            {
                obj.style.display = '';
            }
            function saveCalExp(obj, lbl)
            {   
                if (obj.value == ''){
                    obj.style.display = 'none';
                    
                    obj.value = obj.alt;  // Devuelve el valor que tenía antes
                    return false;
                }
            
                sDat = obj.name + '___' + obj.value;
    			ajaxres = ajaxReq('SaveCalExp', sDat);
    			
    			if (ajaxres == "OK"){
    			    //alert('Guardado');
    			    
    			    obj.alt = obj.value;  // En la propiedad ALT se guarda el valor actual
    			    if (obj.options[obj.selectedIndex].text == ""){
    			        lbl.innerText = "       ";
    			    }else{
    			        lbl.innerText = obj.options[obj.selectedIndex].text;
    			    }
    			    
    			}else{
    			    alert('Error ' + ajaxres);
    			    
                    obj.value = obj.alt;  // Devuelve el valor que tenía antes
    			    return false;
    			}
                obj.style.display = 'none';
            }
            function saveCalOf(obj, lbl)
            {   
                if (obj.value == ''){
                    obj.style.display = 'none';
                    
                    obj.value = obj.alt;  // Devuelve el valor que tenía antes
                    return false;
                }
            
                sDat = obj.name + '___' + obj.value;
    			ajaxres = ajaxReq('SaveCalOf', sDat);
    			
    			if (ajaxres == "OK"){
    			    //alert('Guardado');
    			    
    			    obj.alt = obj.value;  // En la propiedad ALT se guarda el valor actual
    			    if (obj.options[obj.selectedIndex].text == ""){
    			        lbl.innerText = "       ";
    			    }else{
    			        lbl.innerText = obj.options[obj.selectedIndex].text;
    			    }
    			    
    			}else{
    			    alert('Error ' + ajaxres);
    			    
                    obj.value = obj.alt;  // Devuelve el valor que tenía antes
    			    return false;
    			}
                obj.style.display = 'none';
            }
        </script>
    <%end if %>
</HEAD>

<BODY class="BODY_MAIN" <%if Request.Form("FILTER_REPORTTYPE") = "" then %>style="background-image:url('images/background.jpg');background-repeat:no-repeat;"<%end if %>>
<%if Request("XL")="" then %>
<!-- #include file = "include/WaitingIcon.asp" -->
<%end if %>



<FORM action="" method="post" name="thisForm">
<%if Request("XL")="" then %>
<!-- #include file = "ClassMenu.asp" -->
<%end if %>

    <%
    dim ReportNumRowsPerYear: ReportNumRowsPerYear = getReportRows()
    
    
    dim ViewMonths
    if Request.Form("FILTER_VIEWMONTHS")<>"" then
        ViewMonths = CInt(Request.Form("FILTER_VIEWMONTHS"))
    else
        ViewMonths = 4
    end if
    dim StartYear
    if Request.Form("FILTER_YEAR")<>"" then
        StartYear = CInt(Request.Form("FILTER_YEAR"))
    else
        StartYear = Year(Date)
    end if
    dim StartMonth
    if Request.Form("FILTER_STARTMONTH")<>"" then
        StartMonth = CInt(Request.Form("FILTER_STARTMONTH"))
    else
        StartMonth = Month(Date)
    end if
    
    dim YearRowSpan
    if Request.Form("FILTER_LASTYEAR")<>"" then
        YearRowSpan = 1 + (ReportNumRowsPerYear * 2)
    else
        YearRowSpan = 1 + ReportNumRowsPerYear
    end if
    


    dim tableWidth
    tableWidth = "100%"
    if ViewMonths < 4 then
        tableWidth = (250 + ViewMonths*2*Application("ReportHalfWidth")) & "px"
    else
        tableWidth = (250 + ViewMonths*2*Application("ReportHalfWidth")) & "px"
    end if
    %>
    
    <%' TABLA TOP%>
    <TABLE width="100%" ID="TBL_TOP" cellpadding=0 cellspacing=0 class=topTable background="images/a3.gif">
        <TR>
            <TD valign=top width=380 style="padding:10px;">
                <font class=fontTitleTop>
                    <%=IDM_MAINTITLE1 %>
                </font>
                <font class=fontTitleTop2>
                    <br />
                    <%=IDM_MAINTITLE2 %>
                </font>
            </TD>
        </TR>
    </TABLE>
    
    <%response.Flush %>
    
    <% if Request.Form("FILTER_REPORTTYPE") = "0" OR Request.Form("FILTER_REPORTTYPE") = "1" then %>
    
        <%if Request("XL")="" then %>
            <!-- #include file = "ClassMenuReportNavigation.asp" -->
        <%end if %>

        <%' TABLA PRINCIPAL QUE CONTIENE TODO EL CALENDARIO%>
        <TABLE width="<%=tableWidth %>" ID="TBL_MAIN" border=0 style="border-left:1 solid gray;border-right:1 solid gray;border-bottom:1 solid gray;" cellpadding=0 cellspacing=0 bordercolorlight=gainsboro bordercolordark=gray>

            <%=PrepararColumnasCalendario(ViewMonths) %>
            
            <%
            
            dim bra, b
            dim cli, c
            dim sName, sTitle
            dim iter
            if Request.Form("FILTER_REPORTTYPE") = "0" then
                set cli = getClient(Request.Form("FILTER_CLIENT"))
                iter = 0
                for each b in split(Request.Form("FILTER_MULTIBRAND"), ",")
                    set bra = getBrand(CInt(b))
                    if bra.ImageFileNameV<>"" then
                        sName = "<img src='images/Brands/" & bra.ImageFileNameV & "' style='width:60px;border:0' />"
                    else
                        sName = bra.Name
                    end if

                    if iter mod Request.Form("FILTER_BREAKE_EACH") = 0 then
                        if cli.ImageFileNameH<>"" then
                            sTitle = "<img height=60 src='images/Clients/" & cli.ImageFileNameH & "' />"
                        else
                            sTitle = "<font class=font15>" & cli.Name & "</font>"
                        end if
                        Response.Write PintarCalendario(StartYear, StartMonth, ViewMonths, sTitle)
                    end if 
                    
                    Response.Write PintarReportClientBrand(StartYear, StartMonth, ViewMonths, cli, bra, YearRowSpan, ReportNumRowsPerYear, sName)
                    
                    if (iter+1) mod Request.Form("FILTER_BREAKE_EACH") = 0 then
                        Response.Write PintarFilaBlanco(StartYear, StartMonth, ViewMonths)
                        Response.Write PintarPrintPageBreak()
                    end if
                    
                    iter = iter + 1
                next
                
            elseif Request.Form("FILTER_REPORTTYPE") = "1" then 
                set bra = getBrand(Request.Form("FILTER_BRAND"))
                iter = 0
                for each c in split(Request.Form("FILTER_MULTICLIENT"), ",")
                    set cli = getClient(CInt(c))
                    if cli.ImageFileNameV<>"" then
                        sName = "<img src='images/Clients/" & cli.ImageFileNameV & "' style='width:60px;border:0' />"
                    else
                        sName = cli.Name
                    end if

                    if iter mod Request.Form("FILTER_BREAKE_EACH") = 0 then
                        if bra.ImageFileNameH<>"" then
                            sTitle = "<img height=60 src='images/Brands/" & bra.ImageFileNameH & "' />"
                        else
                            sTitle = "<font class=font15>" & bra.Name & "</font>"
                        end if
                        Response.Write PintarCalendario(StartYear, StartMonth, ViewMonths, sTitle)
                    end if

                    Response.Write PintarReportClientBrand(StartYear, StartMonth, ViewMonths, cli, bra, YearRowSpan, ReportNumRowsPerYear, sName)

                    if (iter+1) mod Request.Form("FILTER_BREAKE_EACH") = 0 then
                        Response.Write PintarFilaBlanco(StartYear, StartMonth, ViewMonths)
                        Response.Write PintarPrintPageBreak()
                    end if

                    iter = iter + 1
                next
            

            end if %>
            
            
        </TABLE>
    
    
    <%end if %>
    
    
    <br />
    <br />
    <br />
    <br />
    
    <!-- #include file = "include/EventFunctions2.asp" -->
</FORM>

<!-- #include file = "include/pageBottom.asp" -->

</BODY>

</HTML>