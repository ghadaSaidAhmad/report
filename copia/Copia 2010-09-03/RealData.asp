<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
	Response.Buffer=true
%>

<!-- #include file = "include/EventFunctions1.asp" -->

<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->

<!-- #include file = "RenderFunctions.asp" -->
<!-- #include file = "ClassInclude.asp" -->

<%
dim rst, rst2
set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")

rst.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient

RecoverSQLConnection()

RecoverSession(true) 

if not isAdmin() AND not isInputData() then
	msgError "You are not allowed to view this information", true, true
end if


showMenu = TRUE  ' Muestra el menú de la aplicación
dim arrClients, c
dim WYear
WYear = Request.Form("FILTER_YEAR")
if WYear="" then
    WYear = Year(Date())
end if

Sub Save_click()
    Dim r, IDClient, IDBrand, WYear, WMonth, WHalf
    Dim SQL, sNShops, sPercentComp, SetUpdate
    
    For Each r in Request.Form
        
        if Mid(r, 1, 3) = "NS_" OR Mid(r, 1, 3) = "PC_" then
        
            sNShops = Request.Form(r)
            if sNShops = "" then
                sNShops = "NULL"
            else
                on error resume next
                sNShops = CStr(CInt(Request.Form(r)))
                if Err<>0 then
                    sNShops = "NULL"
                end if
                on error goto 0
            end if
            sPercentComp = Request.Form(r)
            if sPercentComp = "" then
                sPercentComp = "NULL"
            else
                sPercentComp = Replace(sPercentComp, ",", ".")
                on error resume next
                sPercentComp = CStr(CDbl(sPercentComp))
                if Err<>0 then
                    sPercentComp = "NULL"
                end if
                on error goto 0
            end if
            
            IDClient = CInt(Mid(r, 4, 4))
            IDBrand = CInt(Mid(r, 9, 4))
            WYear = CInt(Mid(r, 14, 4))
            WMonth = CInt(Mid(r, 19, 2))
            WHalf = CInt(Mid(r, 22))
            
            SetUpdate = ""
            if Mid(r, 1, 3) = "NS_" then
                SetUpdate = " NShops = " & sNShops
            elseif Mid(r, 1, 3) = "PC_" then
                SetUpdate = " PercentComplaint = " & sPercentComp
            end if
            
            if SetUpdate<>"" then
                SQL = "UPDATE RealData SET " & SetUpdate & _
                " WHERE IDClient=" & IDClient & " AND IDBrand=" & IDBrand & " AND WYear=" & WYear & " AND WMonth=" & WMonth & " AND WHalf=" & WHalf
                ObjConnectionSQL.Execute(SQL)
            end if
        end if
    Next
    
    
End Sub

Sub ApplyFilter_click()

End Sub
%>

<!-- #include file = "include/Idioma.asp" -->


<% 
Select Case EventObject
	case "ApplyFilter" call ApplyFilter_click()
	case "Save" call Save_click()
End Select
%>


<HTML>
<HEAD>
    <TITLE>Real Data</TITLE>
    <LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
    <script language="javascript">
        var datamodified = false;
        
        function windowClose()
        {
            if (thisForm.EventObject.value == "")
            {
                if (datamodified)
                {
        			event.returnValue = '<%=IDM_JS_RealData_DatosModificadosSalirSinGuardar %>';
                }
            }
        }
        
        function save()
        {
            if (checkNumbers())
            {
                _fireEvent('Save','','');
            }else{
                alert('<%=IDM_JS_RealData_ErrorEnValor %>');
            }
        }

        function checkNumbers()
        {
	        for (i=0;i<document.thisForm.elements.length;i++) 
	        {
	            //alert(document.thisForm.elements[i].type);
		        if(document.thisForm.elements[i].type == "text")
		        {
		            if (document.thisForm.elements[i].name != "")
		            {
		                tipo = document.thisForm.elements[i].name.substring(0,3);
		                if (tipo == "PC_" || tipo == "NS_")
		                {
		                    if (isNaN(document.thisForm.elements[i].value))
		                    {
		                        document.thisForm.elements[i].focus();
		                        document.thisForm.elements[i].select();
		                        return false;
		                    }
		                }
	    	        }
	    	    }
	    	}
	    	
	    	return true;
        }
        
        function checkChanges()
        {
            if (datamodified){
                if (confirm('<%=IDM_JS_RealData_DatosModificadosSalirSinGuardar %>?')){
                    return true;
                }
            }else{
                return true;
            }
            return false;
        }
        
        function modif()
        {
            datamodified = true;
        }
    </script>
</HEAD>

<BODY class="BODY_MAIN" <%if Request.Form("FILTER_CLIENT") = "" then %>style="background-image:url('images/background.jpg');background-repeat:no-repeat;"<%end if %> onbeforeunload="windowClose();">

<FORM action="" method="post" name="thisForm">
<!-- #include file = "ClassMenuRealData.asp" -->

    <%
    dim tableWidth
    dim StartMonth, ViewMonths
    tableWidth = "100%"
    ViewMonths = CInt(Request.Form("FILTER_VIEWMONTHS"))
    StartMonth = CInt(Request.Form("FILTER_STARTMONTH"))
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
    
    <!-- #include file = "ClassMenuRealDataNavigation.asp" -->

    <%response.Flush %>
    
    <%if Request.Form("FILTER_CLIENT")="" then %>
        <p align=center><font class="font20">Seleccione un cliente</font></p>
    <%else %>
        
        <p align=right style="padding-right:30px;">
            <img onclick="save();" title="<%=IDM_Save %>" style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/save.png" />
        </p>
        
        <TABLE width="<%=tableWidth %>" ID="TBL_MAIN" border=0 style="border-left:1 solid gray;border-right:1 solid gray;border-bottom:1 solid gray;" cellpadding=0 cellspacing=0 bordercolorlight=gainsboro bordercolordark=gray>
            <%=PrepararColumnasCalendario(ViewMonths) %>
            
            <%
                dim cli, arrBrands, bra, iter, sTitle
                set cli = getClient(Request.Form("FILTER_CLIENT"))
                
                if cli.ImageFileNameH<>"" then
                    sTitle = "<img height=60 src='images/Clients/" & cli.ImageFileNameH & "' />"
                else
                    sTitle = "<font class=font15>" & cli.Name & "</font>"
                end if
                
                Response.Write PintarCalendario(WYear, StartMonth, ViewMonths, sTitle)
                Response.Write PintarColumnasRealData(WYear, StartMonth, ViewMonths)
                arrBrands = getBrands("NOMBRE")
                iter = 0
                for each bra in arrBrands
                    Response.Write PintarRealData(cli, bra, WYear, StartMonth, ViewMonths, iter)
                    iter = iter + 1
                next
                
            %>
            
        </TABLE>
        
    <%end if %>
    
    <br />
    <br />
    <br />
    <br />
    
    <input type=hidden name="FILTER_CLIENT" value="<%=Request.Form("FILTER_CLIENT") %>" />
    <input type=hidden name="FILTER_YEAR" value="<%=WYear %>" />
    
    <!-- #include file = "include/EventFunctions2.asp" -->
</FORM>

<!-- #include file = "include/pageBottom.asp" -->

</BODY>

</HTML>


