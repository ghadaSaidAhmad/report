<script language="javascript">
    function mostrarClients()
    {
        TD_NAV_Clients.style.backgroundColor = '<%=Application("ColorSelectorLight")%>';
        TD_NAV_Brands.style.backgroundColor = '<%=Application("ColorSelectorDark")%>';
        
        TBL_NAV_BRANDS.style.display = 'none';
        TBL_NAV_CLIENTS.style.display = '';
    }
    function mostrarBrands()
    {
        TD_NAV_Clients.style.backgroundColor = '<%=Application("ColorSelectorDark")%>';
        TD_NAV_Brands.style.backgroundColor = '<%=Application("ColorSelectorLight")%>';

        TBL_NAV_BRANDS.style.display = '';
        TBL_NAV_CLIENTS.style.display = 'none';
    }
    
    function prevMonth()
    {
        if (thisForm.FILTER_STARTMONTH.value == 1)
        {
            thisForm.FILTER_YEAR.selectedIndex = thisForm.FILTER_YEAR.selectedIndex - 1;
            thisForm.FILTER_STARTMONTH.selectedIndex = 11; 
        }else{
            thisForm.FILTER_STARTMONTH.selectedIndex = thisForm.FILTER_STARTMONTH.selectedIndex - 1; 
        }
        thisForm.submit();
    }

    function nextMonth()
    {
        if (thisForm.FILTER_STARTMONTH.value == 12)
        {
            thisForm.FILTER_YEAR.selectedIndex = thisForm.FILTER_YEAR.selectedIndex + 1;
            thisForm.FILTER_STARTMONTH.selectedIndex = 0; 
        }else{
            thisForm.FILTER_STARTMONTH.selectedIndex = thisForm.FILTER_STARTMONTH.selectedIndex + 1; 
        }
        thisForm.submit();
    }
</script>
<div ID="MENU_REPORT_NAVIGATION">
    
    <table width="100%" bgcolor="<%=Application("ColorSelectorDark")%>" cellpadding="1" cellspacing="0">
        <tr>
            <td width="20px;"></td>
            <td width="160px;" style="text-align:center;cursor:pointer;" onmouseover="mostrarClients();" id="TD_NAV_Clients"><font class="font15">REPORT CLIENTE</font></td>
            <td width="20px;"></td>
            <td width="160px;" style="text-align:center;cursor:pointer;" onmouseover="mostrarBrands();" id="TD_NAV_Brands"><font class="font15">REPORT MARCA</font></td>
            <td></td>
            <td align="right" style="padding-right:10px;"><font class="font12"><a href="" onclick="prevMonth(); return false;"><font color="white"><%=IDM_PrevMonth %></font></a> | <a href="" onclick="nextMonth(); return false;"><font color="white"><%=IDM_NextMonth %></font></a></td>
        </tr>
    </table>
    
    <table width="100%" bgcolor="<%=Application("ColorSelectorLight")%>" id="TBL_NAV_CLIENTS">
        <tr>
            <td width="100%"><font class="font11">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <%
            dim sSelected1, sSelected2
            arrClients = getClients("NOMBRE")
            for each c in arrClients
                sSelected1 = ""
                sSelected2 = ""
                if isInArray(Request.Form("FILTER_MULTICLIENT"), c.IDClient) then
                    if Request.Form("FILTER_CLIENT")<>"" then
                        if CInt(Request.Form("FILTER_CLIENT")) = c.IDClient then
                            sSelected1 = "<font color=red>"
                            sSelected2 = "</font>"
                        end if
                    end if
                    %>&nbsp;&nbsp;|&nbsp;&nbsp;
                    <a href="" onclick="if (thisForm.Quick_Tipo2.checked){thisForm.Quick_Tipo1.checked=true;}thisForm.FILTER_REPORTTYPE.value='0';thisForm.FILTER_BRAND.value='';thisForm.FILTER_CLIENT.value='<%=c.IDClient %>';thisForm.submit();return false;">
                        <font style="cursor:pointer;">
                            <%=sSelected1%><%=c.Name %><%=sSelected2 %>
                        </font>
                    </a><%
                end if
            next
            %>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            </font>
            </td>
        </tr>
    </table>
    
    <table width="100%" bgcolor="<%=Application("ColorSelectorLight")%>" id="TBL_NAV_BRANDS">
        <tr>
            <td width="100%"><font class="font11">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <%
            arrBrands = getBrands("NOMBRE")
            for each b in arrBrands
                sSelected1 = ""
                sSelected2 = ""
                if isInArray(Request.Form("FILTER_MULTIBRAND"), b.IDBrand) then
                    if Request.Form("FILTER_BRAND")<>"" then
                        if CInt(Request.Form("FILTER_BRAND")) = b.IDBrand then
                            sSelected1 = "<font color=red>"
                            sSelected2 = "</font>"
                        end if
                    end if
                    %>&nbsp;&nbsp;|&nbsp;&nbsp;
                    <a href="" onclick="if (thisForm.Quick_Tipo1.checked){thisForm.Quick_Tipo2.checked=true;}   thisForm.FILTER_REPORTTYPE.value='1';thisForm.FILTER_CLIENT.value='';thisForm.FILTER_BRAND.value='<%=b.IDBrand %>';thisForm.submit();return false;">
                        <font style="cursor:pointer;">
                            <%=sSelected1 %><%=b.Name %><%=sSelected2 %>
                        </font>
                    </a><%
                end if
            next
            %>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            </font>
            </td>
            <td></td>
        </tr>
    </table>
    
    <br /><br />
</div>

<script language="javascript">
<%if Request.Form("FILTER_REPORTTYPE") = "0" then%>
    mostrarClients();
<%else %>
    mostrarBrands();
<%end if%>
</script>

