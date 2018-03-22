<script language="javascript">
    function mostrarClients()
    {
        TD_NAV_Clients.style.backgroundColor = '<%=Application("ColorSelectorLight")%>';
        
        TBL_NAV_CLIENTS.style.display = '';
    }

    function prevYear()
    {
        if (checkChanges()) {
            thisForm.FILTER_YEAR.value = parseInt(thisForm.FILTER_YEAR.value) - 1;
            thisForm.EventObject.value='ChYe';
            thisForm.submit();
        }
    }

    function nextYear()
    {
        if (checkChanges()) {
            thisForm.FILTER_YEAR.value = parseInt(thisForm.FILTER_YEAR.value) + 1;
            thisForm.EventObject.value='ChYe';
            thisForm.submit();
        }
    }
    
</script>
<div ID="MENU_REALDATA_NAVIGATION">
    
    <table width="100%" bgcolor="<%=Application("ColorSelectorDark")%>" cellpadding="1" cellspacing="0">
        <tr>
            <td width="20px;"></td>
            <td width="160px;" style="text-align:center;cursor:pointer;" onmouseover="mostrarClients();" id="TD_NAV_Clients"><font class="font15">CLIENTE</font></td>
            <td></td>
            <td width=250 align="right" style="padding-right:10px;">
                <input type="checkbox" name="ViewNShopsActivity" <%if (Request.Form("ViewNShopsActivity")<>"" AND Request.Form("PageReloaded")<>"") OR Request.Form("PageReloaded")="" then %>checked<%end if %> onclick="if (checkChanges()){ _fireEvent('ChSel','',''); }else{this.checked = !this.checked}" /><font class="font12" color="white"><%=IDM_NTiendasShort %></font>
                <input type="checkbox" name="ViewNShops" <%if (Request.Form("ViewNShops")<>"" AND Request.Form("PageReloaded")<>"") OR Request.Form("PageReloaded")="" then %>checked<%end if %> onclick="if (checkChanges()){ _fireEvent('ChSel','',''); }else{this.checked = !this.checked}" /><font class="font12" color="white"><%=IDM_NTiendasRealShort %></font>
                <input type="checkbox" name="ViewPercentComplaint" <%if (Request.Form("ViewPercentComplaint")<>"" AND Request.Form("PageReloaded")<>"") OR Request.Form("PageReloaded")="" then %>checked<%end if %> onclick="if (checkChanges()){ _fireEvent('ChSel','',''); }else{this.checked = !this.checked}" /><font class="font12" color="white"><%=IDM_PercentComplaintShort %></font>
            </td>
            <td>
                <select name="FILTER_STARTMONTH" style="width:100px;" onchange="if (checkChanges()){ _fireEvent('Ch_St','',''); }else{return false;}">
                    <%
                    dim iMonth, sSelected
                    for iMonth = 1 To 12
                        sSelected = ""
                        if Request.Form("FILTER_STARTMONTH")<>"" then
                            if CInt(Request.Form("FILTER_STARTMONTH")) = iMonth then
                                sSelected = "selected"
                            end if
                        elseif iMonth = Month(Date) then
                            sSelected = "selected"
                        end if
                        %><option value="<%=iMonth %>" <%=sSelected %>><%=locMonthName(iMonth, Idioma) %></option><%
                    next
                    %>
                </select>
                <select name="FILTER_VIEWMONTHS" style="width:40px;" onchange="if (checkChanges()){ _fireEvent('Ch_Vw','',''); }else{return false;}">
                    <%
                    dim NMonths
                    for NMonths = 1 To 13
                        sSelected = ""
                        if Request.Form("FILTER_VIEWMONTHS")<>"" then
                            if CInt(Request.Form("FILTER_VIEWMONTHS")) = NMonths then
                                sSelected = "selected"
                            end if
                        elseif NMonths = CInt(Application("Default_NumMonths")) then
                            sSelected = "selected"
                        end if
                        %><option value="<%=NMonths %>" <%=sSelected %>><%=NMonths %></option><%
                    next
                    %>
                </select>
            </td>
            <td align="right" style="padding-right:10px;"><font class="font12"><a href="" onclick="prevYear(); return false;"><font color="white"><%=IDM_PrevYear %></font></a> | <a href="" onclick="nextYear(); return false;"><font color="white"><%=IDM_NextYear %></font></a></td>
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
                if Request.Form("FILTER_CLIENT")<>"" then
                    if CInt(Request.Form("FILTER_CLIENT")) = c.IDClient then
                        sSelected1 = "<font color=red>"
                        sSelected2 = "</font>"
                    end if
                end if
                %>&nbsp;&nbsp;|&nbsp;&nbsp;
                <a href="" onclick="if (checkChanges()) { thisForm.FILTER_CLIENT.value='<%=c.IDClient %>';thisForm.EventObject.value='ChCli';thisForm.submit();} return false;">
                    <font style="cursor:pointer;">
                        <%=sSelected1%><%=c.Name %><%=sSelected2 %>
                    </font>
                </a><%
            next
            %>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            </font>
            </td>
        </tr>
    </table>
    
    
</div>

<script language="javascript">
    mostrarClients();
</script>

