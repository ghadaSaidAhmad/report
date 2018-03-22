
<script runat=server language=vbscript>
</script>

<script language=javascript>
    function exportExcel()
    {
        thisForm.action = '?XL=1';
        thisForm.target = '_blank';
        thisForm.submit();
        thisForm.action = '';
        thisForm.target = '_self';
    }
    
    function imprimir()
    {
        MAIN_MENU.style.display = 'none';
        MENU_REPORT_NAVIGATION.style.display = 'none';
        TBL_TOP.style.display = 'none';
        print();
        MAIN_MENU.style.display = '';
        MENU_REPORT_NAVIGATION.style.display = '';
        TBL_TOP.style.display = '';
    }
    
    function toggleReportType()
    {
        if (thisForm.FILTER_REPORTTYPE.selectedIndex==0){
            thisForm.FILTER_BRAND.selectedIndex = 0;
        }else if (thisForm.FILTER_REPORTTYPE.selectedIndex==1){
            thisForm.FILTER_CLIENT.selectedIndex = 0;
        }
    }
    
    function applyFilter(selectFirstElement)
    {
        var selectedClients=0, selectedBrands=0, selectedActivityTypes=0;
        var firstSelectedClient=0, firstSelectedBrand=0;
        for (i=0;i<thisForm.FILTER_MULTICLIENT.length;i++)
        {
            if (thisForm.FILTER_MULTICLIENT.options[i].selected){
                if (firstSelectedClient==0)
                {
                    firstSelectedClient = thisForm.FILTER_MULTICLIENT.options[i].value;
                }
                selectedClients++
            }
        }
        for (i=0;i<thisForm.FILTER_MULTIBRAND.length;i++)
        {
            if (thisForm.FILTER_MULTIBRAND.options[i].selected)
            {
                if (firstSelectedBrand==0)
                {
                    firstSelectedBrand = thisForm.FILTER_MULTIBRAND.options[i].value;
                }
                selectedBrands++
            }
        }
        for (i=0;i<thisForm.FILTER_MULTIACTIVITYTYPE.length;i++)
            if (thisForm.FILTER_MULTIACTIVITYTYPE.options[i].selected){selectedActivityTypes++}
        

        if (selectedBrands == 0){
            alert('<%=IDM_JS_SelectSomeBrand %>');
            cambiarPest(Pest05, TBL_FILTER_CLIBRA);
            return false;
        }
        if (selectedClients == 0){
            alert('<%=IDM_JS_SelectSomeClient %>');
            cambiarPest(Pest05, TBL_FILTER_CLIBRA);
            return false;
        }
        if (selectedActivityTypes == 0)
        {
            alert('<%=IDM_JS_SelectSomeActivityType %>');
            cambiarPest(Pest20, TBL_FILTER_DATOS);
            return false;
        }
        
        if (selectFirstElement)
        {
            if (thisForm.FILTER_REPORTTYPE.selectedIndex == 0)
                thisForm.FILTER_CLIENT.value = firstSelectedClient;
            if (thisForm.FILTER_REPORTTYPE.selectedIndex == 1)
                thisForm.FILTER_BRAND.value = firstSelectedBrand;
        }


        _fireEvent('ApplyFilter','','');
        
        thisForm.style.display = 'none';
        DIV_WAIT.style.display = '';
    }
    
    function selectAll(oList, status)
    {   
        for (i = 0; i < oList.options.length; i++)
        {
            oList.options[i].selected = status;
        }
    }

	function cambiarPest(pest, div)
	{
	    TBL_FILTER_CLIBRA.style.display='none';
	    Pest05.className='PEST_NOSELEC';
	    TBL_FILTER_ORG.style.display='none';
	    Pest10.className='PEST_NOSELEC';
	    TBL_FILTER_DATOS.style.display='none';
	    Pest20.className='PEST_NOSELEC';
	    
	    // Muestra el DIV de la pestaña seleccionada
	    div.style.display='';
	    pest.className='PEST_SELEC';
	    
	}
    
    function showMenu()
    {
        MAIN_MENU.style.display = '';
    }
    
    function showFilter()
    {
        try{
            MENU_REPORT_NAVIGATION.style.display = 'none';
        }catch(e){}
        try{
            TBL_MAIN.style.display = 'none';
        }catch(e){}
        try{
            thisForm.BTN_Print.style.display = 'none';
        }catch(e){}
        try{
            thisForm.BTN_Export.style.display = 'none';
        }catch(e){}
        MAIN_FILTER.style.display = '';
        
        document.body.style.backgroundImage = 'url(\'images/background.jpg\')';
        document.body.style.backgroundRepeat = 'no-repeat';
    }
    
    function closeFilter()
    {
        try{
            MENU_REPORT_NAVIGATION.style.display = '';
        }catch(e){}
        try{
            TBL_MAIN.style.display = '';
        }catch(e){}
        try{
            thisForm.BTN_Print.style.display = '';
        }catch(e){}
        try{
            thisForm.BTN_Export.style.display = '';
        }catch(e){}
        MAIN_FILTER.style.display = 'none';

        document.body.style.backgroundImage = '';
        document.body.style.backgroundRepeat = '';
    }
</script>

<%
dim menuWidth
menuWidth = 130
if IsAdmin() then
    menuWidth = 190
end if
if IsAdmin() OR IsInputData() then
    menuWidth = menuWidth + 60
end if
%>

<div ID="MAIN_MENU" style="display:none;text-align:right;position:absolute;right:40px;top:0px;width:<%=menuWidth%>px;height:40px;background-color:White;border-top:3 solid Gray;border-left:3 solid Gray;border-right:3 solid Gray;border-bottom:3 solid Gray;">
    <%if IsAdmin() then %>
        <img onclick="location.href='AdminAppVariables.asp';return false;" alt="<%=IDM_MenuConfig %>" style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/config.png" />
        &nbsp;&nbsp;&nbsp;&nbsp;
    <%end if %>
    <%if IsAdmin() OR IsInputData() then %>
        <img onclick="location.href='RealData.asp';return false;" alt="<%=IDM_MenuInputData %>" style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/btnedit.png" height=30 width=30 />
    <%end if %>
    <img onclick="showFilter();" alt="<%=IDM_MenuFilterReport %>" style="width:30px;height:30px;cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/report.png" />
    <%if Request.Form("FILTER_REPORTTYPE") <> "" then %>
        <img ID="BTN_Print" onclick="imprimir();" alt="<%=IDM_MenuImprimir %>" style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/print.png" />
        <img ID="BTN_Export" onclick="exportExcel();" alt="<%=IDM_MenuExportar %>" style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/icon_excel.png" />
    <%end if %>
</div>


<%
dim MainFilterWidth: MainFilterWidth = 500
%>
<div ID="MAIN_FILTER" style="<%if Request.Form("FILTER_REPORTTYPE") <> "" then %>display:none;<%end if %>position:absolute;top:90px;width:<%=MainFilterWidth%>px;height:450px;background-color:White;border-left:1 solid black;border-top:1 solid black;border-bottom:4 solid black;border-right:4 solid black;" >

    <table border=0 cellpadding=0 cellspacing=0 width="100%">
    <tr height="20px;">
	    <td class="filterTopBar" align="left" style="padding-left:5"><%=IDM_FilterTopBarTitle %></td>
    </tr>
    </table>
    
    <table border=0 cellpadding=0 cellspacing=0 width="100%">
    <tr>
	    <td align="left" style="padding-left:5" rowspan="2" valign="top">
		    <img src="images/report.png" />
		</td>
		<td align="left" style="padding-left:25px;">
		    <FONT class="font20"><strong><%=IDM_FilterTitle%></strong></FONT>
	    </td>
	</tr>
	<tr>
	    <td align="left" style="padding-left:25px;">
	        <font class="font12"><%=IDM_FilterSubTitle %></font>
	    </td>
	</tr>
	</table>

    
    <br />
    <table ID="DIV_PESTANAS" style="width:100%;height:20px;" cellpadding=0 cellspacing=0><tr>
        <td width="20" class="PEST_ESPACIO"><font class=font9>&nbsp;</font></td>
        <td width="80" ID="Pest05" class="PEST_SELEC" onclick="cambiarPest(Pest05, TBL_FILTER_CLIBRA);return false;"><%=IDM_FilterPest05 %></td>
        <td width="10" class="PEST_ESPACIO"><font class=font9>&nbsp;</font></td>
        <td width="80" ID="Pest10" class="PEST_NOSELEC" onclick="cambiarPest(Pest10, TBL_FILTER_ORG);return false;"><%=IDM_FilterPest1 %></td>
        <td width="10" class="PEST_ESPACIO"><font class=font9>&nbsp;</font></td>
        <td width="90" ID="Pest20" class="PEST_NOSELEC" onclick="cambiarPest(Pest20, TBL_FILTER_DATOS);return false;"><%=IDM_FilterPest2 %></td>
        <td class="PEST_ESPACIO">&nbsp;</td>
    </tr></table>
    <br />
    
    
    <div ID="TBL_FILTER_CLIBRA">
        <table width="100%" >
            <tr>
                <td width=120 class="fieldheader"><%=IDM_FilterReport %></td>
                <td>
                    <select name="FILTER_REPORTTYPE" style="width:200px;" onchange="toggleReportType();">
                        <option value="0" <%if Request.Form("FILTER_REPORTTYPE") = "0" then %>selected<%end if %>><%=IDM_FilterReportType0 %></option>
                        <option value="1" <%if Request.Form("FILTER_REPORTTYPE") = "1" then %>selected<%end if %>><%=IDM_FilterReportType1 %></option>
                    </select>
                </td>
            </tr>

            <tr id="TR_SELECTCLIENT" style="display:none;">
                <td width=120 class="fieldheader"><%=IDM_Client %></td>
                <td>
                    <select style="width:200px;" name="FILTER_CLIENT">
                        <option value=""><%=IDM_SelectClient %></option>
                        <%
                        dim lstcli
                        dim arrClients
                        arrClients = getClients()
                        for each lstcli in arrClients
                            sSelected = ""
                            if Request.Form("FILTER_CLIENT")<>"" then
                                if CInt(Request.Form("FILTER_CLIENT")) = lstcli.IDClient then
                                    sSelected = "selected"
                                end if
                            end if
                            %><option value="<%=lstcli.IDClient %>" <%=sSelected %>><%=lstcli.Name %></option><%
                        next
                        %>
                    </select>
                </td>
            </tr>
            <tr id="TR_SELECTBRAND" style="display:none;">
                <td width=120 class="fieldheader"><%=IDM_Brand %></td>
                <td>
                    <select style="width:200px;" name="FILTER_BRAND">
                        <option value=""><%=IDM_SelectBrand %></option>
                        <%
                        dim lstbra
                        dim arrBrands
                        arrBrands = getBrands()
                        for each lstbra in arrBrands
                            sSelected = ""
                            if Request.Form("FILTER_BRAND")<>"" then
                                if CInt(Request.Form("FILTER_BRAND")) = lstbra.IDBrand then
                                    sSelected = "selected"
                                end if
                            end if
                            %><option value="<%=lstbra.IDBrand %>" <%=sSelected %>><%=lstbra.Name %></option><%
                        next
                        %>
                    </select>
                </td>
            </tr>
            <tr id="TR_FILTER_MULTIBRAND">
                <td width=120 class="fieldheader" valign=top><%=IDM_Brands %>
                </td>
                <td>
                    <select id="FILTER_MULTIBRAND" name="FILTER_MULTIBRAND" multiple style="width:200px;height:100px;" class=textfield>
                        <%
                        dim splitBrands
                        splitBrands = split(Request.Form("FILTER_MULTIBRAND"), ",")
                        
                        arrBrands = getBrands()
                        for each lstbra in arrBrands
                            sSelected = ""
                            if Request.Form("FILTER_MULTIBRAND")<>"" then
                                if isInArray(splitBrands, lstbra.IDBrand) then
                                    sSelected = "selected"
                                end if
                            end if
                            %><option value="<%=lstbra.IDBrand %>" <%=sSelected %>><%=lstbra.Name %></option><%
                        next
                        %>
                    </select>
                </td>
                <td valign=top>
                    <font class=font10>
                    <input type="button" onclick="selectAll(thisForm.FILTER_MULTIBRAND, true);return false;" value="<%=IDM_FilterSelectAll %>" class="button" style="width:100px;" />
                    <br />
                    <input type="button" onclick="selectAll(thisForm.FILTER_MULTIBRAND, false);return false;" value="<%=IDM_FilterUnselectAll %>" class="button" style="width:100px;" />
                    
                    <br /><br />
                    <a title="<%=IDM_JS_ListKeepPressedCtrl %>" href="" onclick="alert('<%=IDM_JS_ListKeepPressedCtrl %>');return false;">¿selección múltiple?</a>
                    </font>
                </td>
            </tr>
            <tr id="TR_FILTER_MULTICLIENT">
                <td width=120 class="fieldheader" valign=top><%=IDM_Clients %>
                </td>
                <td>
                    <select id="FILTER_MULTICLIENT" name="FILTER_MULTICLIENT" multiple style="width:200px;height:100px;" class=textfield>
                        <%
                        dim splitClients
                        splitClients = split(Request.Form("FILTER_MULTICLIENT"), ",")
                        
                        
                        arrClients = getClients()
                        for each lstcli in arrClients
                            sSelected = ""
                            if Request.Form("FILTER_MULTICLIENT")<>"" then
                                if isInArray(splitClients, lstcli.IDClient) then
                                    sSelected = "selected"
                                end if
                            end if
                            %><option value="<%=lstcli.IDClient %>" <%=sSelected %>><%=lstcli.Name %></option><%
                        next
                        %>
                    </select>
                </td>
                <td valign=top>
                    <font class=font10>
                    <input type="button" onclick="selectAll(thisForm.FILTER_MULTICLIENT, true);return false;" value="<%=IDM_FilterSelectAll %>" class="button" style="width:100px;" />
                    <br />
                    <input type="button" onclick="selectAll(thisForm.FILTER_MULTICLIENT, false);return false;" value="<%=IDM_FilterUnselectAll %>" class="button" style="width:100px;" />

                    <br /><br />
                    <a title="<%=IDM_JS_ListKeepPressedCtrl %>" href="" onclick="alert('<%=IDM_JS_ListKeepPressedCtrl %>');return false;">¿selección múltiple?</a>
                    </font>
                </td>
            </tr>
        </table>
        
        <br />
        <%if Request.Form("FILTER_REPORTTYPE") <> "" then %>
            <div style="position:absolute;bottom:10px;left:20px;">
                <input class=button type=button value="<%=IDM_FilterClose %>" onclick="closeFilter();" />
            </div>
        <%end if %>
        <div style="position:absolute;bottom:10px;right:20px;">
            <input type=button class=button value="<%=IDM_FilterNext %>" onclick="cambiarPest(Pest10, TBL_FILTER_ORG);" />
            <input class=button type=button value="<%=IDM_FilterApply %>" onclick="applyFilter(true);" />
        </div>

    </div>
    
    <div ID="TBL_FILTER_ORG" style="display:none;">
        <table width="100%" >
            <tr>
                <td width="160" class="fieldheader"><%=IDM_FilterStart %></td>
                <td>
                    <select name="FILTER_YEAR" style="width:100px;">
                        <%
                        dim iYear, sSelected
                        for iYear = 2008 to Year(Date) + 1 
                            sSelected = ""
                            if Request.Form("FILTER_YEAR")<>"" then
                                if CInt(Request.Form("FILTER_YEAR")) = iYear then
                                    sSelected = "selected"
                                end if
                            elseif iYear = Year(Date) then
                                sSelected = "selected"
                            end if
                            %>
                            <option value="<%=iYear %>" <%=sSelected %>><%=iYear %></option>
                        <%next %>
                    </select>
                    <select name="FILTER_STARTMONTH"  style="width:100px;">
                        <%
                        dim iMonth
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
            </tr>
            <tr>
                <td class="fieldheader"><%=IDM_FilterMonths %></td>
                <td>
                    <select name="FILTER_VIEWMONTHS"  style="width:100px;">
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
            </tr>
            <tr>
                <td width="140" class="fieldheader"><%=IDM_FilterLastYear %></td>
                <%
                dim sChecked
                if Request.Form("FILTER_LASTYEAR")<>"" then
                    sChecked = "checked"
                elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ViewLastYear")="YES" then
                    sChecked = "checked"
                end if
                %>
                <td>
                    <input style="width:30px;height:30px;" type="checkbox" name="FILTER_LASTYEAR" <%=sChecked %> />
                </td>
            </tr>
            <tr>
                <td class="fieldheader" ><%=IDM_FilterSaltoCada %></td>
                <%
                dim sBreakEach
                if Request.Form("FILTER_VIEWMONTHS")<>"" then
                    sBreakEach = Request.Form("FILTER_BREAKE_EACH")
                else
                    sBreakEach = Application("Default_BreakEach")
                end if
                %>
                <td>
                    <input style="width:30px;" type=text name="FILTER_BREAKE_EACH" value="<%=sBreakEach %>" />
                    <font style="font-family:Arial;font-size:12px;font-weight:bold;"></font>
                </td>
            </tr>
        </table>

        <br />
        <%if Request.Form("FILTER_REPORTTYPE") <> "" then %>
            <div style="position:absolute;bottom:10px;left:20px;">
                <input class=button type=button value="<%=IDM_FilterClose %>" onclick="closeFilter();" />
            </div>
        <%end if %>
        <div style="position:absolute;bottom:10px;right:20px;">
            <input type=button class=button value="<%=IDM_FilterPrevious %>" onclick="cambiarPest(Pest05, TBL_FILTER_CLIBRA);" />
            <input type=button class=button value="<%=IDM_FilterNext %>" onclick="cambiarPest(Pest20, TBL_FILTER_DATOS);" />
            <input class=button type=button value="<%=IDM_FilterApply %>" onclick="applyFilter(true);" />
        </div>

    </div>
    
    <div id="TBL_FILTER_DATOS" style="display:none;">
        <table width="100%" >
            <tr>
                <%
                sChecked = ""
                if Request.Form("FILTER_SHOWREALDATA_NSHOPS")<>"" then
                    sChecked = "checked"
                elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ShowRealData_NShops")="YES" then
                    sChecked = "checked"
                end if
                %>
                <td width=120 class="fieldheader"><%=IDM_NTiendasReal %></td>
                <td><input style="width:30px;height:30px;" type=checkbox name="FILTER_SHOWREALDATA_NSHOPS" <%=sChecked%> /></td>
                <%
                sChecked = ""
                if Request.Form("FILTER_SHOWREALDATA_PERCENTCOMPLAINT")<>"" then
                    sChecked = "checked"
                elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ShowRealData_PercentComplaint")="YES" then
                    sChecked = "checked"
                end if
                %>
                <td width=120 class="fieldheader"><%=IDM_PercentComplaint %></td>
                <td><input style="width:30px;height:30px;" type=checkbox name="FILTER_SHOWREALDATA_PERCENTCOMPLAINT" <%=sChecked%> /></td>
            </tr>
            <tr>
                <%
                sChecked = ""
                if Request.Form("FILTER_SHOWGENERALTHEME")<>"" then
                    sChecked = "checked"
                elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ShowGeneralTheme")="YES" then
                    sChecked = "checked"
                end if
                %>
                <td width=120 class="fieldheader"><%=IDM_GeneralTheme %></td>
                <td><input style="width:30px;height:30px;" type=checkbox name="FILTER_SHOWGENERALTHEME" <%=sChecked%> /></td>
                <%
                sChecked = ""
                if Request.Form("FILTER_SHOWNR")<>"" then
                    sChecked = "checked"
                elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ShowNR")="YES" then
                    sChecked = "checked"
                end if
                %>
                <td width=120 class="fieldheader">NR</td>
                <td><input style="width:30px;height:30px;" type=checkbox name="FILTER_SHOWNR" <%=sChecked%> /></td>
            </tr>
            <tr>
                <%
                sChecked = ""
                if Request.Form("FILTER_SHOWFC")<>"" then
                    sChecked = "checked"
                elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ShowFC")="YES" then
                    sChecked = "checked"
                end if
                %>
                <td width=120 class="fieldheader">FC</td>
                <td><input style="width:30px;height:30px;" type=checkbox name="FILTER_SHOWFC" <%=sChecked%> /></td>
                <%
                sChecked = ""
                if Request.Form("FILTER_SHOWNRVSLY")<>"" then
                    sChecked = "checked"
                elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ShowNRvsLY")="YES" then
                    sChecked = "checked"
                end if
                %>
                <td width=120 class="fieldheader">%NR vs LY</td>
                <td><input style="width:30px;height:30px;" type=checkbox name="FILTER_SHOWNRVSLY" <%=sChecked%> /></td>
            </tr>
        </table>
        
        <hr color="Silver"/>
        
        <table width="100%" style="display:none;">
            <tr>
                <td width=120 class="fieldheader" valign=top>Tipos
                </td>
                <td >
                    <select name="FILTER_MULTIACTIVITYTYPE" multiple style="width:200px;height:100px;">
                        <%
                        dim arrTipos
                        dim tip
                        dim splitActivityTypes
                        splitActivityTypes = split(Request.Form("FILTER_MULTIACTIVITYTYPE"), ",")
                        arrTipos = getActivityTypes(Idioma)
                        for each tip in arrTipos
                            sSelected = ""
                            if Request.Form("FILTER_MULTIACTIVITYTYPE")="" then
                                sSelected = "selected"
                            else
                                if isInArray(splitActivityTypes, tip.ID) then
                                    sSelected = "selected"
                                end if
                            end if
                            %><option value="<%=tip.ID %>" <%=sSelected %>><%=tip.Name %></option><%
                        next
                        %>
                        
                    </select>
                </td>
                <td valign=top>
                    <input type="button" onclick="selectAll(thisForm.FILTER_MULTIACTIVITYTYPE, true);return false;" value="<%=IDM_FilterSelectAll %>" class="button" style="width:100px;" />
                    <br />
                    <input type="button" onclick="selectAll(thisForm.FILTER_MULTIACTIVITYTYPE, false);return false;" value="<%=IDM_FilterUnselectAll %>" class="button" style="width:100px;" />
                </td>
            </tr>
        </table>

        <br />
        <%if Request.Form("FILTER_REPORTTYPE") <> "" then %>
            <div style="position:absolute;bottom:10px;left:20px;">
                <input class=button type=button value="<%=IDM_FilterClose %>" onclick="closeFilter();" />
            </div>
        <%end if %>
        <div style="position:absolute;bottom:10px;right:20px;">
            <input type=button class=button value="<%=IDM_FilterPrevious %>" onclick="cambiarPest(Pest10, TBL_FILTER_ORG);" />
            <input type=button class=button value="<%=IDM_FilterApply %>" onclick="applyFilter(true);" />
        </div>

    </div>
    
    
</div>



<script language=javascript>
    toggleReportType();
    
    MAIN_FILTER.style.left = (screen.width / 2) - <%=MainFilterWidth %>/2 ;
    
</script>