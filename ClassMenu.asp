
<script runat=server language=vbscript>
</script>

<script language="javascript" type="text/javascript">
    function exportExcel()
    {
        thisForm.action = 'SOAXL.asp';
        thisForm.target = '_blank';
        thisForm.submit();
        thisForm.action = '';
        thisForm.target = '_self';
    }
    
    function imprimir()
    {
        if (document.getElementById('MENU_REPORT_NAVIGATION').style.display != 'none'){
            document.getElementById('TOPMARGIN').style.display = 'none';
            document.getElementById('DIV_TOP_BAR').style.display = 'none';
            document.getElementById('MENU_REPORT_NAVIGATION').style.display = 'none';
            print();
            document.getElementById('TOPMARGIN').style.display = '';
            document.getElementById('DIV_TOP_BAR').style.display = '';
            document.getElementById('MENU_REPORT_NAVIGATION').style.display = '';
        }else{
            alert('Visualice primero en pantalla el report a imprimir');
        }
    }
    
    function toggleReportType()
    {
        if (thisForm.FILTER_REPORTTYPE.selectedIndex==0){
            thisForm.FILTER_BRAND.selectedIndex = 0;
        }else if (thisForm.FILTER_REPORTTYPE.selectedIndex==1){
            thisForm.FILTER_CLIENT.selectedIndex = 0;
        }
    }
    
    function applyFilter(selectFirstElement, excelexport)
    {
        var selectedClients=0, selectedBrands=0;
        var firstSelectedClient=0, firstSelectedBrand=0;
        
        <%if Request.Form("FILTER_REPORTTYPE")="" then%>
            if (!thisForm.Quick_Tipo1.checked && !thisForm.Quick_Tipo2.checked && !thisForm.Quick_Personalizado.checked)
            {
                alert('<%=IDM_JS_SelectQuickReport%>');
                
                return false;
            }
        <%end if %>
        
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
        
        if (selectFirstElement)
        {
            if (thisForm.FILTER_REPORTTYPE.selectedIndex == 0)
                thisForm.FILTER_CLIENT.value = firstSelectedClient;
            if (thisForm.FILTER_REPORTTYPE.selectedIndex == 1)
                thisForm.FILTER_BRAND.value = firstSelectedBrand;
        }

        if (excelexport){
            
            if (thisForm.Quick_Tipo1.checked || thisForm.Quick_Tipo2.checked){
                alert('Debe crear un report personalizado y seleccionar un cliente o una marca');
                return false;
            }
            
            // Si quieren un report de 1 cliente - varias marcas, tienen que tener seleccionado SÓLO 1 cliente
            if (thisForm.FILTER_REPORTTYPE.selectedIndex == 0 && selectedClients>1){
                alert('Seleccione sólo el cliente que quiere exportar');
                cambiarPest(Pest05, TBL_FILTER_CLIBRA);
                return false;
            }
            // Si quieren un report de 1 marca - varios clientes, tienen que tener seleccionada SÓLO 1 marca
            else if (thisForm.FILTER_REPORTTYPE.selectedIndex == 1 && selectedBrands>1){
                alert('Seleccione sólo la marca que quiere exportar');
                cambiarPest(Pest05, TBL_FILTER_CLIBRA);
                return false;
            }
            
            exportExcel();
        }else{
            _fireEvent('ApplyFilter','','');
            
            thisForm.style.display = 'none';
            DIV_WAIT.style.display = '';
        }
        
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
	    TBL_FILTER_QUICK.style.display='none';
	    Pest02.className='PEST_NOSELEC';
	    TBL_FILTER_CLIBRA.style.display='none';
	    Pest05.className='PEST_NOSELEC';
	    TBL_FILTER_ORG.style.display='none';
	    Pest10.className='PEST_NOSELEC';
	    TBL_FILTER_DATOS.style.display='none';
	    Pest20.className='PEST_NOSELEC';
	    
	    try{ // Puede estar escondido
	        TBL_FILTER_EXTRAS.style.display='none';
	        Pest30.className='PEST_NOSELEC';
	    }catch(e){}
	    
	    // Muestra el DIV de la pestaña seleccionada
	    div.style.display='';
	    pest.className='PEST_SELEC';
	    
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
            //document.getElementById('BTN_Print').style.display = 'none';
        }catch(e){}
        try{
            //document.getElementById('BTN_Export').style.display = 'none';
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
            document.getElementById('BTN_Print').style.display = '';
        }catch(e){}
        try{
            document.getElementById('BTN_Export').style.display = '';
        }catch(e){}
        MAIN_FILTER.style.display = 'none';

        document.body.style.backgroundImage = '';
        document.body.style.backgroundRepeat = '';
    }
    
    function SelectorQuick(tipo)
    {
        thisForm.Quick_Tipo1.checked = false;
        thisForm.Quick_Tipo2.checked = false;
        thisForm.Quick_Personalizado.checked = false;

        if (tipo == 1){
            thisForm.Quick_Tipo1.checked = true;
            
            thisForm.FILTER_REPORTTYPE.selectedIndex = 0;
            selectAll(thisForm.FILTER_MULTICLIENT, true);
            selectAll(thisForm.FILTER_MULTIBRAND, true);
            thisForm.FILTER_YEAR.value = <%=Year(Date) %>;
            thisForm.FILTER_STARTMONTH.selectedIndex = 0;
            thisForm.FILTER_VIEWMONTHS.value = <%=12 %>;

            Pest05.style.display = 'none';
            Pest10.style.display = 'none';
            Pest20.style.display = 'none';
            thisForm.BTN_SIGUIENTE_QUICK.style.display = 'none';
        }else if (tipo == 2){
            thisForm.Quick_Tipo2.checked = true;

            thisForm.FILTER_REPORTTYPE.selectedIndex = 1;
            selectAll(thisForm.FILTER_MULTICLIENT, true);
            selectAll(thisForm.FILTER_MULTIBRAND, true);
            thisForm.FILTER_YEAR.value = <%=Year(Date) %>;
            thisForm.FILTER_STARTMONTH.selectedIndex = 0;
            thisForm.FILTER_VIEWMONTHS.value = <%=12 %>;

            Pest05.style.display = 'none';
            Pest10.style.display = 'none';
            Pest20.style.display = 'none';
            thisForm.BTN_SIGUIENTE_QUICK.style.display = 'none';
        }else{
            thisForm.Quick_Personalizado.checked = true;
            ///// thisForm.FILTER_REPORTTYPE.selectedIndex = 0;
            
            
            //selectAll(thisForm.FILTER_MULTICLIENT, false);
            //selectAll(thisForm.FILTER_MULTIBRAND, false);
            
            Pest05.style.display = '';
            Pest10.style.display = '';
            Pest20.style.display = '';
            thisForm.BTN_SIGUIENTE_QUICK.style.display = '';
        }
    }


    function falseTextBoxLastYear()
    {
        thisForm.FILTER_LASTYEAR.checked = thisForm.FalseTextBox_LastYear.checked;
    }
    function lanzarQuery(){
        if (document.getElementById('IDReportQuery').selectedIndex > 0){
            location.href='ExportQuery.asp?q=' + document.getElementById('IDReportQuery').value;
        }else{
            alert('<%=IDM_BtnReportQuery_AlertJS %>');
        }
    }
    
</script>

<!-- #include file = "ClassTopButtons.asp" -->


<%
dim MainFilterWidth: MainFilterWidth = 500
%>
<div ID="MAIN_FILTER" style="left:300px;<%if Request.Form("FILTER_REPORTTYPE") <> "" then %>display:none;<%end if %>position:absolute;top:90px;width:<%=MainFilterWidth%>px;height:450px;background-color:#e7ebf6;border-left:1px solid black;border-top:1px solid black;border-bottom:1px solid black;border-right:1px solid black;margin:0;" >

    <table border=0 cellpadding=0 cellspacing=0 width="100%">
    <tr height="25px;">
	    <td class="filterTopBar" align="left" style="padding-left:5px,">
	        <div style="float:left;padding-left:2px;padding-top:0px;"><img src="images/form1.png" style="height:23px;"/></div>
	        <div style="float:left;padding-top:4px;padding-left:15px;"><%=IDM_FilterTopBarTitle %></div>
	    </td>
    </tr>
    </table>
    
    <table border=0 cellpadding=0 cellspacing=0 width="100%" style="background-color:white;">
    <tr>

		<td align="left" style="padding-left:25px;padding-top:10px;">
		    <FONT class="font20"><strong><%=IDM_FilterTitle%></strong></FONT>
	    </td>
	</tr>
	<tr style="height:60px;">
	    <td align="left" valign="top" style="padding-left:25px;padding-top:10px;">
	        <font class="font12"><%=IDM_FilterSubTitle %></font>
	    </td>
	</tr>
	</table>
	
    <table ID="DIV_PESTANAS" style="width:100%;height:20px;" cellpadding=0 cellspacing=0><tr>
        <td width="20" class="PEST_ESPACIO">&nbsp;</td>
        <td width="80" ID="Pest02" class="PEST_SELEC" onclick="cambiarPest(Pest02, TBL_FILTER_QUICK);return false;"><%=IDM_FilterPest02 %></td>
        <td width="5" class="PEST_ESPACIO">&nbsp;</td>
        <td width="80" style="display:none;" ID="Pest05" class="PEST_NOSELEC" onclick="cambiarPest(Pest05, TBL_FILTER_CLIBRA);return false;"><%=IDM_FilterPest05 %></td>
        <td width="5" class="PEST_ESPACIO">&nbsp;</td>
        <td width="80" style="display:none;" ID="Pest10" class="PEST_NOSELEC" onclick="cambiarPest(Pest10, TBL_FILTER_ORG);return false;"><%=IDM_FilterPest1 %></td>
        <td width="5" class="PEST_ESPACIO">&nbsp;</td>
        <td width="90" style="display:none;" ID="Pest20" class="PEST_NOSELEC" onclick="cambiarPest(Pest20, TBL_FILTER_DATOS);return false;"><%=IDM_FilterPest2 %></td>
        
        <%if isAdmin() then %>
            <td width="5" class="PEST_ESPACIO">&nbsp;</td>
            <td width="90" ID="Pest30" class="PEST_NOSELEC" onclick="cambiarPest(Pest30, TBL_FILTER_EXTRAS);return false;" style="color:red;"><%=IDM_FilterPest30 %></td>
        <%end if %>
        
        <td class="PEST_ESPACIO">&nbsp;</td>
    </tr></table>
    <br />
    
    <%
    dim sCheckedLastYear
    if Request.Form("FILTER_LASTYEAR")<>"" then
        sCheckedLastYear = "checked"
    elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ViewLastYear")="YES" then
        sCheckedLastYear = "checked"
    end if
    %>

    <div ID="TBL_FILTER_QUICK">
        
        <p align=center><font class=font20><b>Standard Reports</b></font></p>
        <table width="70%" align=center>
            <tr>
                <td >
                    <input onclick="SelectorQuick(1);" <%if Request.Form("Quick_Tipo")="1" then %>checked<%end if %> NAME="Quick_Tipo" ID="Quick_Tipo1" type="radio" style="height:30px;width:30px;" value="1" />
                    <font class=font20 onclick="SelectorQuick(1);return false;" style="cursor:pointer;">
                    <%=IDM_FilterReportType0Todas %>
                    </font>
                </td>
            </tr>
            <tr>
                <td >
                    <input onclick="SelectorQuick(2);" <%if Request.Form("Quick_Tipo")="2" then %>checked<%end if %> NAME="Quick_Tipo" ID="Quick_Tipo2" type="radio" style="height:30px;width:30px;" value="2" />
                    <font class=font20 onclick="SelectorQuick(2);return false;" style="cursor:pointer;">
                    <%=IDM_FilterReportType1Todas %>
                    </font>
                </td>
            </tr>
            <tr>
                <td >
                    <input onclick="SelectorQuick(99);" <%if Request.Form("Quick_Tipo")="99" then %>checked<%end if %> NAME="Quick_Tipo" ID="Quick_Personalizado" type="radio" style="height:30px;width:30px;" value="99" />
                    <font class=font20 onclick="SelectorQuick(99);return false;" style="cursor:pointer;">
                    <%=IDM_Personalizado %>
                    </font>
                </td>
            </tr>
            <tr>
                <td>
                    <div style="float:right;padding-top:5px;padding-left:5px;cursor:pointer;" onclick="document.getElementById('FalseTextBox_LastYear').checked = !document.getElementById('FalseTextBox_LastYear').checked;falseTextBoxLastYear();" ><font class=font15><%=IDM_FilterLastYear %></font></div>
                    <div style="float:right;"><input type="checkbox" style="width:30px;height:30px;" id="FalseTextBox_LastYear" name="FalseTextBox_LastYear" onclick="falseTextBoxLastYear();" <%=sCheckedLastYear %>/></div>
                </td>
            </tr>
        </table>
        

        <div style="position:absolute;bottom:0px;width:100%;height:50px;background-color:white;border-top:1px solid black;">
            <%if Request.Form("FILTER_REPORTTYPE") <> "" then %>
                <div style="position:absolute;bottom:10px;left:20px;">
                    <input class=button type=button value="<%=IDM_FilterClose %>" onclick="closeFilter();" />
                </div>
            <%end if %>
            <div style="position:absolute;bottom:10px;right:20px;">
                <input ID="BTN_SIGUIENTE_QUICK" style="display:none;" type=button class=button value="<%=IDM_FilterNext %>" onclick="cambiarPest(Pest05, TBL_FILTER_CLIBRA);" />

                <%if Application("BotonExportExcel") = "YES" then %>
                    <input class=button type=button value="<%=IDM_ExportExcel %>" onclick="applyFilter(true, true);" />
                <%end if %>
                <input class=button type=button value="<%=IDM_FilterApply %>" onclick="applyFilter(true, false);" />
            </div>
        </div>
    </div>
    
    <div ID="TBL_FILTER_CLIBRA" style="display:none;">
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
                        arrClients = getClients("ORDEN")
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
                        arrBrands = getBrands("ORDEN")
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
                        
                        arrBrands = getBrands("ORDEN")
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
                        
                        
                        arrClients = getClients("ORDEN")
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
        
        <div style="position:absolute;bottom:0px;width:100%;height:50px;background-color:white;border-top:1px solid black;">
            <%if Request.Form("FILTER_REPORTTYPE") <> "" then %>
                <div style="position:absolute;bottom:10px;left:20px;">
                    <input class=button type=button value="<%=IDM_FilterClose %>" onclick="closeFilter();" />
                </div>
            <%end if %>
            <div style="position:absolute;bottom:10px;right:20px;">
                <input type=button class=button value="<%=IDM_FilterPrevious %>" onclick="cambiarPest(Pest02, TBL_FILTER_QUICK);" />
                <input type=button class=button value="<%=IDM_FilterNext %>" onclick="cambiarPest(Pest10, TBL_FILTER_ORG);" />

                <%if Application("BotonExportExcel") = "YES" then %>
                    <input class=button type=button value="<%=IDM_ExportExcel %>" onclick="applyFilter(true, true);" />
                <%end if %>
                <input class=button type=button value="<%=IDM_FilterApply %>" onclick="applyFilter(true, false);" />
            </div>
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
                <td>
                    <input style="width:30px;height:30px;" type="checkbox" name="FILTER_LASTYEAR" <%=sCheckedLastYear %> />
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

        <div style="position:absolute;bottom:0px;width:100%;height:50px;background-color:white;border-top:1px solid black;">
            <%if Request.Form("FILTER_REPORTTYPE") <> "" then %>
                <div style="position:absolute;bottom:10px;left:20px;">
                    <input class=button type=button value="<%=IDM_FilterClose %>" onclick="closeFilter();" />
                </div>
            <%end if %>
            <div style="position:absolute;bottom:10px;right:20px;">
                <input type=button class=button value="<%=IDM_FilterPrevious %>" onclick="cambiarPest(Pest05, TBL_FILTER_CLIBRA);" />
                <input type=button class=button value="<%=IDM_FilterNext %>" onclick="cambiarPest(Pest20, TBL_FILTER_DATOS);" />
                <%if Application("BotonExportExcel") = "YES" then %>
                    <input class=button type=button value="<%=IDM_ExportExcel %>" onclick="applyFilter(true, true);" />
                <%end if %>
                <input class=button type=button value="<%=IDM_FilterApply %>" onclick="applyFilter(true, false);" />
            </div>
        </div>

    </div>
    
    <div id="TBL_FILTER_DATOS" style="display:none;">
        <table width="100%" >
            <tr>
                <%
                dim sChecked: sChecked = ""
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
                if Request.Form("FILTER_TOTALSHOPS")<>"" then
                    sChecked = "checked"
                elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ShowTotalShops")="YES" then
                    sChecked = "checked"
                end if
                %>
                <td width=120 class="fieldheader"><%=IDM_NTiendasTOTAL %></td>
                <td><input style="width:30px;height:30px;" type=checkbox name="FILTER_TOTALSHOPS" <%=sChecked%> /></td>
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
            <tr>
                <%
                sChecked = ""
                if Request.Form("FILTER_SHOWKPIQUALITY")<>"" then
                    sChecked = "checked"
                elseif Request.Form("FILTER_REPORTTYPE")="" AND Application("Default_ShowKPIQuality")="YES" then
                    sChecked = "checked"
                end if
                %>
                <td width=120 class="fieldheader"><%=IDM_KPIQuality %></td>
                <td><input style="width:30px;height:30px;" type=checkbox name="FILTER_SHOWKPIQUALITY" <%=sChecked%> /></td>
            </tr>
            <%if IsInputQuality() then %>
                <tr>
                    <%
                    sChecked = ""
                    if Request.Form("FILTER_SHOWQUALITY")<>"" then
                        sChecked = "checked"
                    end if
                    %>
                    <td width=120 class="fieldheader">Quality</td>
                    <td><input style="width:30px;height:30px;" type=checkbox name="FILTER_SHOWQUALITY" <%=sChecked%> /></td>
                </tr>
            <%end if %>
        </table>
        
        
        <div style="position:absolute;bottom:0px;width:100%;height:50px;background-color:white;border-top:1px solid black;">
            <%if Request.Form("FILTER_REPORTTYPE") <> "" then %>
                <div style="position:absolute;bottom:10px;left:20px;">
                    <input class=button type=button value="<%=IDM_FilterClose %>" onclick="closeFilter();" />
                </div>
            <%end if %>
            <div style="position:absolute;bottom:10px;right:20px;">
                <input type=button class=button value="<%=IDM_FilterPrevious %>" onclick="cambiarPest(Pest10, TBL_FILTER_ORG);" />
                
                <%if Application("BotonExportExcel") = "YES" then %>
                    <input class=button type=button value="<%=IDM_ExportExcel %>" onclick="applyFilter(true, true);" />
                <%end if %>
                <input type=button class=button value="<%=IDM_FilterApply %>" onclick="applyFilter(true, false);" />
            </div>
        </div>

    </div>

    <%if isAdmin() then %>
        <div id="TBL_FILTER_EXTRAS" style="display:none;">
            <table width="100%" cellpadding="5px">
                <tr>
                    <td>
                        <select name="IDReportQuery">
                            <option value=""><%=IDM_BtnReportQuery_Option %></option>
                            <%
                            SQL = "SELECT * FROM ReportQuery ORDER BY Orden "
                            rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
                            while not rst.EOF
                                %><option value="<%=rst("ID") %>"><%=rst("Nombre") %></option><%
                                rst.MoveNext
                            wend
                            rst.Close
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td valign="top">
                        <input type="button" class="button" value="<%=IDM_BtnReportQuery %>" onclick="lanzarQuery();return false;" />
                    </td>
                    <td valign="top"><font class="font12"><%=IDM_BtnReportQueryTxt %></font></td>
                </tr>
            </table>
        </div>
    <%end if %>
    
</div>



<script language="javascript" type="text/javascript">
    <%if Request.Form("Quick_Tipo")<>"" then %>
        SelectorQuick(<%=Request.Form("Quick_Tipo") %>);
    <%end if %>
    toggleReportType();
    
    document.getElementById('MAIN_FILTER').style.left = (screen.width / 2) - <%=MainFilterWidth %>/2 ;
    
</script>