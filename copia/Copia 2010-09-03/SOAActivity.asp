<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
%>

<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/EventFunctions1.asp" -->

<!-- #include file = "RenderFunctions.asp" -->
<!-- #include file = "ClassInclude.asp" -->


<%
dim rst, rst2, sSelected
set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")

rst.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient

RecoverSQLConnection()

RecoverSession(true) 


%>
<!-- #include file = "include/Idioma.asp" -->

<%
dim ID: ID = Request("ID")
if ID = "" then ID = -1

dim IDClient: IDClient = Request("IDClient")
dim IDBrand: IDBrand = Request("IDBrand")
dim WYear: WYear = Request("WYear")
dim WMonth: WMonth = Request("WMonth")
dim WHalf: WHalf = Request("WHalf")
dim FILTER_YEAR: FILTER_YEAR = CInt(Request("FILTER_YEAR"))
dim FILTER_STARTMONTH: FILTER_STARTMONTH = CInt(Request("FILTER_STARTMONTH"))
dim FILTER_VIEWMONTHS: FILTER_VIEWMONTHS = CInt(Request("FILTER_VIEWMONTHS"))
dim FILTER_MULTIBRAND: FILTER_MULTIBRAND = Request("FILTER_MULTIBRAND")
dim FILTER_MULTICLIENT: FILTER_MULTICLIENT = Request("FILTER_MULTICLIENT")

dim DataModified: DataModified = Request("DataModified")

dim bottomNavigate: bottomNavigate = ""

Sub Save_click()
    
    ' ************************************************************************************
    ' USA LA VARIABLE act QUE ES LA ACTIVIDAD ACTUAL (YA SEA NUEVA O EN EDICIÓN)
    ' ************************************************************************************
    
    act.IDClient = IDClient
    act.IDBrand = IDBrand
    act.WYear = WYear
    act.WMonth = WMonth
    act.WHalf = WHalf
    act.Name = Request.Form("Name")
    
    act.Oferta = Request.Form("Oferta")
    act.IDRatio = Request.Form("IDRatio")
    act.Folleto = Request.Form("Folleto")
    act.Cabecera = Request.Form("Cabecera")
    act.NShops = Request.Form("NShops")
    'act.PercentComplaint = Request.Form("PercentComplaint")
    act.PercentComplaint = ""
    act.IDStatus = Request.Form("IDStatus")
    act.Adicional = Request.Form("Adicional")

    on error resume next
    saveActivity(act)
    if Err<>0 then
        bottomMessage = Err.Description
    end if
    on error goto 0
    
    
    'Recarga los datos en la variable 'act'
    set act = getActivity(act.ID)
    
    
    if Request.Form("CloseWindow") <> "" then
        %><script language="JavaScript">try{window.close();}catch(e){} try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} try{window.opener.applyFilter(false);}catch(e){}</script><%
    end if
    if Request.Form("Navigate") = "NAVIGATE_TO" then
        bottomNavigate = "navigateTo(" & Request.Form("NavigateToYear") & ", " & Request.Form("NavigateToMonth") & ", " & Request.Form("NavigateToHalf") & ");"
    elseif Request.Form("Navigate") = "NAVIGATE_TO_CLIENT" then
        bottomNavigate = "navigateToClient('" & Request.Form("NavigateToClient") & "');"
    elseif Request.Form("Navigate") = "NAVIGATE_TO_BRAND" then
        bottomNavigate = "navigateToBrand('" & Request.Form("NavigateToBrand") & "');"
    end if
    
    DataModified = "1"
    
End Sub


Sub Delete_click(delID)
    
    on error resume next
    deleteActivity(delID)

    if Err<>0 then
        bottomMessage = "Error deleting activity"
    else
        ' Cierra la ventana
        
        set act = new Activity
        
        ID = -1
        
        DataModified = "1"

        if FALSE then
        %><script language="JavaScript">try{window.close();}catch(e){} try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} try{window.opener.applyFilter(false);}catch(e){}</script><%
        end if
    end if
    
End Sub


Sub Copy_click()
    
    session("Clipboard") = "1"
    
    session("Clipboard_Name") = Request.Form("Name")
    session("Clipboard_Oferta") = Request.Form("Oferta")
    session("Clipboard_IDRatio") = Request.Form("IDRatio")
    session("Clipboard_Folleto") = Request.Form("Folleto")
    session("Clipboard_Cabecera") = Request.Form("Cabecera")
    session("Clipboard_NShops") = Request.Form("NShops")
    ''' session("Clipboard_PercentComplaint") = Request.Form("PercentComplaint")
    session("Clipboard_IDStatus") = Request.Form("IDStatus")
    session("Clipboard_Adicional") = Request.Form("Adicional")
    
End Sub

Sub ClearClipboard_click()
    
    session("Clipboard") = ""

    session("Clipboard_Name") = ""
    session("Clipboard_Oferta") = ""
    session("Clipboard_IDRatio") = ""
    session("Clipboard_Folleto") = ""
    session("Clipboard_Cabecera") = ""
    session("Clipboard_NShops") = ""
    ''' session("Clipboard_PercentComplaint") = ""
    session("Clipboard_IDStatus") = ""
    session("Clipboard_Adicional") = ""
    
End Sub

Sub SaveAndCopy_click()
    
    Save_click()
    
    Copy_click()
    
End Sub

Sub Paste_click()
    
    ' Ya existe una actividad --> Copia los datos del portapapeles y guarda
    if session("Clipboard")<>"" then

        act.IDClient = IDClient
        act.IDBrand = IDBrand
        act.WYear = WYear
        act.WMonth = WMonth
        act.WHalf = WHalf

        ' Hay que inicializar si están vacíos
        if isNull(act.NShops) then act.NShops = ""

        ''' if isNull(act.PercentComplaint) then act.PercentComplaint = ""
        act.PercentComplaint = ""

        if session("Clipboard_Name")<>"" then act.Name = session("Clipboard_Name")
        if session("Clipboard_Oferta")<>"" then act.Oferta = session("Clipboard_Oferta")
        if session("Clipboard_IDRatio")<>"" then act.IDRatio = session("Clipboard_IDRatio")
        if session("Clipboard_Folleto")<>"" then act.Folleto = session("Clipboard_Folleto")
        if session("Clipboard_Cabecera")<>"" then act.Cabecera = session("Clipboard_Cabecera")
        if session("Clipboard_NShops")<>"" then act.NShops = session("Clipboard_NShops")
        '''' if session("Clipboard_PercentComplaint")<>"" then act.PercentComplaint = session("Clipboard_PercentComplaint")
        if session("Clipboard_IDStatus")<>"" then act.IDStatus = session("Clipboard_IDStatus")
        if session("Clipboard_Adicional")<>"" then act.Adicional = session("Clipboard_Adicional")
        
        on error resume next
        saveActivity(act)
        if Err<>0 then
            bottomMessage = Err.Description
        end if
        on error goto 0
        
        
        'Recarga los datos en la variable 'act'
        set act = getActivity(act.ID)
        
        DataModified = "1"
    end if
    
End Sub


' ************************************************************************************
' Tiene que crear la variable 'act' antes de realizar ninguna acción
' ************************************************************************************
dim act
if CInt(ID) > -1 then
    on error resume next
    set act = getActivity(ID)
    if Err<>0 then
        %>
        <br /><br /><br /><br /><br /><br />
        <table align=center width=300 style="border:1 solid gray;"><tr height=200><td align=center><font style="font-family:Arial;"><%=Err.Description %></font><br /><br /><input type=button value="Close" onclick="try{window.close();}catch(e){}try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} try{window.opener.applyFilter(false);}catch(e){} " /></td></tr></table>
        <%
        Response.End
    end if
    on error goto 0
else
    ' Es un elemento nuevo
    set act = getActivityFromDate(IDClient, IDBrand, WYear, WMonth, WHalf)
end if


' ************************************************************************************
' Ejecución de los eventos
' ************************************************************************************
Select Case EventObject
	case "Save" Save_click()
	case "Delete" Delete_click(EventParam1)
	case "Copy" Copy_click()
	case "Paste" Paste_click()
	case "SaveAndCopy" SaveAndCopy_click()
	case "ClearClipboard" ClearClipboard_click()
End Select


dim aCli, aBra
set aCli = getClient(IDClient)
set aBra = getBrand(IDBrand)


%>

<HTML>
<HEAD>
    <TITLE><%=IDM_Activity %></TITLE>
    <LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
    <script language=javascript>
        var dataModified = false;

        function closeWindow()
        {
            if (dataModified)
            {
                <%if Application("AutoSaveActivity") = "NO" then %>
                if (confirm('<%=IDM_JS_DatosModificadosGuardar %>')){
                <%end if %>
                    thisForm.CloseWindow.value = '1';
                    Save();
                    return false;
                <%if Application("AutoSaveActivity") = "NO" then %>
                }
                <%end if %>
            }
            
          <%if DataModified<>"" then %> 
            try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} 
            try{window.opener.applyFilter(false);}catch(e){} 
          <%end if %> 
          
          try{window.close();}catch(e){}
        }
        function cancelWindow()
        {
            if (dataModified)
            {
                if (confirm('<%=IDM_JS_DatosModificadosGuardar %>')){
                    thisForm.CloseWindow.value = '1';
                    Save();
                    return false;
                }
            }
            
          <%if DataModified<>"" then %> 
            try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} 
            try{window.opener.applyFilter(false);}catch(e){} 
          <%end if %> 
          
          try{window.close();}catch(e){}
        }
        function changeMade()
        {
            dataModified = true;
            thisForm.BTN_Cancelar.style.display = '';
        }
        function Save()
        {
            if (thisForm.Oferta.value=='' && thisForm.Folleto.value=='' && thisForm.Cabecera.value=='' && thisForm.Adicional.value=='' ) {
                alert('<%=IDM_JS_RellenarAlgunCampo %>'); 
                return false;
            }
            
            _fireEvent('Save', '', '');
        }
        function Delete(id)
        {
            _fireConfirm('Delete', id, '', '');
        }
        function Copy()
        {
            if (thisForm.Oferta.value=='' && thisForm.Folleto.value=='' && thisForm.Cabecera.value=='' && thisForm.Adicional.value=='' ) {
                alert('<%=IDM_JS_RellenarAlgunCampo %>'); 
                return false;
            }

            if (dataModified)
            {
                <%if Application("AutoSaveActivity") = "NO" then %>
                if (confirm('<%=IDM_JS_DatosModificadosGuardarCambiar %>')){
                <%end if %>
                    _fireEvent('SaveAndCopy', '', '');
                    return false;
                <%if Application("AutoSaveActivity") = "NO" then %>
                }
                <%end if %>
            }


            _fireEvent('Copy', '', '');
        }
        function Paste()
        {
            _fireEvent('Paste', '', '');
        }
        function ClearClipboard()
        {
            _fireEvent('ClearClipboard', '', '');
        }
        function navigateTo(ToYear, ToMonth, ToHalf)
        {
            if (dataModified)
            {
                <%if Application("AutoSaveActivity") = "NO" then %>
                if (confirm('<%=IDM_JS_DatosModificadosGuardarCambiar %>')){
                <%end if %>
                    thisForm.NavigateToYear.value = ToYear;
                    thisForm.NavigateToMonth.value = ToMonth;
                    thisForm.NavigateToHalf.value = ToHalf;
                    thisForm.Navigate.value = 'NAVIGATE_TO';
                    Save();
                    return false;
                <%if Application("AutoSaveActivity") = "NO" then %>
                }
                <%end if %>
            }
            
            thisForm.ID.value = '';
            thisForm.WYear.value = ToYear;
            thisForm.WMonth.value = ToMonth;
            thisForm.WHalf.value = ToHalf;
            thisForm.action = 'SOAActivity.asp';
            thisForm.submit();
            
            //location.href='SOAActivity.asp?IDClient=<%=IDClient%>&IDBrand=<%=IDBrand%>&WYear=' + ToYear + '&WMonth=' + ToMonth + '&WHalf=' + ToHalf + '&DataModified=<%=DataModified %>&FILTER_YEAR=<%=FILTER_YEAR %>&FILTER_STARTMONTH=<%=FILTER_STARTMONTH %>&FILTER_VIEWMONTHS=<%=FILTER_VIEWMONTHS %>';
        }
        
        function navigateToClient(idClient)
        {
            if (dataModified)
            {
                <%if Application("AutoSaveActivity") = "NO" then %>
                if (confirm('<%=IDM_JS_DatosModificadosGuardarCambiar %>')){
                <%end if %>
                    thisForm.NavigateToClient.value = idClient;
                    thisForm.Navigate.value = 'NAVIGATE_TO_CLIENT';
                    Save();
                    return false;
                <%if Application("AutoSaveActivity") = "NO" then %>
                }
                <%end if %>
            }
            
            thisForm.ID.value = '';
            thisForm.IDClient.value = idClient;
            thisForm.action = 'SOAActivity.asp';
            thisForm.submit();

            //location.href='SOAActivity.asp?IDClient=' + idClient + '&IDBrand=<%=IDBrand%>&WYear=<%=WYear %>&WMonth=<%=WMonth %>&WHalf=<%=WHalf %>&DataModified=<%=DataModified %>&FILTER_YEAR=<%=FILTER_YEAR %>&FILTER_STARTMONTH=<%=FILTER_STARTMONTH %>&FILTER_VIEWMONTHS=<%=FILTER_VIEWMONTHS %>&FILTER_MULTIBRAND=<%=FILTER_MULTIBRAND %>&FILTER_MULTICLIENT=<%=FILTER_MULTICLIENT %>';
        }

        function navigateToBrand(idBrand)
        {
            if (dataModified)
            {
                <%if Application("AutoSaveActivity") = "NO" then %>
                if (confirm('<%=IDM_JS_DatosModificadosGuardarCambiar %>')){
                <%end if %>
                    thisForm.NavigateToBrand.value = idBrand;
                    thisForm.Navigate.value = 'NAVIGATE_TO_BRAND';
                    Save();
                    return false;
                <%if Application("AutoSaveActivity") = "NO" then %>
                }
                <%end if %>
            }
            
            thisForm.ID.value = '';
            thisForm.IDBrand.value = idBrand;
            thisForm.action = 'SOAActivity.asp';
            thisForm.submit();

            //location.href='SOAActivity.asp?IDClient=' + idClient + '&IDBrand=<%=IDBrand%>&WYear=<%=WYear %>&WMonth=<%=WMonth %>&WHalf=<%=WHalf %>&DataModified=<%=DataModified %>&FILTER_YEAR=<%=FILTER_YEAR %>&FILTER_STARTMONTH=<%=FILTER_STARTMONTH %>&FILTER_VIEWMONTHS=<%=FILTER_VIEWMONTHS %>&FILTER_MULTIBRAND=<%=FILTER_MULTIBRAND %>&FILTER_MULTICLIENT=<%=FILTER_MULTICLIENT %>';
        }        
		function ControlOnlyNumbers(e){
			var keycode;
			if (window.event) keycode = window.event.keyCode;
			else if (e) keycode = e.which;
			else return false;
			
			if ((keycode>=48 && keycode<=57) ){
				return true
			}else{
				e.which = 0;
				return false
			}
		}
		
		function toggleDropDown(view)
		{
		    var sDisplay = "";
            if (view){ sDisplay = ""; } else { sDisplay = "none"; }
            var x = document.getElementsByTagName("select");

            for (i = 0; i < x.length; i++) {
                x[i].style.display = sDisplay;
            }
		}
		
		function DIVClient_MouseOver()
		{
		    toggleDropDown(false);
		    DIV_ClientNavigator.style.display='';
		}
		function DIVClient_MouseOut()
		{
		    toggleDropDown(true);
            DIV_ClientNavigator.style.display='none';
		}
		function DIVBrand_MouseOver()
		{
		    toggleDropDown(false);
		    DIV_BrandNavigator.style.display='';
		}
		function DIVBrand_MouseOut()
		{
		    toggleDropDown(true);
            DIV_BrandNavigator.style.display='none';
		}
    </script>
</HEAD>

<BODY leftmargin=0 topmargin=0 >

 
<FORM action="" method="post" name="thisForm">
    
<%=getClientNavigation() %>
<%=getBrandNavigation() %>
    
    <table style="width:100%;">
        <tr>
            <td align=left><%if aCli.ImageFileNameH <> "" then %><img height="50" src="images/Clients/<%=aCli.ImageFileNameH %>" /><%else %><font class="font12"><%=aCli.Name %></font><%end if %></td>
            <td align=right><%if aBra.ImageFileNameH <> "" then %><img height="50" src="images/Brands/<%=aBra.ImageFileNameH %>" /><%else %><font class="font12"><%=aBra.Name %></font><%end if %></td>
        </tr>
        <tr>
            <td align=left><a href="" onclick="DIVClient_MouseOver();return false;" ><font class="font10"><%=IDM_ActivityChangeClient %></font></a></td>
            <td align=right><a href="" onclick="DIVBrand_MouseOver();return false;" ><font class="font10"><%=IDM_ActivityChangeBrand %></font></a></td>
        </tr>
    </table>

    
    <%=PintarCalendarioNavegacionActividad(IDClient, IDBrand, FILTER_YEAR, FILTER_STARTMONTH, FILTER_VIEWMONTHS, WYear, WMonth, WHalf) %>

    <table style="width:100%;height:40px;background-image:url('images/Grad5.gif'); ">
        <tr>
            <td align="left" width=140>
                <%if CInt(act.ID) <> -1 then %>
                    <input type=button class="button" value="<%=IDM_Delete %>" style="width:65px;" onclick="Delete(<%=act.ID %>);" />
                <%end if %>
                <input ID="BTN_Cancelar" type=button class="button" value="<%=IDM_Cancel %>" style="display:none;width:65px;" onClick="cancelWindow();" />
            </td>
            <td align=center>
                <input type=button class="button" title="<%=IDM_CopyAlt %>" value="<%=IDM_Copy %>" style="width:65px;" onclick="Copy();" />
                <%if session("Clipboard") <> "" then %>
                    <input type=button class="button" title="<%=IDM_PasteAlt %>" value="<%=IDM_Paste %>" style="width:65px;" onclick="Paste();" />
                    <input type=button class="button" title="<%=IDM_ClearAlt %>" value="<%=IDM_Clear %>" style="width:65px;" onclick="ClearClipboard();" />
                <%end if %>
            </td>
            <td align="right" width=140>
                <input type=button class="button" value="<%=IDM_Save %>" style="width:65px;" onclick="Save();" />
                <input type=button class="button" value="<%=IDM_Close %>" style="width:65px;" onClick="closeWindow();" />
            </td>
        </tr>
    </table>

    <table style="width:100%;height:30px;">
        <tr>
            <td valign=top width=150 class="fieldheader" ><%=IDM_Oferta %></td>
            <td>
                <textarea onchange="changeMade();" name="Oferta" class="textfield" style="width:100%;height:60px;"><%=act.Oferta %></textarea>
            </td>
        </tr>

        <tr>
            <td class="fieldheader"><%=IDM_Ratio %></td>
            <td>
                <select name="IDRatio" style="width:100%;" class="textfield" onchange="changeMade();">
                <%
                dim rats, r
                rats = getActivityRatios(Idioma)
                for each r in rats
                    sSelected = ""
                    if act.IDRatio <> "" then
                        if CInt(act.IDRatio) = r.ID then
                            sSelected = "selected"
                        end if
                    end if
                    %><option value="<%=r.ID %>" <%=sSelected %>><%=r.Name %></option><%
                next
                %>
                </select>
            </td>
        </tr>

        <tr>
            <td valign=top class="fieldheader" ><%=IDM_Folleto %></td>
            <td>
                <textarea onchange="changeMade();" name="Folleto" class="textfield" style="width:100%;height:60px;"><%=act.Folleto %></textarea>
            </td>
        </tr>
        
        <tr>
            <td valign=top class="fieldheader" ><%=IDM_Cabecera %></td>
            <td>
                <textarea onchange="changeMade();" name="Cabecera" class="textfield" style="width:100%;height:60px;"><%=act.Cabecera %></textarea>
            </td>
        </tr>

    </table>
    
    <table style="width:100%;height:30px;">
        <tr>
            <td width=150 class="fieldheader"><%=IDM_NTiendas %></td>
            <td>
                <input name="NShops" value="<%=act.NShops %>" class="textfield" type="text" onKeyPress="return ControlOnlyNumbers(event)" onchange="changeMade();" style="width:100%;" />
            </td>
        </tr>
        
        <%''' REAL DATA ''' %>
        <tr>
            <td width=150 class="fieldheader"><%=IDM_NTiendasReal %></td>
            <td>
                <input readonly name="" value="<%=getRealData(WYear, WMonth, WHalf, IDClient, IDBrand).NShops %>" class="textfieldreadonly" type="text" onKeyPress="return ControlOnlyNumbers(event)" onchange="changeMade();" style="width:100%;" />
            </td>
            <td width=150 class="fieldheader"><%=IDM_PercentComplaint %></td>
            <td>
                <input readonly name="" value="<%=getRealData(WYear, WMonth, WHalf, IDClient, IDBrand).PercentComplaint %>" class="textfieldreadonly" type="text" onKeyPress="return ControlOnlyNumbers(event)" onchange="changeMade();" style="width:100%;" />
            </td>
        </tr>
    </table>
    
    <table style="width:100%;height:30px;">
        <tr>
            <td width=150 class="fieldheader"><%=IDM_Status %></td>
            <td>
                <%
                dim arrStatus
                dim s, iSt
                arrStatus = getActivityStatuses()
                iSt = 1
                for each s in arrStatus
                    sSelected = ""
                    if act.IDStatus<>-1 then
                        if CInt(act.IDStatus) = s.ID then
                            sSelected = "checked"
                        end if
                    end if
                    %><input style="width:25px;height:25px;" type="radio" name="IDStatus" id="ST_<%=s.ID %>" value="<%=s.ID %>" <%=sSelected %> onchange="changeMade();" /><font class="font15"><%=s.Name %>&nbsp;&nbsp;&nbsp;</font><%
                    
                    if iSt mod 4 = 0 then
                        %><br /><%
                    end if
                    
                    iSt = iSt + 1
                next
                %>
            </td>
        </tr>

        <tr>
            <td valign=top class="fieldheader" ><%=IDM_Adicional %></td>
            <td>
                <textarea onchange="changeMade();" name="Adicional" class="textfield" style="width:100%;height:60px;"><%=act.Adicional %></textarea>
            </td>
        </tr>

        <%if CInt(act.ID) <> -1 then %>
            <tr height=20><td></td></tr>
            <tr>
                <td valign=top class="fieldheader" style="border-top:1 solid silver;"><%=IDM_LastUpdatedBy %></td>
                <td style="border-top:1 solid silver;"><font class=font12>
                    <%=act.LastUpdatedBy %>
                    &nbsp;-&nbsp;
                    <%=act.LastUpdatedDate %>
                    </font>
                </td>
            </tr>
        <%end if %>
            
    </table>
    
    
    
    <input type=hidden name="ID" value="<%=act.ID %>" />
    <input type=hidden name="IDClient" value="<%=IDClient %>" />
    <input type=hidden name="IDBrand" value="<%=IDBrand %>" />
    <input type=hidden name="WYear" value="<%=WYear %>" />
    <input type=hidden name="WMonth" value="<%=WMonth %>" />
    <input type=hidden name="WHalf" value="<%=WHalf %>" />
    <input type=hidden name="FILTER_YEAR" value="<%=FILTER_YEAR %>" />
    <input type=hidden name="FILTER_STARTMONTH" value="<%=FILTER_STARTMONTH %>" />
    <input type=hidden name="FILTER_VIEWMONTHS" value="<%=FILTER_VIEWMONTHS %>" />
    <input type=hidden name="FILTER_MULTIBRAND" value="<%=FILTER_MULTIBRAND %>" />
    <input type=hidden name="FILTER_MULTICLIENT" value="<%=FILTER_MULTICLIENT %>" />

    <input type=hidden name="DataModified" value="<%=DataModified %>" />

    <input type=hidden name="Navigate" value="" />
    <input type=hidden name="NavigateToYear" value="" />
    <input type=hidden name="NavigateToMonth" value="" />
    <input type=hidden name="NavigateToHalf" value="" />
    <input type=hidden name="NavigateToClient" value="" />
    <input type=hidden name="NavigateToBrand" value="" />
    <input type=hidden name="CloseWindow" value="" />

    <!-- #include file = "include/EventFunctions2.asp" -->

</FORM>


<!-- #include file = "include/pageBottom.asp" -->

<script language=javascript>
    thisForm.Oferta.focus();
    
    <%=bottomNavigate %>
</script>

</BODY>

</HTML>