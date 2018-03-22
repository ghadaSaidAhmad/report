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
dim rst, rst2, SQL, sSelected
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

dim WYear: WYear = Request("WYear")
dim WMonth: WMonth = Request("WMonth")
dim WHalf: WHalf = Request("WHalf")

dim FILTER_YEAR: FILTER_YEAR = CInt(Request("FILTER_YEAR"))
dim FILTER_STARTMONTH: FILTER_STARTMONTH = CInt(Request("FILTER_STARTMONTH"))
dim FILTER_VIEWMONTHS: FILTER_VIEWMONTHS = CInt(Request("FILTER_VIEWMONTHS"))

dim DataModified: DataModified = Request("DataModified")

dim bottomNavigate: bottomNavigate = ""

Sub Save_click()
    
    ' ************************************************************************************
    ' USA LA VARIABLE gthm QUE ES EL TEMA ACTUAL (YA SEA NUEVA O EN EDICIÓN)
    ' ************************************************************************************

    gthm.IDClient = IDClient
    gthm.WYear = WYear
    gthm.WMonth = WMonth
    gthm.WHalf = WHalf
    gthm.Name = Request.Form("Name")
    
    gthm.IDTheme = Request.Form("IDTheme")
    
    on error resume next
    saveGeneralTheme(gthm)
    if Err<>0 then
        bottomMessage = Err.Description
    end if
    on error goto 0
    
    set gthm = getGeneralTheme(gthm.ID)
    
    if Request.Form("CloseWindow") <> "" then
        %><script language="JavaScript">try{window.close();}catch(e){} try{window.opener.applyFilter(false);}catch(e){}</script><%
    end if
    if Request.Form("Navigate") = "NAVIGATE_TO" then
        bottomNavigate = "navigateTo(" & Request.Form("NavigateToYear") & ", " & Request.Form("NavigateToMonth") & ", " & Request.Form("NavigateToHalf") & ");"
    end if

    DataModified = "1"

End Sub

Sub Delete_click(delID)
    
    on error resume next
    deleteGeneralTheme(delID)

    if Err<>0 then
        bottomMessage = "Error deleting General Theme"
    else
        ' Cierra la ventana
        
        set gthm = new GeneralTheme
        
        DataModified = "1"

        if FALSE then
        %><script language="JavaScript">try{window.close();}catch(e){}try{window.opener.applyFilter(false);}catch(e){}</script><%
        end if
    end if
    
End Sub


' ************************************************************************************
' Tiene que crear la variable 'gthm' antes de realizar ninguna acción
' ************************************************************************************
dim gthm
if CInt(ID) > -1 then
    on error resume next
    set gthm = getGeneralTheme(ID)
    if Err<>0 then
        %>
        <br /><br /><br /><br /><br /><br />
        <table align=center width=300 style="border:1 solid gray;"><tr height=200><td align=center><font style="font-family:Arial;"><%=Err.Description %></font><br /><br /><input type=button value="Close" onclick="try{window.close();}catch(e){}try{window.opener.applyFilter(false);}catch(e){} " /></td></tr></table>
        <%
        Response.End
    end if
    on error goto 0
else
    ' Es un elemento nuevo
    set gthm = getGeneralThemeFromDate(IDClient, WYear, WMonth, WHalf)
end if


' ************************************************************************************
' Ejecución de los eventos
' ************************************************************************************
Select Case EventObject
	case "Save" Save_click()
	case "Delete" Delete_click(EventParam1)
End Select


dim aCli
set aCli = getClient(IDClient)

%>
<HTML>
<HEAD>
    <TITLE><%=IDM_GeneralTheme %></TITLE>
    <LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
    <script language="javascript">
        
        var dataModified = false;
        var themeList;

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
            
          <%if Request("PageReloaded")<>"" OR Request("DataModified")<>"" then %> 
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
            
          <%if Request("PageReloaded")<>"" OR Request("DataModified")<>"" then %> 
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
            if ((thisForm.Name.value=='') && (thisForm.IDTheme.selectedIndex==0 || thisForm.IDTheme.selectedIndex==-1))
            {
                alert('<%=IDM_JS_SeleccioneTematica %>');
                return false;
            }

            _fireEvent('Save', '', '');
        }
        function Delete(id)
        {
            _fireConfirm('Delete', id, '', '');
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
            
            location.href='SOAGeneralTheme.asp?IDClient=<%=IDClient%>&WYear=' + ToYear + '&WMonth=' + ToMonth + '&WHalf=' + ToHalf + '&DataModified=<%=DataModified %>&FILTER_YEAR=<%=FILTER_YEAR %>&FILTER_STARTMONTH=<%=FILTER_STARTMONTH %>&FILTER_VIEWMONTHS=<%=FILTER_VIEWMONTHS %>';
        }
        function cambioTematica()
        {
            if (thisForm.IDTheme.selectedIndex==0){
                BTN_ModifTheme.style.display = 'none';
            }else{
                BTN_ModifTheme.style.display = '';
            }

            if (thisForm.IDTheme.selectedIndex>0)
            {
                if (themeList.Themes[thisForm.IDTheme.selectedIndex-1].ImageFileName != ''){
                    document.images["IMG_Theme"].src = 'images/Themes/' + themeList.Themes[thisForm.IDTheme.selectedIndex-1].ImageFileName;
                    document.images["IMG_Theme"].style.display='';
                }else{
                    document.images["IMG_Theme"].style.display='none';
                }
            }else{
                document.images["IMG_Theme"].style.display='none';
            }
        }

        function cargarTematicas(defaultID)
        {
            var sDat
            sDat = '<%=IDClient %>';
            <%if gthm.IDTheme<>"" then %>
                sDat += ';<%=gthm.IDTheme %>';
            <%end if %>
			ajaxres = ajaxReq('ListaTematicas', sDat);

			themeList = eval('('+ ajaxres +')');
			
			
			//Borra la lista de Themes actual
			thisForm.IDTheme.options.length=0;
			
			// Incluye en la lista el primer elemento vacío
			var objOption = document.createElement("option");
			objOption.text = '' ;
			objOption.value = '-1';
			thisForm.IDTheme.options.add(objOption)
			
			
			//Rellena la lista con los nuevos Themes
			for (i=0;i<themeList.Themes.length;i++){
			    var bSelected = false;
				objOption = document.createElement("option");
				objOption.text = themeList.Themes[i].Name ;
				objOption.value = themeList.Themes[i].id;
				thisForm.IDTheme.options.add(objOption)
			}
			
			if (defaultID!=''){
    			thisForm.IDTheme.value = defaultID;
    	    }else{
    	        thisForm.IDTheme.selectedIndex = 0;
    	    }
    	    
        }
        function nuevaTematicaCreada(id)
        {
            cargarTematicas(id);
            if (thisForm.IDTheme.value != id){
                thisForm.IDTheme.value = id;
            }else{
                cambioTematica();
            }
            
            changeMade();
        }
        function tematicaBorrada(id)
        {
            cargarTematicas(id);
            thisForm.IDTheme.value = id;
            changeMade();
        }
        function tematicaImagenBorrada()
        {
            document.images["IMG_Theme"].src = '';
            document.images["IMG_Theme"].style.display='none';
        }
    </script>
</HEAD>

<BODY leftmargin=0 topmargin=0 >

<FORM action="" method="post" name="thisForm">
    
    <table style="width:100%;">
        <tr>
            <td align=left><%if aCli.ImageFileNameH <> "" then %><img height="50" src="images/Clients/<%=aCli.ImageFileNameH %>" /><%else %><font class="font12"><%=aCli.Name %></font><%end if %></td>
            <td><%=PintarCalendarioNavegacionGeneralTheme(IDClient, FILTER_YEAR, FILTER_STARTMONTH, FILTER_VIEWMONTHS, WYear, WMonth, WHalf) %></td>
        </tr>
    </table>
    
    
    <table style="width:100%;height:10px;background-color:Black;"><tr><td></td></tr></table>
    <table style="width:100%;height:40px;background-image:url('images/Grad5.gif'); ">
        <tr>
            <td align=left>
                <%if CInt(gthm.ID) <> -1 then %>
                    <input type=button class="button" value="<%=IDM_Delete %>" style="width:55px;" onclick="Delete(<%=gthm.ID %>);" />
                <%end if %>
                <input ID="BTN_Cancelar" type=button class="button" value="<%=IDM_Cancel %>" style="display:none;width:65px;" onClick="cancelWindow();" />
            </td>
            <td align="right">
                <input type=button class="button" value="<%=IDM_Save %>" style="width:55px;" onclick="Save();" />
                <input type=button class="button" value="<%=IDM_Close %>" style="width5:55px;" onClick="closeWindow();" />
            </td>
        </tr>
    </table>
    
    <table style="width:100%;height:30px;">
        <tr>
            <td width=150 valign=top class="fieldheader"><%=IDM_Tematica %></td>
            <td valign=top >
                <select name="IDTheme" style="width:100%;" class="textfield" onchange="changeMade(); cambioTematica();"></select>
                <img ID="IMG_Theme" src="" style="width:<%=Application("ThemeImageWidth") %>px;" />
            </td>
            <td width=90 align=right valign=top>
                <a title="<%=IDM_ModificarTematica %>" ID="BTN_ModifTheme" href="" onclick="window.open('SOAAddTheme.asp?ID=' + thisForm.IDTheme.value + '&IDClient=<%=IDClient %>','ADMTHM','width=500,height=300,top=200,left=250,scrollbars');return false;"><img onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/edit30.png" style="border:2px solid white;" /></a>
                <a title="<%=IDM_NuevaTematica %>" href="" onclick="window.open('SOAAddTheme.asp?ID=-1&IDClient=<%=IDClient %>','ADMTHM','width=500,height=300,top=200,left=250,scrollbars');return false;"><img onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/add.png" style="border:2px solid white;" /></a>
            </td>
        </tr>

    </table>
    
    <table style="width:100%;height:30px;">
        <tr>
            <td valign=top width=150 class="fieldheader" ><%=IDM_ExtraInfo %></td>
            <td>
                <textarea onchange="changeMade();" name="Name" class="textfield" style="width:100%;height:60px;"><%=gthm.Name %></textarea>
            </td>
        </tr>

        <%if CInt(gthm.ID) <> -1 then %>
            <tr height=20><td></td></tr>
            <tr>
                <td valign=top width=150 class="fieldheader" style="border-top:1 solid silver;"><%=IDM_LastUpdatedBy %></td>
                <td style="border-top:1 solid silver;"><font class=font12>
                    <%=gthm.LastUpdatedBy %>
                    &nbsp;-&nbsp;
                    <%=gthm.LastUpdatedDate %>
                    </font>
                </td>
            </tr>
        <%end if %>
        
    </table>
    
    
    
    <input type=hidden name="ID" value="<%=gthm.ID %>" />
    <input type=hidden name="IDClient" value="<%=IDClient %>" />
    <input type=hidden name="WYear" value="<%=WYear %>" />
    <input type=hidden name="WMonth" value="<%=WMonth %>" />
    <input type=hidden name="WHalf" value="<%=WHalf %>" />
    <input type=hidden name="FILTER_YEAR" value="<%=FILTER_YEAR %>" />
    <input type=hidden name="FILTER_STARTMONTH" value="<%=FILTER_STARTMONTH %>" />
    <input type=hidden name="FILTER_VIEWMONTHS" value="<%=FILTER_VIEWMONTHS %>" />
    
    <input type=hidden name="DataModified" value="<%=DataModified %>" />

    <input type=hidden name="Navigate" value="" />
    <input type=hidden name="NavigateToYear" value="" />
    <input type=hidden name="NavigateToMonth" value="" />
    <input type=hidden name="NavigateToHalf" value="" />
    <input type=hidden name="CloseWindow" value="" />

    <!-- #include file = "include/EventFunctions2.asp" -->

</FORM>


<!-- #include file = "include/pageBottom.asp" -->

<script language=javascript>
    cargarTematicas('<%=gthm.IDTheme %>');
    cambioTematica();

    <%=bottomNavigate %>
</script>

</BODY>

</HTML>