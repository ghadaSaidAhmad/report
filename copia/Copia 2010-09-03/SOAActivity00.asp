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

dim IDClient: IDClient = Request("IDClient")
dim IDBrand: IDBrand = Request("IDBrand")
dim WYear: WYear = Request("WYear")
dim WMonth: WMonth = Request("WMonth")
dim WHalf: WHalf = Request("WHalf")
dim IDType: IDType = Request("IDType")

Sub Save_click()
    
    dim act
    set act = new Activity00
    
    act.ID = CInt(Request("ID"))
    act.IDClient = IDClient
    act.IDBrand = IDBrand
    act.WYear = WYear
    act.WMonth = WMonth
    act.WHalf = WHalf
    act.IDType = IDType
    act.Name = Request.Form("Name")
    
    act.IDTheme = Request.Form("IDTheme")
    act.IDRatio = Request.Form("IDRatio")
    
    on error resume next
    saveActivity00(act)
    if Err<>0 then
        bottomMessage = Err.Description
    else
        'Si era nuevo, ahora tiene un ID

        ID = act.ID

    end if
    on error goto 0
    
    if Request.Form("CloseWindow") <> "" then
        %><script language="JavaScript">try{window.close();}catch(e){} try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} try{window.opener.applyFilter(false);}catch(e){}</script><%
    end if
    
End Sub

Sub Delete_click(delID)
    
    on error resume next
    deleteActivity00(delID)

    if Err<>0 then
        bottomMessage = "Error deleting activity"
    else
        ' Cierra la ventana
        
        %><script language="JavaScript">try{window.close();}catch(e){} try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} try{window.opener.applyFilter(false);}catch(e){}</script><%
    end if
    
End Sub


Select Case EventObject
	case "Save" Save_click()
	case "Delete" Delete_click(EventParam1)
End Select


dim act

if CInt(ID) > -1 then
    set act = getActivity00(ID)
else
    ' Es un elemento nuevo
    set act = new Activity00
end if

dim aType, aCli, aBra
set aType = getActivityType(IDType, Idioma)
set aCli = getClient(IDClient)
set aBra = getBrand(IDBrand)

%>
<HTML>
<HEAD>
    <TITLE>Activity</TITLE>
    <LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
    <script>
        
        var dataModified = false;
        var themeList;
        
        function closeWindow()
        {
            if (dataModified)
            {
                if (confirm('<%=IDM_JS_DatosModificadosGuardar %>')){
                    thisForm.CloseWindow.value = '1';
                    Save();
                    return false;
                }
            }
            
          <%if Request("PageReloaded")<>"" then %> 
            try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} 
            try{window.opener.applyFilter(false);}catch(e){} 
          <%end if %> 
          
          try{window.close();}catch(e){}
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
            <%if act.IDTheme<>"" then %>
                sDat += ';<%=act.IDTheme %>';
            <%end if %>
			ajaxres = ajaxReq('ListaTematicas', sDat);
			
			themeList = eval('('+ ajaxres +')');
			
			
			//Borra la lista de Themes actual
			thisForm.IDTheme.options.length=0;
			
			// Incluye en la lista el primer elemento vacío
			var objOption = document.createElement("option");
			objOption.text = '  <%=IDM_SELECT_TemasDe %> <%=aCli.Name %>' ;
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
            dataModified = true;
        }
        function tematicaBorrada(id)
        {
            cargarTematicas(id);
            thisForm.IDTheme.value = id;
            dataModified = true;
        }
        function tematicaImagenBorrada()
        {
            document.images["IMG_Theme"].src = '';
            document.images["IMG_Theme"].style.display='none';
        }
    </script>
</HEAD>

<BODY leftmargin=0 topmargin=0 >

<FORM action="SOAActivity00.asp?ID=<%=ID %>" method="post" name="thisForm">
    
    
    <table style="width:100%;height:40px;background-image:url('images/Grad5.gif'); ">
        <tr>
            <td valign="middle" style="padding-left:10px;">
                <font class="wopenTitle">
                    <%dim nTitleChars: nTitleChars = 25 %>
                    <%=left(act.GridText, nTitleChars) %><%if len(act.GridText)>nTitleChars then %>...<%end if %>
                </font>
            </td>
            <td align="right" width="180px;">
                <input type=button class="button" value="<%=IDM_Save %>" style="width:55px;" onclick="Save();" />
                <%if ID <> "-1" then %>
                    <input type=button class="button" value="<%=IDM_Delete %>" style="width:55px;" onclick="Delete(<%=ID %>);" />
                <%end if %>
                <input type=button class="button" value="<%=IDM_Close %>" style="width:55px;" onClick="closeWindow();" />
            </td>
        </tr>
    </table>
    
    <table style="width:100%;height:30px;border:2 solid gray;">
        <tr>
            <td valign=top width=100 class="fieldheader" ><%=IDM_ActivityType %></td>
            <td><font class="font12"><%=aType.Name %></font></td>
            <%
            dim sTiming
            if WHalf = 1 then
                sTiming = IDM_1aQuincena
            else
                sTiming = IDM_2aQuincena
            end if
            sTiming = sTiming & " " & locMonthName(WMonth, Idioma)
            sTiming = sTiming & " " & WYear
            %>
            <td align=right><font class="font15"><b><%=sTiming %></b></font></td>
        </tr>
        <tr>
            <td valign=top width=100 class="fieldheader" ><%=IDM_Client %></td>
            <td><font class="font12"><%=aCli.Name %></font></td>
        </tr>
        <tr>
            <td valign=top width=100 class="fieldheader" ><%=IDM_Brand %></td>
            <td><font class="font12"><%=aBra.Name %></font></td>
        </tr>

    </table>
    
    <%if FALSE then %>
        <table style="width:100%;height:30px;">
            <tr>
                <td width=100 valign=top class="fieldheader"><%=IDM_Tematica %></td>
                <td valign=top >
                    <select name="IDTheme" style="width:100%;" class="textfield" onchange="dataModified=true; cambioTematica();"></select>
                    <img ID="IMG_Theme" src="" style="width:<%=Application("ThemeImageWidth") %>px;" />
                </td>
                <td width=90 align=right valign=top>
                    <a title="<%=IDM_ModificarTematica %>" ID="BTN_ModifTheme" href="" onclick="window.open('SOAAddTheme.asp?ID=' + thisForm.IDTheme.value + '&IDClient=<%=IDClient %>','ADMTHM','width=500,height=300,top=200,left=250,scrollbars');return false;"><img src="images/edit30.png" style="border:0px;" /></a>
                    <a title="<%=IDM_NuevaTematica %>" href="" onclick="window.open('SOAAddTheme.asp?ID=-1&IDClient=<%=IDClient %>','ADMTHM','width=500,height=300,top=200,left=250,scrollbars');return false;"><img src="images/add.png" style="border:0px;" /></a>
                </td>
            </tr>

        </table>
    <%end if %>
    
    
    <table style="width:100%;height:30px;">
        <%if TRUE then %>
        <tr>
            <td valign=top width=100 class="fieldheader" ><%=IDM_ExtraInfo %></td>
            <td>
                <textarea onchange="dataModified=true;" name="Name" class="textfield" style="width:100%;height:60px;"><%=act.Name %></textarea>
            </td>
        </tr>
        <%else %>
            <input type=hidden name="Name" value="" />
        <%end if %>

        <tr>
            <td width=100 class="fieldheader"><%=IDM_Ratio %></td>
            <td>
                <select name="IDRatio" style="width:100%;" class="textfield" onchange="dataModified=true;">
                <%
                set rats = getActivityRatios(Idioma)
                %>
                </select>
            </td>
        </tr>

        <%if ID <> "-1" then %>
            <tr height=20><td></td></tr>
            <tr>
                <td valign=top width=100 class="fieldheader" style="border-top:1 solid silver;"><%=IDM_LastUpdatedBy %></td>
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
    <input type=hidden name="IDType" value="<%=IDType %>" />
    
    <input type=hidden name="CloseWindow" value="" />

    <!-- #include file = "include/EventFunctions2.asp" -->

</FORM>


<!-- #include file = "include/pageBottom.asp" -->

<script language=javascript>
    cargarTematicas('<%=act.IDTheme %>');
    cambioTematica();
</script>

</BODY>

</HTML>