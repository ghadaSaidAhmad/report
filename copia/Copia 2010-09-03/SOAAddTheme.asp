<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
%>

<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->

<!-- #include file = "ClassInclude.asp" -->



<%
dim rst, rst2, sChecked
set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")

rst.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient

RecoverSQLConnection()

RecoverSession(true) 


%>
<!-- #include file = "include/Idioma.asp" -->

<%
dim ID: ID = Request.QueryString("ID")
if ID = "" then ID = -1
dim IDClient: IDClient = Request.QueryString("IDClient")
if IDClient="" then IDClient = -1
dim PageReloaded
if Request.QueryString("U")<>"" then
    PageReloaded = "1"
else
    PageReloaded = ""
end if
dim CloseWindow

dim Name
dim indBaja
dim ImageFileName


Dim DestinationPath
DestinationPath = Server.mapPath("images/Themes")

'Create upload form
'Using Pure-ASP file upload
Dim Form: Set Form = New ASPForm %><!--#INCLUDE FILE="include/UploadClass.asp"--><% 
Server.ScriptTimeout = 2000
Form.SizeLimit = &HA00000
Dim MaxFileSize: MaxFileSize = 1024*1024*3

If Form.State = 0 Then 'Completed
	Dim File
	
    EventObject = Form.Item("EventObject")
    EventParam1 = Form.Item("EventParam1")
    EventParam2 = Form.Item("EventParam2")
    CloseWindow = Form.Item("CloseWindow")

    Name = Form.Item("Name")
    indBaja = Form.Item("indBaja")

	'For Each File In Form.Files
	dim iItem: iItem = 0
	dim fso: set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	For each File in Form.Files.Items
		if File.Length<=MaxFileSize Then
		    if File.isFile AND File.FileName <> "" then
    		    'Response.Write "File.FileName [" & File.FileName & "]<br>"
    		    'Response.Write "File.Name [" & File.Name & "]<br>"
    		    'Response.Write "File.FilePath [" & File.FilePath & "]<br>"
    		    'Response.Write "File.isFile [" & File.isFile & "]<br>"
    		    'Response.Write "File.Length [" & File.Length & "]<br>"
	    	    'Response.Write "MaxFileSize [" & MaxFileSize & "]<br>"

        		File.Save DestinationPath
        		
        		on error resume next
        		
        		if fso.FileExists(DestinationPath & "\" & Right("0000" & ID, 5) & "_" & File.FileName) then
        		    fso.DeleteFile DestinationPath & "\" & Right("0000" & ID, 5) & "_" & File.FileName
        		end if
        		
        		fso.MoveFile DestinationPath & "\" & File.FileName, DestinationPath & "\" & Right("0000" & ID, 5) & "_" & File.FileName
        		
        		ImageFileName = Right("0000" & ID, 5) & "_" & File.FileName
        		
        		if Err<>0 then
    				Response.Write "<script language=""JavaScript"">alert('Hubo un error subiendo el fichero.\n\rPor favor, inténtelo de nuevo y contacte con \n\rel administrador si el error persiste');history.back(-1);</script>"
        		    response.End
        		end if
        		on error goto 0
        		
        	end if
		else
			Response.Write "<script language=""JavaScript"">alert('Ha sobrepasado el tamaño máximo del fichero (" & (MaxFileSize / 1024 / 1024) & " MB)');history.back(-1);</script>"
			Response.End
		end if
	    
	    
        iItem = iItem + 1
    Next
ElseIf Form.State > 10 then
	Const fsSizeLimit = &HD
	Select case Form.State
		case fsSizeLimit: 
			Response.Write "<script language=""JavaScript"">alert('Tamaño máximo del fichero sobrepasado " & Form.TotalBytes & "B (Máximo " & Form.SizeLimit & "B)');history.back(-1);</script>"
			Response.End
		case else 
		    response.Write "kklvk"
			'Response.Write "<script language=""JavaScript"">alert('Error subiendo fichero');history.back(-1);</script>"
			'Response.End
	end Select
End If



dim EventObject, EventParam1, EventParam2






Sub Save_click()
    
    dim thm
    set thm = new Theme
    thm.ID = CInt(ID)
    thm.IDClient = IDClient
    if indBaja<>"" then
        thm.indBaja = 1
    else
        thm.indBaja = 0
    end if
    
    thm.Name = Name
    
    if ImageFileName<>"" then
        thm.ImageFileName = ImageFileName
    end if
    
    on error resume next
    saveTheme(thm)
    if Err<>0 then
        bottomMessage = Err.Description
    else
        'Si era nuevo, ahora tiene un ID

        ID = thm.ID

    end if
    on error goto 0
    
    if CloseWindow <> "" then
        %>
        <script language="JavaScript">
            try{window.close();}catch(e){} 
            try{window.opener.nuevaTematicaCreada(<%=ID%>);} catch(e){} 
        </script><%
    end if
    
End Sub

Sub Delete_click(DelID)
    
    on error resume next
    deleteTheme(delID)
    if Err<>0 then
        bottomMessage = "Error deleting theme"
    else
        ' Cierra la ventana
        
        %><script language="JavaScript">try{window.opener.tematicaBorrada(<%=DelID %>);} catch(e){} try{window.close();}catch(e){}</script><%
    end if
    
End Sub


Sub RemoveImage_click()
    
	dim fso: set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	dim thm, file
	set thm = getTheme(CInt(ID))
	file = thm.ImageFileName
	
	on error resume next
	if fso.FileExists(DestinationPath & "\" & file) then
	    fso.DeleteFile DestinationPath & "\" & file
	end if
	on error goto 0
	
	
    on error resume next
    removeThemeImage(CInt(ID))
    if Err<>0 then
        bottomMessage = Err.Description
    end if
    on error goto 0
    
End Sub


Select Case EventObject
	case "Save" Save_click()
	case "Delete" Delete_click(EventParam1)
	case "RemoveImage" RemoveImage_click()
End Select


dim thm
if CInt(ID) > -1 then
    set thm = getTheme(ID)
else
    set thm = new Theme
end if

dim aCli, clientName
if IDClient <> "-1" then
    set aCli = getClient(IDClient)
    clientName = aCli.Name
end if
%>

<HTML>
<HEAD>
    <TITLE><%=IDM_Tematica %></TITLE>
    <LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
    <script language="javascript">
        var dataModified = false;
        
        function _fireEvent (Objeto, Param1, Param2)
        {	
	        thisForm.EventObject.value = Objeto;
	        thisForm.EventParam1.value = Param1;
	        thisForm.EventParam2.value = Param2;			
	        thisForm.submit();
        }
        function _fireConfirm(Objeto, Param1, Param2, MSG)
        {
	        if (MSG!=""){
		        if (confirm(MSG)){
			        _fireEvent(Objeto,Param1,Param2);
		        }
	        }
	        else if (window.confirm("Click OK to continue. Click Cancel to abort.")){
		        _fireEvent(Objeto,Param1,Param2);
	        }
        }
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
            
            <%if PageReloaded <> "" then %>
                try{window.opener.nuevaTematicaCreada(<%=ID %>);} catch(e){} 
            <%end if %> 
            
            try{window.close();}catch(e){}
            
        }
        function Save()
        {
            _fireEvent('Save', '', '');
        }
        
        function Delete(id)
        {
            _fireConfirm('Delete', id, '', '');
        }

    </script>
</HEAD>

<BODY leftmargin=0 topmargin=0 >

<FORM action="SOAAddTheme.asp?U=1&ID=<%=ID %>&IDClient=<%=IDClient %>" method="post" name="thisForm" enctype="multipart/form-data">
    
    
    <table style="width:100%;height:40px;background-image:url('images/Grad5.gif'); ">
        <tr>
            <td valign="middle" style="padding-left:10px;">
                <font class="wopenTitle">
                    <%if IDClient <> "-1" then %>
                        <%=IDM_TemaDeCliente %>
                    <%else %>
                        <%=IDM_GeneralTheme %>
                    <%end if %>
                    
                    <%if FALSE then %>
                        <%dim nTitleChars: nTitleChars = 15 %>
                        <%=left(thm.Name, nTitleChars) %><%if len(thm.Name)>nTitleChars then %>...<%end if %>
                    <%end if %>
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
        <%if clientName <> "" then %>
            <tr>
                <td valign=top width=120 class="fieldheader" ><%=IDM_Client %></td>
                <td><font class="font12"><%=clientName %></font></td>
            </tr>
        <%else %>
            <tr>
                <td valign=top width=120 class="fieldheader" ><%=IDM_Type %></td>
                <td><font class="font12"><%=IDM_GenericTheming %></font></td>
            </tr>
        <%end if %>

    </table>
    
    <table style="width:100%;height:30px;">
        <tr>
            <td valign=top width=120 class="fieldheader" ><%=IDM_Nombre %></td>
            <td>
                <input type="text" name="Name" class="textfield" value="<%=thm.Name %>" style="width:100%;" onchange="dataModified = true;" />
            </td>
        </tr>
        
        <tr>
            <td valign=top width=120 class="fieldheader" ><%=IDM_Image %></td>
            <td>
                <%if CInt(ID)<>-1 then %>
                    <%if thm.ImageFileName<>"" then %>
                        <table width="100%"><tr>
                            <td><img src="images/Themes/<%=replace(Server.URLEncode(thm.ImageFileName),"+"," ") %>" width="<%=Application("ThemeImageWidth") %>" /></td>
                            <td width=120><input type=button class=button value="<%=IDM_RemoveImage %>" onclick="_fireEvent('RemoveImage','','');return false;" /></td>
                        </tr></table>
                    <%else %>
                        <input class="textfield" type=file name="File" onchange="dataModified = true;" style="width:100%" />
                    <%end if %>
                <%else %>
                    <font class="font11"><%=IDM_GuardeTemaParaAgregarImg %></font>
                <%end if %>
            </td>
        </tr>
        
        <%if ID <> "-1" then %>
            <tr>
                <td valign=top width=120 class="fieldheader" ><%=IDM_indBaja %></td>
                <td>
                    <%
                    sChecked = ""
                    if thm.ID > -1 then
                        if thm.indBaja <> 0 then
                            sChecked = "checked"
                        end if
                    end if
                    %>
                    <input type="checkbox" name="indBaja" <%=sChecked %> onchange="dataModified = true;" />
                </td>
            </tr>
        <%end if %>
        
        <%if ID <> "-1" then %>
            <tr height=20><td></td></tr>
            <tr>
                <td valign=top width=120 class="fieldheader" style="border-top:1 solid silver;"><%=IDM_LastUpdatedBy %></td>
                <td style="border-top:1 solid silver;"><font class=font12>
                    <%=thm.LastUpdatedBy %>
                    &nbsp;-&nbsp;
                    <%=thm.LastUpdatedDate %>
                    </font>
                </td>
            </tr>
        <%end if %>

    </table>
    
    
    
    
    
    <input type=hidden name="CloseWindow" value="" />
    
    <!-- #include file = "include/EventFunctions2.asp" -->

</FORM>


<!-- #include file = "include/pageBottom.asp" -->

<script language=javascript>
    thisForm.Name.focus();
</script>

</BODY>

</HTML>