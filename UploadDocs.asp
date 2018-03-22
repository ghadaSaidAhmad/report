<%@language=VBScript%>
<%Response.Expires=0
	Response.Buffer=true%>

<%
TMP = Request.QueryString("TMP")
Accion = Request.QueryString("AC")
CurrentFileName = Request.QueryString("CFN")

if Accion<>"" then
    
    dim fn
    if TMP<>"" then
        fn = TMP & "_" & CurrentFileName
    else
        fn = CurrentFileName
    end if
    
    set fso = Server.CreateObject("Scripting.FileSystemObject")
    if fso.FileExists(Server.mapPath("images/Themes") & "\" & fn) then
        fso.DeleteFile Server.mapPath("images/Themes") & "\" & fn
    end if
    
    %>
    <script language="JavaScript">
        window.parent.fileRemoved();
        //location.href='UploadDocs.asp?TMP=<%=TMP %>';
    </script>
    <%
    Response.End
    
end if

    
    Dim FileName 'Nombre del fichero subido
    
    
    Dim DestinationPath
	DestinationPath = Server.mapPath("images/Themes")

	'Create upload form
	'Using Pure-ASP file upload
	Dim Form: Set Form = New ASPForm %><!--#INCLUDE FILE="UploadClass.asp"--><% 

	Server.ScriptTimeout = 2000
	Form.SizeLimit = &HA00000
	MaxFileSize = 1024*1024*3

	If Form.State = 0 Then 'Completed
		Dim File
		

		'For Each File In Form.Files
		iItem = 0
		For each File in Form.Files.Items
    		'Response.Write "File.FileName [" & File.FileName & "]<br>"
    		'Response.Write "File.Name [" & File.Name & "]<br>"
    		'Response.Write "File.FilePath [" & File.FilePath & "]<br>"
    		'Response.Write "File.isFile [" & File.isFile & "]<br>"
    		'Response.Write "File.Length [" & File.Length & "]<br>"
	    	'Response.Write "MaxFileSize [" & MaxFileSize & "]<br>"
			
			fileOK = false
			
			if File.Length<=MaxFileSize Then
			    if File.isFile then
            		File.Save DestinationPath
            		
            		
            		set fso = Server.CreateObject("Scripting.FileSystemObject")
            		
            		on error resume next
            		
            		if fso.FileExists(DestinationPath & "\" & TMP & "_" & File.FileName) then
            		    fso.DeleteFile DestinationPath & "\" & TMP & "_" & File.FileName
            		end if
            		
            		fso.MoveFile DestinationPath & "\" & File.FileName, DestinationPath & "\" & TMP & "_" & File.FileName
            		
            		if Err<>0 then
        				Response.Write "<script language=""JavaScript"">alert('Hubo un error subiendo el fichero.\n\rPor favor, inténtelo de nuevo y contacte con \n\rel administrador si el error persiste');history.back(-1);</script>"
            		    response.End
            		end if
            		on error goto 0
            		
            		fileOK = true
            		
            	end if
			else
				Response.Write "<script language=""JavaScript"">alert('Ha sobrepasado el tamaño máximo del fichero (" & (MaxFileSize / 1024 / 1024) & " MB)');history.back(-1);</script>"
				Response.End
			end if

    		
            CurrentFileName = File.FileName
            
		    if Trim(CurrentFileName)<>"" then
		        
                
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

    
%>
<html>
<head>
    <LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
</head>
<body topmargin=0 leftmargin=0 bgcolor="silver">

<%if TMP<>"" AND (Request.QueryString("U")<>"" OR CurrentFileName<>"") then %>

    <form name="thisForm" method="post" action="?TMP=<%=TMP %>" >
        <div align=left style="padding-left:30px;">&nbsp;&nbsp;&nbsp;
        <a href="" title="Quitar fichero adjunto" onclick="if (!confirm('Desea quitar el documento adjunto?')){ return false; } thisForm.action += '&AC=DEL&CFN=<%=CurrentFileName %>'; thisForm.submit();return false;">
            <img src="images/borrar.gif" border=0 />
        </a>
        &nbsp;&nbsp;&nbsp;
        <a target="_blank" href="images/Themes/<%=TMP & "_" & CurrentFileName %>"><font class=font12><%=CurrentFileName %></font></a>
        <input type=hidden name="File_UPLOADED" value="OK" />
        </div>
    </form>
    <script language="JavaScript">
        window.parent.fileSelected('<%=CurrentFileName %>');
    </script>
    
<%elseif TMP="" then %>

    <form name="thisForm" method="post" action="" >
        <div align=left style="padding-left:30px;">&nbsp;&nbsp;&nbsp;
        <a href="" title="Quitar fichero adjunto" onclick="if (!confirm('Desea quitar el documento adjunto?')){ return false; } thisForm.action += '?AC=DEL&CFN=<%=CurrentFileName %>'; thisForm.submit();return false;">
            <img src="images/borrar.gif" border=0 />
        </a>
        &nbsp;&nbsp;&nbsp;
        <a target="_blank" href="images/Themes/<%=CurrentFileName %>"><font class=font12><%=Mid(CurrentFileName, 22) %></font></a>
        </div>
    </form>
    
<%else %>

    <form name="thisForm" method="post" action="?TMP=<%=TMP %>&U=1" enctype="multipart/form-data">
        <div align=left style="padding-left:50px;">
            <font class=font12>Seleccione el fichero pulsando este botón: </font>
            <input type=file name="File" style="width:0px;" onchange="showIMGWAIT();thisForm.submit();"/>    
        </div>
    </form>
    
<%end if %>


<script language="JavaScript">
    function hideIMGWAIT()
    {
        try{
            window.parent.thisForm.IMG_WaitConfirmo.style.display='none';
        }
        catch(e)
        {
        }        
    }
    
    function showIMGWAIT()
    {
        try{
            window.parent.thisForm.IMG_WaitConfirmo.style.display='';
        }
        catch(e)
        {
        }        
    }
    

    // Esconde la imagen WAIT
    hideIMGWAIT();
    
</script>


</body>
</html>
