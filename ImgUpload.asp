<%Response.Expires=0%>

<%

if Session("WI_ADMNombreUsuario")="" then
	Response.End
end if

Lang = "ES"
LangFolder = "/"


set conn = CreateObject("ADODB.Connection")
set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")
conn.Open Session("WI_ConnectTo")

%>

<!--#include file="include/adovbs.asp" -->
<!--#include file="include/SrvrFunctions.asp" -->
<!--#include file="include/EventFunctions1.asp" -->

<!--#include file="FolderSettings.asp" -->
<!--#include file="LangFunctions.asp" -->


<%


Sub DocUpload_click()
    
    
	'Crea un IDDocumento y recupera el número ****************
	conn.BeginTrans
	
	SQL = "INSERT INTO Documento (Nombre) VALUES ('')"
	conn.Execute(SQL)
	
	SQL = "SELECT MAX(ID) AS MAXID FROM Documento "
	rst.Open SQL, conn, adOpenStatic, adLockReadOnly
	NewIDDoc = rst("MAXID")
	rst.Close
	
	conn.CommitTrans
	'*******************************************************
    
   	Set FFichero = Upload.Files("Fichero")
   	
	on error resume next
	NFichero = FFichero.FileName
	if NFichero<>"" then FFichero.MoveVirtual session("WI_DataVFolder") & "/DOCS/DC" & Right("000" & NewIDDoc, 4) & "-" & NFichero
	if NFichero<>"" then NFichero = FFichero.FileName
	
	if Err<>0 then
	    %><script language="JavaScript">try{alert('Error subiendo fichero');history.go(-1);}catch(e){}</script><%
	    Exit Sub
	end if
	on error goto 0

	if NFichero="" then
		strNFichero = "NULL"
	else
		strNFichero = "N'" & replace(left(NFichero,150),"'","''") & "'"
	end if
    
    
    SQL = "UPDATE Documento " & _
	" SET " & _
	" idRelacionado = " & IDRelacionado & ", " & _
	" Nombre = N'" & replace(left(Nombre,150), "'", "''") & "', " & _
	" URLDoc = " & strNFichero & ", " & _
	" Tipo = '" & Tipo & "', " & _
	" Idioma = '" & Idioma & "' " & _
	" WHERE ID= " & NewIDDoc
	on error resume next
	conn.Execute(SQL)
	if Err<>0 then
	    %>
	    <p align=center>
	    <font face="Arial" size=3>
	    Error subiendo archivo
	    <br /><br />
	    <b>
	    <%=Err.Description %>
	    </b>
	    <br /><br />
	    Por favor, cierre esta ventana y vuelva a intentar.
	    </font>
	    <br /><br />
	    <input type=button value="Cerrar" onclick="try{window.close()}catch(e){}" />
	    
	    </p>
	    <%
	    Response.End
	    Exit Sub
	end if
	on error goto 0

	%>
	<script language="JavaScript">
		try{
		    window.opener.thisForm.EventObject.value='DocUploaded';
		    window.opener.thisForm.submit();
		}catch(e){}
		try{window.close();}catch(e){}
	</script>
	<%

    
End Sub




IDRelacionado = Request("ID")
if IDRelacionado = "" then
    IDRelacionado = Request.QueryString("IDRelacionado")
end if
if IDRelacionado = "" then
    response.Write "Invalid ID"
    response.End
end if

Tipo = Request("T")
if Tipo = "" then
    response.Write "Invalid Tipo"
    response.End
end if



set rst = CreateObject("ADODB.RecordSet")
dim filesys


if Request("UP")<>"" then
    
   	'SAVE THE UPLOADED THE FILES
	Set Upload = Server.CreateObject("Persits.Upload")
	Upload.IgnoreNoPost = True
	Upload.OverwriteFiles = FALSE
	
	on error resume next
	Upload.SaveVirtual(session("WI_DataVFolder") & "/DOCS")
	if Err<>0 then
	    %>
	    <p align=center>
	    <font face="Arial" size=3>
	    Error subiendo archivo
	    <br /><br />
	    <b>
	    <%=Err.Description %>
	    </b>
	    <br /><br />
	    Por favor, cierre esta ventana y vuelva a intentar.
	    </font>
	    <br /><br />
	    <input type=button value="Cerrar" onclick="try{window.close()}catch(e){}" />
	    
	    </p>
	    <%
	    response.End
	end if
	on error goto 0

	Nombre = Upload.Form.Item("Nombre")
	Idioma = Upload.Form.Item("Idioma")
	
	EventObject = Upload.Form.Item("EventObject")
	EventParam1 = Upload.Form.Item("EventParam1")
	EventParam2 = Upload.Form.Item("EventParam2")


	select case EventObject
		case "DocUpload" call DocUpload_click()
	end select

    
end if

%>


<html>
<head>
<title>Nuevo documento</title>
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function UploadFile(){
    
    if (thisForm.Nombre.value==''){alert('Por favor, escriba un nombre para el documento');return false;}
    if (thisForm.Fichero.value==""){alert('Por favor, seleccione un fichero');return false;}

	thisForm.EventObject.value="DocUpload";

	thisForm.action = "DocUpload.asp?UP=1&ID=<%=IDRelacionado%>&T=<%=Tipo %>";
	thisForm.encoding="multipart/form-data";
	thisForm.submit();
	
}	

</script>


</head>

<body topmargin=0 leftmargin=0 rightmargin=0>
<form name="thisForm" method="post" action="" enctype="multipart/form-data">
	
	<TABLE width="100%" cellspacing=0>
		<TR height=40 bgcolor=#32a4e7>
			<TD style="padding-left:15px;"><font class="about_title1"><font color=white>Nuevo documento</TD>
			<TD align=right>
				<input type="button" value="Guardar" onclick="UploadFile();" id=BtnUpload name=BtnUpload>
				<input type="button" value="Cancelar" onclick="window.close();">
			</TD>
		</TR>
	</TABLE>
	
	<TABLE width="100%" cellspacing=0 >
	    <TR>
	        <td width=80 style="padding-left:15px;"><font class=font12>Nombre</font></td>
		    <TD><input type="text" style="width:100%;font-size:9px;" name="Nombre"></TD>
	    </TR>

	    <TR>
	        <td style="padding-left:15px;"><font class=font12>Idioma</td>
		    <TD>
		        <select name="Idioma" style="width:150px;font-size:9px;">
		            <%
		            SQL = "SELECT id, Nombre FROM Idioma ORDER BY id"
		            rst.Open SQL, conn, adOpenStatic, adLockReadOnly
		            while not rst.EOF
		                strSelected = ""
		                if rst("id") = "ES" then
		                    strSelected = " selected "
		                end if
		                %>
    		            <option value="<%=rst("id") %>" <%=strSelected %>><%=rst("Nombre") %></option>
		                <%
		                rst.MoveNext
		            wend
		            rst.Close
    	            %>
		        </select>
		    </TD>
	    </TR>
	    <TR>
	        <td style="padding-left:15px;"><font class=font12>Fichero</td>
		    <TD><input type="file" style="WIDTH: 100%" name="Fichero" style="font-size:9px;background-color:Gainsboro;"></TD>
	    </TR>
	    
	</TABLE>




<input type="hidden" name="EventObject" >
<input type="hidden" name="EventParam1" >
<input type="hidden" name="EventParam2" >


</form>

</body>
