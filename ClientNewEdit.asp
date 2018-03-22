<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
	Response.Buffer=true
	%>


<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/EventFunctions1.asp" -->

<%
	RecoverSQLConnection()
	RecoverSession(true) 
%>
<!-- #include file = "include/Idioma.asp" -->


<%
dim rst, SQL

	Set rst = Server.CreateObject("ADODB.Recordset")

	SQL = "SELECT u.IDEmpleado " & _
	" FROM Users u " & _
	" INNER JOIN UserGroup g ON u.IDEmpleado=g.IDEmpleado " & _
	" WHERE g.IDGroup=0 AND u.IDEmpleado=" & session("IDEmpleado")
	'PrintSQL
	rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
	if rst.EOF then
		Response.Write "You are not allowed to see this page"
		Response.End
	end if
	rst.Close
%>

<%

'-------------------------------------------------------------
'-------------------------------------------------------------
Sub SaveClient_click()
	dim rstUs, SQL, sOrden, sActivateForms
	on error resume next

	if indBaja<>"" then
		indBaja = 1
	else
		indBaja = 0
	end if
	
	if activateForms<>"" then
	    sActivateForms = 1
	else
        sActivateForms = 0
	end if
	
	sOrden = CInt(Orden) & ""
	if Err<>0 then
	    msgError "El campo Orden debe ser numérico", true, true
	    exit sub
	end if
	Err.Clear
	
	Set rstUs = Server.CreateObject("ADODB.Recordset")
	if IDClient<>"" then	'No es un cliente nuevo
		
		SQL = "UPDATE Client SET Name=N'" & replace(LEFT(Name, 50), "'", "''") & "', " & _
		" ShortName=N'" & replace(ShortName, "'", "''") & "', " & _
		" SiebelCode=N'" & replace(LEFT(SiebelCode, 50), "'", "''") & "', " & _
		" Orden = " & sOrden & ", " & _
		" indBaja=" & indBaja & ", " & _
		" activateForms = " & sActivateForms & _
		" WHERE IDClient=" & IDClient
		ObjConnectionSQL.Execute SQL
		if Err<>0 then
			MsgError "Error 10: " & Err.Description, true, true
		end if
		
		Response.Write "<script language=""JavaScript"">try{window.opener.window.thisForm.submit();window.close()}catch(e){}</script>"
		Response.End
		exit sub

	else	'Es un cliente nuevo
		
		SQL = "INSERT INTO Client (Name, ShortName, indBaja, SiebelCode, Orden, activateForms) " & _
		" VALUES (N'" & replace(LEFT(Name, 50), "'", "''") & "', " & _
		" N'" & LEFT(replace(ShortName, "'", ""), 17) & "', " & _
		" " & indBaja & ", " & _
		" N'" & replace(LEFT(SiebelCode, 50),"'","''") & "', " & _
		" " & sOrden & ", " & sActivateForms & ")"
		ObjConnectionSQL.Execute SQL
		if Err<>0 then
			MsgError "Error 20: " & Err.Description, true, true
			exit sub
		end if

		Response.Write "<script language=""JavaScript"">try{window.opener.window.thisForm.submit();window.close()}catch(e){}</script>"
		Response.End
		exit sub
	end if
	
End Sub



'-------------------------------------------------------------
'-------------------------------------------------------------
'Manté els valors en el formulari ----------------------------
dim IDClient: IDClient = request("IDClient")
dim Name: Name = Request.Form("Name")
dim ShortName: ShortName = Request.Form("ShortName")
dim SiebelCode: SiebelCode = Request.Form("SiebelCode")
dim Orden: Orden = Request.Form("Orden")
dim activateForms: activateForms = Request.Form("activateForms")

dim indBaja: indBaja = Request.Form("indBaja")



'-------------------------------------------------------------
'-------------------------------------------------------------
'Reconeix l'event --------------------------------------------
EventObject = request("EventObject")
EventParam1 = request("EventParam1")
EventParam2 = request("EventParam2")
if EventObject = "SaveClient" then
	call SaveClient_click()

end if

%>



<script language="JavaScript">
function SaveClient(){
	if (thisForm.Name.value==''){alert('Escriba el nombre del cliente'); return false;}
	if (thisForm.Orden.value==''){alert('Escriba el orden'); return false;}
	_fireEvent('SaveClient','','')
}

</script>

<html>
<head>
	<title><%=IDM_ClientNewEditTitle%></title>
	<LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
</head>
<body topmargin=0 leftmargin=5>

<form action="ClientNewEdit.asp" method="post" name="thisForm">

	<table style="border:2 solid slateGray" border=0 cellpadding=0 cellspacing=0 width="100%">
	<tr>
		<td bgcolor=slateGray width=35></td>
		<td bgcolor=slateGray style="padding-left:5" align="left">
			<FONT class="font20"><font color=white><STRONG><%=IDM_ClientNewEditTitle%></STRONG></FONT>
		</td>
		<td bgcolor=slateGray align=right>
			<%'Botons de guardar, cancel·lar%>
			<input alt="Close" type="image" style="border:0" src="images/iCancel.gif" value="<%=IDM_Close%>" onclick="window.close();return false;">
			<input alt="Save" type="image" style="border:0" src="images/iSave.gif" value="<%=IDM_Save%>" onclick="SaveClient();return false;">
		</td>
	</tr>
	</table>

<br>


	<%'Están editando o creando un nuevo usuario
		if IDClient<>"" then	'Están EDITANDO
			SQL = "select c.* " & _
			" FROM Client c " & _
			" WHERE c.IDClient=" & IDClient
			
			rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
			if not rst.EOF then
				Name = rst("Name")
				ShortName = rst("ShortName")
				SiebelCode = rst("SiebelCode")
				indBaja = rst("indBaja")
				Orden = rst("Orden")
				activateForms = rst("activateForms")
			else
				MsgError "", true, true
			end if
			rst.Close
		end if%>
		
		

		<table border="0" width="100%" cellpadding=0 cellspacing=0>
		<tr><td>
		  <table border=0 width="100%" cellpadding=2 cellspacing=0>
				<tr><td colspan=4>
<!--					
					<input style="width:80" type="button" value="Save" onclick="SaveUser();">
-->				</td></tr>
				<tr>
					<td	valign="top" width=150>
  						<font class="font11">
  							<%=IDM_Name%>
  						</font>
					</td>
					<td valign="top" width=220>
						<input style="width:200" class="recuadro" maxlength=50 name="Name" value="<%=Name%>">
						<input type="hidden" name="IDClient" id="IDClient" value="<%=IDClient%>">
					</td>
				</tr>
				<%if FALSE then %>
				    <tr>
					    <td	valign="top" >
  					      <font class="font11"><%=IDM_ShortName%></font>
					    </td>
					    <td valign="top">
						    <input style="width:200" class="recuadro" name="ShortName" maxlength=17 value="<%=ShortName%>">
					    </td>
				    </tr>
				<%end if %>
				<tr>
					<td	valign="top" >
  					  <font class="font11"><%=IDM_PlanTo%></font>
					</td>
					<td valign="top">
						<input style="width:200" class="recuadro" name="SiebelCode" maxlength=50 value="<%=SiebelCode%>">
					</td>
				</tr>
				<tr>
					<td	valign="top" >
  					  <font class="font11"><%=IDM_Orden%></font>
					</td>
					<td valign="top">
						<input style="width:200" class="recuadro" name="Orden" maxlength=5 value="<%=Orden%>">
					</td>
				</tr>
				<tr>
					<td	valign="top" >
  					  <font class="font11"><%=IDM_ClientFormsActivated%></font>
					</td>
					<td valign="top">
						<input type="checkbox" class="recuadro" name="activateForms" <%if activateForms<>0 then%> checked <%end if%>>
					</td>
				</tr>
				<tr>
					<td	valign="top" >
  					  <font class="font11"><%=IDM_Deleted%></font>
					</td>
					<td valign="top">
						<input type="checkbox" class="recuadro" name="indBaja" <%if indBaja<>0 then%> checked <%end if%>>
					</td>
				</tr>
			</table>
		</td></tr>
		</table>
				

<!-- #include file = "include/EventFunctions2.asp" -->



</form>

<script language="JavaScript">window.focus();</script>

</body>
</html>

<%Response.Flush%>