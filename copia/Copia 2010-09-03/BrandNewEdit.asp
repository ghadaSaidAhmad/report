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



'-------------------------------------------------------------
'-------------------------------------------------------------
Sub SaveBrand_click()
dim rstUs, SQL, sOrden

	on error resume next

	if indBaja<>"" then
		indBaja = 1
	else
		indBaja = 0
	end if

	sOrden = CInt(Orden) & ""
	if Err<>0 then
	    msgError "El campo Orden debe ser numérico", true, true
	    exit sub
	end if
	Err.Clear

	Set rstUs = Server.CreateObject("ADODB.Recordset")
	if IDBrand<>"" then	'No es un Brand nuevo
		
		SQL = "UPDATE Brand SET Name=N'" & replace(LEFT(Name, 100), "'", "''") & "', " & _
		" ShortName=N'" & LEFT(replace(ShortName, "'", ""), 17) & "', " & _
		" SiebelCode=N'" & replace(LEFT(SiebelCode, 50), "'", "''") & "', " & _
		" Orden=" & sOrden & ", " & _
        " indBaja=" & indBaja & " WHERE IDBrand=" & IDBrand
		ObjConnectionSQL.Execute SQL
		if Err<>0 then
			MsgError "Error 10: " & Err.Description, true, true
			exit sub
		end if
		
		Response.Write "<script language=""JavaScript"">try{window.opener.window.thisForm.submit();window.close()}catch(e){}</script>"
		Response.End
		exit sub

	else	'Es un Brand nuevo
		
		SQL = "INSERT INTO Brand (Name, ShortName, indBaja, SiebelCode, Orden) " & _
		" VALUES (N'" & replace(LEFT(Name, 100), "'", "''") & "', " & _
		" N'" & LEFT(replace(ShortName, "'", ""), 17) & "', " & _
		" " & indBaja & ", " & _
		" N'" & replace(LEFT(SiebelCode, 50), "'", "''") & "', " & _
		" " & sOrden & ")"
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
dim IDBrand: IDBrand = request("IDBrand")
dim Name: Name = Request.Form("Name")
dim ShortName: ShortName = Request.Form("ShortName")
dim SiebelCode: SiebelCode = Request.Form("SiebelCode")
dim Orden: Orden = Request.Form("Orden")
dim indBaja: indBaja = Request.Form("indBaja")
dim FormName, FormDate




'-------------------------------------------------------------
'-------------------------------------------------------------
'Reconeix l'event --------------------------------------------
EventObject = request("EventObject")
EventParam1 = request("EventParam1")
EventParam2 = request("EventParam2")
if EventObject = "SaveBrand" then
	call SaveBrand_click()

end if

%>



<script language="JavaScript">
function SaveBrand(){
	if (thisForm.Name.value==''){alert('Write the Brand Name'); return false;}
	if (thisForm.ShortName.value==''){alert('Write the Brand Short Name'); return false;}
	if (thisForm.Orden.value==''){alert('Escriba el orden'); return false;}
	_fireEvent('SaveBrand','','')
}

</script>

<html>
<head>
	<title><%=IDM_BrandNewEditTitle%></title>
	<LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
</head>
<body topmargin=0 leftmargin=5>

<form action="BrandNewEdit.asp" method="post" name="thisForm">

	<table style="border:2 solid slateGray" border=0 cellpadding=0 cellspacing=0 width="100%">
	<tr>
		<td bgcolor=slateGray width=35></td>
		<td bgcolor=slateGray style="padding-left:5" align="left">
			<FONT class="font20"><font color=white><STRONG><%=IDM_BrandNewEditTitle%></STRONG></FONT>
		</td>
		<td bgcolor=slateGray align=right>
			<%'Botons de guardar, cancel·lar%>
			<input alt="Close" type="image" style="border:0" src="images/iCancel.gif" value="Close" onclick="window.close();return false;">
			<input alt="Save" type="image" style="border:0" src="images/iSave.gif" value="Save" onclick="SaveBrand();return false;">
		</td>
	</tr>
	</table>

<br>


	<%'Están editando o creando un nuevo usuario
		if IDBrand<>"" then	'Están EDITANDO
			SQL = "select b.*, f.Name AS [FormName], CONVERT(varchar,f.DateFrom,103) AS [FormDate] " & _
			" FROM Brand b " & _
			" LEFT JOIN Form f ON b.idForm = f.idForm " & _
			" WHERE b.IDBrand=" & IDBrand
			
			rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
			if not rst.EOF then
				Name = rst("Name")
				ShortName = rst("ShortName")
				SiebelCode = rst("SiebelCode")
				indBaja = rst("indBaja")
				Orden = rst("Orden")
				FormName = rst("FormName")
				FormDate = rst("FormDate")
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
					<td	width=70>
  						<font class="font11">
  							<%=IDM_Brand%>
  						</font>
					</td>
					<td colspan=3>
						<input style="width:200" class="recuadro" maxlength=100 name="Name" value="<%=Name%>">
						<input type="hidden" name="IDBrand" id="IDBrand" value="<%=IDBrand%>">
					</td>
				</tr>
				<tr>
					<td>
  					  <font class="font11"><%=IDM_ShortName%></font>
					</td>
					<td colspan=3>
						<input style="width:200" class="recuadro" name="ShortName" maxlength=17 value="<%=ShortName%>">
					</td>
				</tr>
				<tr>
					<td>
  					  <font class="font11"><%=IDM_BrandCode%></font>
					</td>
					<td colspan=3>
						<input style="width:200" class="recuadro" name="SiebelCode" maxlength=50 value="<%=SiebelCode%>">
					</td>
				</tr>
				<tr>
					<td	valign="top" width=70>
  					  <font class="font11"><%=IDM_Orden%></font>
					</td>
					<td valign="top">
						<input style="width:200" class="recuadro" name="Orden" maxlength=5 value="<%=Orden%>">
					</td>
				</tr>
				<tr>
				    <td>
				      <font class="font11"><%=IDM_Deleted%></font>
				    </td>
				    <td>
					    <input type="checkbox" class="recuadro" name="indBaja" <%if indBaja<>0 then%> checked <%end if%>>
				    </td>
				</tr>
				<%if FALSE then %>
				    <tr>
					    <td>
  					      <font class="font11"><%=IDM_Order%></font>
					    </td>
					    <td>
						    <input size=3 class="recuadro" name="NOrder" value="<%=NOrder%>">
					    </td>
				    </tr>
				<%end if%>
				<%if FALSE then%>
					<tr>
						<td><font class="font11"><%=IDM_Image%></td>
						<td colspan=3>
							<INPUT type="hidden" name="TargetURL" value="<%=Application("URLINTRA") & "/Docs"%>">
							<INPUT type="hidden" name="FFile">
							<INPUT TYPE=checkbox name="subirFichero" onclick="if (this.checked) {thisForm.borrarFichero.checked=true}">
							<INPUT TYPE="file" style="width:180;" name="file1">
						</td>
					</tr>
					<tr>
						<td></td>
						<td colspan=3>
							<INPUT TYPE=checkbox name="borrarFichero" onclick="if (subirFichero.checked) {this.checked=true}"> 
							<font class="font11">Borrar Fichero
						</td>
					</tr>
				<%end if%>
				<%if IDBrand <> "" then %>
				    <tr>
				        <td><font class="font11"><%=IDM_Form %></font></td>
				        <td><font class="font11">
				            <%if FormName<>"" then%>
    				            <%=FormName & " (" & FormDate & ")" %>
    				        <%end if %>
				            </font>
				        </td>
				    </tr>
				<%end if %>
			</table>
		</td></tr>
		</table>
				


<!-- #include file = "include/EventFunctions2.asp" -->



</form>

<script language="JavaScript">window.focus();</script>

</body>
</html>

<%Response.Flush%>
