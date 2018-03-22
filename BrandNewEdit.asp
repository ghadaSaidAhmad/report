<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
	Response.Buffer=true
	%>


<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/EventFunctions1.asp" -->
<!-- #include file = "ClassBrand.asp" -->
<%
	RecoverSQLConnection()
	RecoverSession(true) 
%>
<!-- #include file = "include/Idioma.asp" -->

<%
dim rst, SQL
dim b

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
		
	    set b = getBrand(IDBrand)
	    b.Name = Name
	    b.ShortName = ShortName
	    b.SiebelCode = SiebelCode
	    b.Orden = Orden
	    b.indBaja = indBaja
	    
	    for iSubcat = 0 to 9
	        b.arrNShops(iSubcat) = arrNShops(iSubcat)
	    next
	    
	    saveBrand(b)
	    
		if Err<>0 then
			MsgError "Error 10: " & Err.Description, true, true
			exit sub
		end if
		Err.Clear
		
		Response.Write "<script language=""JavaScript"">try{window.opener.window.thisForm.submit();window.close()}catch(e){}</script>"
		Response.End
		exit sub

	else	'Es un Brand nuevo
		
		
		set b = new Brand
		b.Name = Name
	    b.ShortName = ShortName
	    b.SiebelCode = SiebelCode
	    b.Orden = Orden
	    b.indBaja = indBaja

	    saveBrand(b)
		
		if Err<>0 then
			MsgError "Error 20: " & Err.Description, true, true
			exit sub
		end if
		Err.Clear
		
		Response.Write "<script language=""JavaScript"">try{window.opener.window.thisForm.submit();window.close()}catch(e){}</script>"
		Response.End
		exit sub
	end if
	
End Sub


'' Borra el descriptivo de la subcategoría en la marca y todos los valores en las Activity para esta marca e ID de subcategoría (0..9)
Sub deleteSubcat_click(id)
dim SQL
    
    on error resume next
    id = CInt(id)
	if Err<>0 then
		MsgError "Error 30: " & Err.Description, true, true
		exit sub
	end if
	Err.Clear
    
    SQL = "UPDATE Brand SET NShops" & id & " = '' WHERE IDBrand = " & IDBrand
    ObjConnectionSQL.Execute SQL
	if Err<>0 then
		MsgError "Error 40: " & Err.Description, true, true
		exit sub
	end if
	Err.Clear
    
    
    SQL = "UPDATE Activity SET NShops" & id & " = NULL WHERE IDBrand = " & IDBrand
    ObjConnectionSQL.Execute SQL
	if Err<>0 then
		MsgError "Error 50: " & Err.Description, true, true
		exit sub
	end if
	Err.Clear
    
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

dim arrNShops(9)
dim iSubcat

for iSubcat = 0 to 9
    arrNShops(iSubcat) = Request.Form("NShops" & iSubcat)
next


'-------------------------------------------------------------
'-------------------------------------------------------------
'Reconeix l'event --------------------------------------------
EventObject = request("EventObject")
EventParam1 = request("EventParam1")
EventParam2 = request("EventParam2")
select case EventObject 
    case "SaveBrand" call SaveBrand_click()
    case "deleteSubcat" call deleteSubcat_click(EventParam1)

end select

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
		    
		    on error resume next
		    set b = getBrand(IDBrand)
		    Name = b.Name
	        ShortName = b.ShortName
	        SiebelCode = b.SiebelCode
	        Orden = b.Orden
	        indBaja = b.indBaja
	        FormName = b.FormName
	        FormDate = b.FormDate
	        
	        for iSubcat = 0 to 9
	            arrNShops(iSubcat) = b.arrNShops(iSubcat)
	        next
	        
		    if Err.number<>0 then
				MsgError Err.Description, true, true
		    end if
		    on error goto 0

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
				<%if FormName <> "" then %>
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
				
				<%if IDBrand <> "" then %>
				
				    <tr>
				        <td colspan="10"><hr /></td>
				    </tr>
				    
				    <%
				    for iSubcat = 0 to 9%>
				        <tr>
				            <td><font class="font11"><%=IDM_Subcategory & " " & (iSubcat + 1)%> </font></td>
					        <td colspan=3>
						        <input style="width:100px" class="recuadro" name="NShops<%=iSubcat %>" maxlength=50 value="<%=arrNShops(iSubcat)%>">
						        <input type="button" style="width:80px" value="Borrar" onclick="_fireConfirm('deleteSubcat', '<%=iSubcat %>', '', '<%=IDM_JS_ConfirmDeleteSubcat %>');" />
					        </td>
				        </tr>
				    <%next%>
				    
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
