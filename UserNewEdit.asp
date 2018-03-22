<%@language=VBScript%>
<%Option Explicit
    Response.Expires=0
	Response.Buffer=true
	%>


<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/EventFunctions1.asp" -->

<%
dim rst, tmp, SQL, temp
set rst = CreateObject("ADODB.Recordset")%>
<%
	RecoverSQLConnection()
	RecoverSession(true) 
%>
<!-- #include file = "include/Idioma.asp" -->


<%
if not isAdmin() then
	msgError "You are not allowed to view this information", true, true
end if
%>

<%

'-------------------------------------------------------------
'-------------------------------------------------------------
Sub SearchUser_click(S)
	SearchUser = S
End Sub

Sub SaveUser_click()
dim ArrGroups, g, rstUs, strSql
	
	Set rstUs = Server.CreateObject("ADODB.Recordset")
	if NewUser="" then	'No es un usuario nuevo

		if IDNewGroup<>"" then
			
			ArrGroups = Split(IDNewGroup, ",")
			For Each g In ArrGroups

				SQL = "SELECT * from UserGroup WHERE IDEmpleado=" & IDUser & " AND IDGroup=" & g
				set temp = ObjConnectionSQL.Execute(SQL)
				if temp.Eof then
					
					SQL = "INSERT INTO UserGroup (IDEmpleado, IDGroup) VALUES (" & IDUser & ", " & g & ")"
					ObjConnectionSQL.Execute SQL
					
				end if
				Set temp = Nothing
				
			Next
			
		else
			SQL = "UPDATE Users SET Idioma='" & Idioma & "' WHERE IDEmpleado=" & IDUser
			ObjConnectionSQL.Execute SQL
			
			Response.Write "<script language=""JavaScript"">try{window.opener.window.thisForm.submit();}catch(e){}</script>"
			'Response.End
		end if
	else	'Es un usuario nuevo
		
		SQL = "select e.ApellidosNombre, e.IDEmpleado, e.NTUser, u.IDEmpleado AS Existe " & _
		" FROM EmpleadosGlobal e " & _
		" LEFT JOIN Users u  ON u.IDEmpleado=e.IDEmpleado " & _
		" WHERE e.IDEmpleado=" & IDNewUser

		rstUs.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
		if not rstUs.EOF then
			if not isnull(rstUs("Existe")) then
				MsgError "Error 10: " & IDM_UserNewEditErr10, true, true
				exit sub
			end if
			
			if isNull(rstUs("NTUser")) OR rstUs("NTUser")="" then
			    MsgError "Falta el NTUser del empleado en la gestión de empleados. Póngase en contacto con la persona que debe gestionarlo. Hasta que no se solucione, el empleado no podrá acceder a la aplicación.", false, false
			end if
			
		else
			MsgError "Error 20: " & IDM_UserNewEditErr20, true, true
			exit sub
		end if
		FullName = rstUs("ApellidosNombre")
		rstUs.Close
		
		strSql = "insert into Users (IDEmpleado, Idioma)values(" & IDNewUser & ", '" & Idioma & "')"
		ObjConnectionSQL.Execute(strSql)
		
		if IDNewGroup<>"" then

			ArrGroups = Split(IDNewGroup, ",")
			For Each g In ArrGroups
				
				SQL = "insert into UserGroup (IDEmpleado,IDGroup) values (" & IDNewUser & ",'" & g & "')"
				ObjConnectionSQL.Execute(SQL)
				
			Next


		end if
		
		IDUser = IDNewUser
		Response.Write "<script language=""JavaScript"">try{window.opener.window.thisForm.submit();}catch(e){}</script>"
		'Response.End
	end if
	
	NewUser=""
	
End Sub

Sub DeleteUserGroup_click(DelIDUser,DelIDGroup)
	SQL = "DELETE FROM UserGroup WHERE IDEmpleado=" & DelIDUser & " and IDGroup=" & DelIDGroup
	ObjConnectionSQL.Execute SQL
End Sub

Sub AddGroups_click()
    dim ArrGroups, g, temp

	ArrGroups = Split(IDNewGroup, ",")
	For Each g In ArrGroups
		
		SQL = "SELECT * from UserGroup WHERE IDEmpleado=" & request("IDUser") & " AND IDGroup=" & g
		set temp = ObjConnectionSQL.Execute(SQL)
		if temp.Eof then
					
			SQL = "INSERT INTO UserGroup (IDEmpleado, IDGroup) VALUES (" & request("IDUser") & ", " & g & ")"
			ObjConnectionSQL.Execute SQL
					
		end if
		Set temp = Nothing
				
	Next
	
End Sub

Sub DelGroup_click(DelIDGroup)
	on error resume next
	dim strSql
	strSql = "delete from UserGroup where IDEmpleado=" & request("IDUser") & " and IDGroup='" & DelIDGroup & "'"
	ObjConnectionSQL.Execute strSql
End Sub



dim strErrorMsg
dim SearchUser


'-------------------------------------------------------------
'-------------------------------------------------------------
'Manté els valors en el formulari ----------------------------
dim IDUser: IDUser = request("IDUser")
dim NewUser: NewUser = request("NewUser")
dim IDNewUser: IDNewUser = request("IDNewUser")
dim IDNewGroup: IDNewGroup = request("IDNewGroup")
dim FullName: FullName = request("FullName")
Idioma = Request.Form("Idioma")

SearchUser = request("SearchUser")


%>


<%
'-------------------------------------------------------------
'-------------------------------------------------------------
'Reconeix l'event --------------------------------------------
EventObject = request("EventObject")
EventParam1 = request("EventParam1")
EventParam2 = request("EventParam2")
if EventObject = "SearchUser" then
	call SearchUser_click(EventParam1)
elseif EventObject = "Search" then
	call Search_click()
elseif EventObject = "SaveUser" then
	call SaveUser_click()
elseif EventObject = "AddGroups" then
	call AddGroups_click()
elseif EventObject = "DelGroup" then
	call DelGroup_click(EventParam1)
elseif EventObject = "DeleteUserGroup" then
	call DeleteUserGroup_click(EventParam1,EventParam2)

end if

%>



<script language="JavaScript">
function SaveUser(){
	<%if NewUser="" then%>
		if (thisForm.IDUser.value==''){
			alert('<%=IDM_SelectUser%>');}
	<%else%>
		if (thisForm.IDNewUser.value==''){
			alert('<%=IDM_SelectUser%>');}
	<%end if%>
	_fireEvent('SaveUser','','')
}

</script>


<html>
<head>
	<title><%=IDM_UserNewEditTitle%></title>
	<LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
</head>
<body topmargin=0 leftmargin=5>

<form action="UserNewEdit.asp" method="post" name="thisForm">

    <table style="width:100%;height:40px;background-image:url('images/Grad5.gif'); ">
        <tr>
            <td valign="middle" style="padding-left:10px;">
                <font class="wopenTitle"><%=IDM_UserNewEditTitle%></font>
            </td>
            <td align="right" width="180px;">
			    <input class=button alt="Save" type="button" style="width:55px;" value="<%=IDM_Save%>" onclick="SaveUser();return false;">
			    <input class=button alt="Close" type="button" style="width:55px;" value="<%=IDM_Close%>" onclick="window.close();return false;">
            </td>
        </tr>
    </table>

    <br>


	<%'Están editando o creando un nuevo usuario
		if NewUser="" then	'Están EDITANDO
			SQL = "select e.ApellidosNombre AS FullName, u.Idioma " & _
			" FROM Users u " & _
			" INNER JOIN EmpleadosGlobal e ON u.IDEmpleado=e.IDEmpleado " & _
			" WHERE u.IDEmpleado=" & IDUser
			'printDebug "SQL", SQL
			rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
			if not rst.EOF then
				FullName = rst("FullName")
				Idioma = rst("Idioma")
			else
				MsgError "", true, true
			end if
			rst.Close
		end if%>
		
		

		<table border="0" width="100%" cellpadding=0 cellspacing=0>
		<tr><td>
		  <table border=0 width="100%" cellpadding=2 cellspacing=0>
				<tr><td colspan=4></td></tr>
				<tr>
					<td	valign="top" width=90 class="fieldheader">
						<%
						if IDUser="" then
  							Response.Write IDM_NewUser
  						else
  							Response.Write IDM_User
  						end if
  						%>
					</td>
					<td valign="top" width=220 >
						
						<%if IDUser="" then%>
							<SELECT id=IDNewUser name=IDNewUser style="width:200">
								<%	dim strSearchUser: strSearchUser = SearchUser
									if strSearchUser<>"" then
										if strSearchUser="*" then strSearchUser="%"
										strSearchUser = " WHERE indBaja=0 AND (c.NTUser LIKE '%" & strSearchUser & "%' OR c.ApellidosNombre LIKE '%" & strSearchUser & "%')"
									else
										strSearchUser = " WHERE indBaja=0 AND c.ApellidosNombre = ''"
									end if
									strSearchUser = strSearchUser & " AND c.IDEmpleado NOT IN (SELECT IDEmpleado FROM Users) "

									SQL = "SELECT c.IDEmpleado, c.NTUser, c.ApellidosNombre AS FullName " & _
									" FROM EmpleadosGlobal c " & _
									" LEFT JOIN Users u ON c.IDEmpleado=u.IDEmpleado " & strSearchUser & _
									" ORDER BY c.ApellidosNombre"
									response.write SQL
									rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
									While not rst.Eof
										IF rst("IDEmpleado") <> "" THEN
											Response.Write "<OPTION VALUE=""" & rst("IDEmpleado") & """"
											'if rst("IDEmpleado") = IDUser then
											'	Response.Write " selected "
											'end if
											if rst("FullName")="" or isnull(rst("FullName")) then
												Response.Write ">" & rst("NTUser") & "</OPTION>"
											else
												Response.Write ">" & rst("FullName") & "</OPTION>"
											end if
										END IF
										rst.MoveNext
									Wend
									rst.Close
								%>
								</SELECT>
								<a onclick="
									if (srch = prompt('Search User (Type NTUser or Name)','*')){
										_fireEvent('SearchUser',srch,'');
									}
									return false;
								" href=""><img align=middle border=0 alt="Search User" src="images/ilupa.gif"></a>
						<%else%>
							<input readonly style="width:200" class="textfield" name="FullName" id="FullName" value="<%=FullName%>">
							<input type="hidden" name="IDUser" id="IDUser" value="<%=IDUser%>">
						<%end if%>
					</td>
				</tr>
				<tr>
					<td	valign="top" width="70" class="fieldheader"><%=IDM_Idioma%></td>
					<td>
						<select id="Idioma" name="Idioma" style="width:200" >
							<%SQL = "SELECT IDIdioma AS ID, Description FROM Idioma ORDER BY Description"
							rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
							While not rst.Eof
								Response.Write "<OPTION VALUE=""" & rst("ID") & """"
								if Idioma<>"" then
									if ucase(rst("ID")) = ucase(Idioma) then
										Response.Write " selected "
									end if
								end if
								Response.Write ">" & rst("Description") & "</OPTION>"
								rst.MoveNext
							Wend
							rst.Close%>
						</select>
					</td>
				</tr>
				<tr>
					<td	valign="top" width=70 class="fieldheader"><%=IDM_Group%>
					</td>
					<td valign="top">
						<%if NewUser="" then%>
							<SCRIPT LANGUAGE="JavaScript">
								function AddGroups(){
									if (thisForm.IDNewGroup.selectedIndex>=0) {
										_fireEvent('AddGroups','','');
									}else{
									    alert('<%=IDM_SelectGroup %>');
									}
								}
							</SCRIPT>
							&nbsp;<a onclick="AddGroups();return false;" href=""><img src="images/add.gif" border="0" />&nbsp;<font class="font10"><%=IDM_AddToGroup%></font></a>
							<br />
						<%end if%>
						<SELECT id=IDNewGroup name=IDNewGroup style="width:200;height:80;" multiple >
							<%
							SQL = "SELECT IDGroup, Description " & _
							" FROM Groups " & _
							" WHERE Description NOT LIKE 'Reserved %' " & _
							" AND IDGroup NOT IN (" & _
							" SELECT IDGroup FROM UserGroup WHERE IDEmpleado='" & IDUser & "' " & _
							" ) " & _
							" ORDER BY Description"
							rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
							While not rst.Eof
								Response.Write "<OPTION VALUE=""" & rst("IDGroup") & """>" & rst("Description") & "</OPTION>"
								rst.MoveNext
							Wend
							rst.Close%>
						</SELECT>
					</td>
				</tr>
				
				<%if NewUser="" then%>
				    <tr><td colspan="6" style="border-bottom:2 solid black;"><font class="font8">&nbsp;</font></td></tr>
					<tr>
						<td colspan="6" class="fieldheader">
						<font class="font11"><b><%=IDM_GroupsAssTo%><br /><br /></b></font>
						
							<TABLE BORDER="0" width="100%" cellpadding=0 cellspacing=0>
							<%
							SQL = "select g.IDGroup,g.Description " & _
							" FROM UserGroup us " & _
							" INNER JOIN Groups g ON g.IDGroup=us.IDGroup " & _
							" WHERE us.IDEmpleado=" & IDUser & " ORDER BY g.Description"
							rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
							if not rst.EOF then
								While not rst.EOF%>
									<tr>
										<td width=30 align="left"><a href="javascript:if (confirm(('Click OK to continue. Click Cancel to abort.'))) _fireEvent('DelGroup','<%=rst("IDGroup")%>','');"><img src="images/delete.png" border="0" alt="<%=IDM_Delete%>"></a></td>
										<td align="left" class="font10"><%=rst("Description")%></td>
									</tr>
								<%rst.MoveNext
								Wend
							else%>
								<tr>
									<td colspan="2"><font face=Arial size=1 color=red><%=IDM_NoRecordsFound%></font></td>
								</tr>
							<%end if%>
						</td>
					</tr>
			<%end if%>	


<%


'-------------------------------------------------------------
'-------------------------------------------------------------
'Manté els valors en el formulari %>
<INPUT type="hidden" id=NewUser name=NewUser value="<%=NewUser%>">



<INPUT type="hidden" id=SearchUser name=SearchUser value="<%=SearchUser%>">



<!-- #include file = "include/EventFunctions2.asp" -->



</form>


</body>
</html>

<%Response.Flush%>
