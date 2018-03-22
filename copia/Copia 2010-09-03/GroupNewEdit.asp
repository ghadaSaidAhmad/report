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
dim rst, SQL
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


'-------------------------------------------------------------
'-------------------------------------------------------------
Sub DeleteUserGroup_click(DelIDUser,DelIDGroup)
    dim SQL
	SQL = "DELETE FROM UserGroup WHERE IDEmpleado=" & DelIDUser & " and IDGroup=" & DelIDGroup
	ObjConnectionSQL.Execute SQL
End Sub

Sub SaveGroup_click()
dim rstGr
	Set rstGr = Server.CreateObject("ADODB.Recordset")
	if IDGroup<>"" then
		if NewGroup="" then	'No es un group nuevo
			SQL = "SELECT * FROM Groups WHERE IDGroup=" & IDGroup
			rstGr.Open SQL, ObjConnectionSQL, adOpenDynamic, adLockOptimistic
		end if
	else	'Es un group nuevo
		SQL = "SELECT * FROM Groups WHERE IDGroup=0"
		rstGr.Open SQL, ObjConnectionSQL, adOpenDynamic, adLockOptimistic
		rstGr.AddNew
	end if
	rstGr("Description") = Description
	rstGr("Observations") = null
	if Observations<>"" then rstGr("Observations") = Observations

	rstGr.Update
    rstGr.Close
    
    if IDGroup="" then
        SQL = "SELECT MAX(IDGroup) FROM Groups"
        rstGr.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        IDGroup = rstGr.Fields(0)
        rstGr.close
    end if
    
	Response.Write "<script language=""JavaScript"">try{window.opener.window.thisForm.submit();}catch(e){}</script>"
	'Response.End

End Sub

Sub DelGroup_click(DelIDGroup)
	on error resume next
	strSql = "delete from UserGroup where IDEmpleado=" & request("IDUser") & " and IDGroup='" & DelIDGroup & "'"
	ObjConnectionSQL.Execute strSql
End Sub



dim strErrorMsg
dim SearchUser


'-------------------------------------------------------------
'-------------------------------------------------------------
'Manté els valors en el formulari ----------------------------
dim IDGroup: IDGroup = request("IDGroup")
dim NewGroup: NewGroup = request("NewGroup")
dim Description: Description = request("Description")
dim Observations: Observations = request("Observations")



'-------------------------------------------------------------
'-------------------------------------------------------------
'Reconeix l'event --------------------------------------------
if EventObject = "SearchUser" then
	call SearchUser_click(EventParam1)
elseif EventObject = "SaveGroup" then
	call SaveGroup_click()
elseif EventObject = "DeleteGroup" then
	call DeleteGroup_click(EventParam1)
elseif EventObject = "AddGroup" then
	call AddGroup_click(EventParam1)
elseif EventObject = "DelGroup" then
	call DelGroup_click(EventParam1)
elseif EventObject = "DeleteUserGroup" then
	call DeleteUserGroup_click(EventParam1,EventParam2)

end if

%>


<script language="JavaScript">
function SaveGroup(){
	if (thisForm.Description.value==''){
		alert('<%=IDM_WriteGroupDescr%>');}
	else{
		_fireEvent('SaveGroup','')}
	
}

function DeleteGroup(IDGr){
	if (confirm(("<%=IDM_JS_ClickOKCancel%>")))
		_fireEvent('DeleteGroup',IDGr)
}
function DeleteUserGroup(IDUser,IDGroup){
	if (confirm(("<%=IDM_JS_ClickOKCancel%>")))
		_fireEvent('DeleteUserGroup',IDUser,IDGroup)
}
</script>


<html>
<head>
<title><%=IDM_GroupNewEditTitle%></title>
	<LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
</head>
<body topmargin=0 leftmargin=5>

<form action="GroupNewEdit.asp" method="post" name="thisForm">

    <table style="width:100%;height:40px;background-image:url('images/Grad5.gif'); ">
        <tr>
            <td valign="middle" style="padding-left:10px;">
                <font class="wopenTitle"><%=IDM_GroupNewEditTitle%></font>
            </td>
            <td align="right" width="180px;">
			    <input class=button alt="Save" type="button" style="width:55px;" value="<%=IDM_Save%>" onclick="SaveGroup();return false;">
			    <input class=button alt="Close" type="button" style="width:55px;" value="<%=IDM_Close%>" onclick="window.close();return false;">
            </td>
        </tr>
    </table>


<br>
		
<%if NewGroup="" then	'Están EDITANDO
	SQL = "SELECT * FROM Groups WHERE IDGroup='" & IDGroup & "'"
	rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
	if not rst.EOF then
		Description = rst("Description")
		Observations = rst("Observations")
	else
		MsgError "", true, true
	end if
	rst.Close
end if%>

<table border="0" width="100%" cellpadding=0 cellspacing=0>
<tr><td>
  <table border=0 width="100%" cellpadding=2 cellspacing=0>
		<tr><td colspan=4>
		</td></tr>
		<tr>
			<td	valign="top" width=120 class="fieldheader"><%=IDM_GroupDescr%></td>
			<td valign="top">
				<input maxlength=50 class="textfield" type="text" style="width:200" name="Description" id="Description" value="<%=Description%>">
				<input type="hidden" name="IDGroup" id="IDGroup" value="<%=IDGroup%>">
			</td>
		</tr>
		<tr>
			<td	valign="top" width=120 class="fieldheader"><%=IDM_Observations%></td>
			<td valign="top">
				<textarea class="textfield" style="width:200;height:80" name="Observations" id="Observations"><%=Observations%></textarea>
			</td>
		</tr>



		<%if IDGroup<>"" then%>
			<tr>
    		    <tr><td colspan="6" style="border-bottom:2 solid black;"><font class="font8">&nbsp;</font></td></tr>
				<tr>
				<td colspan="6" class="fieldheader">
					<font class="font11"><b><%=IDM_UsersAssTo%><br /><br /></b></font>
					<TABLE BORDER="0" width="100%" cellpadding=0 cellspacing=0>
						<%SQL = "select u.IDEmpleado, e.ApellidosNombre AS FullName " & _
						" FROM UserGroup us " & _
						" INNER JOIN Users u ON u.IDEmpleado=us.IDEmpleado " & _
						" INNER JOIN EmpleadosGlobal e ON e.IDEmpleado=u.IDEmpleado " & _
						" WHERE us.IDGroup=" & IDGroup & " ORDER BY u.FullName"
						rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
						if not rst.EOF then
							While not rst.EOF%>
								<tr>
									<td width=30 align="left"><a href="" onclick="DeleteUserGroup('<%=rst("IDEmpleado")%>',<%=IDGroup%>);return false;"><img src="images/delete.png" alt="<%=IDM_DeleteUserFrom%> '<%=rst("FullName")%>'" border=0></a></td>
									<td align="left" class="font10"><%=rst("Fullname")%></td>
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
  </table>
</td></tr></table>


<!-- #include file = "include/EventFunctions2.asp" -->


</form>


</body>
</html>

<%Response.Flush%>