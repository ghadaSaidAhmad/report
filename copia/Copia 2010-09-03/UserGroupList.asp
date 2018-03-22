<%@language=VBScript%>
<%Option Explicit
    Response.Expires=0
	Response.Buffer=true
	%>

<!-- #include file = "include/ListFunctions.asp" -->
<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/EventFunctions1.asp" -->

<!-- #include file = "ClassInclude.asp" -->

<%
dim rst, SQL, strSql, strSelected, strOrder, vecWidth
set rst = CreateObject("ADODB.RecordSet")%>


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
Sub SearchUser_click(S)
	SearchUser = S
End Sub

Sub DeleteUser_click(DelIDUser)
	
	'Set rstDel = Server.CreateObject("ADODB.Recordset")
	'SQL = "SELECT * FROM Call WHERE IDStatus<>999 AND  '" & DelIDUser & "'"
	'rstDel.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
	
	on error resume next
	SQL = "DELETE FROM Users WHERE IDEmpleado=" & DelIDUser
	ObjConnectionSQL.Execute SQL
	if Err<>0 then
		MsgError "Error 10: " & IDM_UserGroupListErr10, false, false
	else

		SQL = "DELETE FROM UserGroup WHERE IDEmpleado=" & DelIDUser
		ObjConnectionSQL.Execute SQL
	end if
	
End Sub

Sub DeleteUserGroup_click(DelIDUser,DelIDGroup)
	SQL = "DELETE FROM UserGroup WHERE IDEmpleado=" & DelIDUser & " and IDGroup=" & DelIDGroup
	ObjConnectionSQL.Execute SQL
End Sub

Sub DeleteGroup_click(DelIDGroup)
	dim rstDel, SQL
	if cint(DelIDGroup)>10 then
		Set rstDel = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT g.IDGroup " & _
		" FROM Groups g " & _
		" INNER JOIN UserGroup ug ON g.IDGroup=ug.IDGroup WHERE g.IDGroup=" & DelIDGroup
		rstDel.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
		if rstDel.EOF then
			rstDel.Close
			on error resume next
			SQL = "DELETE FROM Groups WHERE IDGroup=" & DelIDGroup
			ObjConnectionSQL.Execute SQL
			if Err<>0 then
				MsgError "Error 20: " & IDM_UserGroupListErr20, false, false
			end if
		else
			rstDel.Close
			MsgError IDM_UserGroupListErr30, false, false
		end if
	else
		msgError IDM_UserGroupListErr40, false, false
	end if
	
End Sub

Sub SearchUsers_click()
	
	Pagina = "1"
	Order = ""
	
End Sub

Sub Search_click()
	'Srch_TxtSearch = TxtSearch
	SearchAssignedTo = AssignedTo
	SearchGroup = SGroup
	Pagina = "1"
End Sub

dim strErrorMsg
dim SearchUser


'-------------------------------------------------------------
'-------------------------------------------------------------
'Manté els valors en el formulari ----------------------------
dim IDUser: IDUser = request("IDUser")
dim IDGroup: IDGroup = request("IDGroup")
dim Description: Description = request("Description")
dim Pestana: Pestana = request("Pestana")
dim NewUser: NewUser = request("NewUser")
dim NewGroup: NewGroup = request("NewGroup")
dim IDNewGroup: IDNewGroup = request("IDNewGroup")
dim IDNewUser: IDNewUser = request("IDNewUser")
dim FullName: FullName = request("FullName")
dim EMail: EMail=request("EMail")
dim AssignedTo: AssignedTo = request("assignedTo")
dim SearchAssignedTo: SearchAssignedTo = request("SearchAssignedTo")
dim SGroup: SGroup = request("SGroup")
dim SearchGroup: SearchGroup = request("SearchGroup")
if Pestana="" then
	Pestana="1"
end if
dim Pagina: Pagina = request("Pagina")
dim Order: Order = request("Order")

SearchUser = request("SearchUser")
dim SearchClient: SearchClient = request("SearchClient")

dim EstadoASP: EstadoASP = request("EstadoASP")




'-------------------------------------------------------------
'-------------------------------------------------------------
'Reconeix l'event --------------------------------------------
if EventObject = "SearchUser" then
	call SearchUser_click(EventParam1)
elseif EventObject = "Search" then
	call Search_click()
elseif EventObject = "DeleteUser" then
	call DeleteUser_click(EventParam1)
elseif EventObject = "DeleteGroup" then
	call DeleteGroup_click(EventParam1)
elseif EventObject = "DeleteClient" then
	call DeleteClient_click(EventParam1)
elseif EventObject = "DelGroup" then
	call DelGroup_click(EventParam1)
elseif EventObject = "DeleteUserGroup" then
	call DeleteUserGroup_click(EventParam1,EventParam2)

elseif EventObject = "Siguiente" then
	call Siguiente_click()
elseif EventObject = "Anterior" then
	call Anterior_click()
elseif EventObject = "Primero" then
	call Primero_click()
elseif EventObject = "Ultimo" then
	call Ultimo_click()

end if





if Pagina = "" then
	Pagina = "1"
end if

%>


<html>
<head>
<title><%=IDM_UserGroupListTitle%></title>

	<link rel="StyleSheet" href="include/style.css" type="text/css" />

    <script type="text/javascript">
        function editElement(val , tipo){
	        if (tipo=="User"){
		        window.open('UserNewEdit.asp?IDUser=' + val,'User','top=200,left=300,width=400,height=400,scrollbars');
	        }
	        else if (tipo=="Group"){
		        window.open('GroupNewEdit.asp?IDGroup=' + val,'Group','top=200,left=300,width=400,height=300,scrollbars');
	        }
        }

        function DeleteElement(val, tipo){
	        if (tipo=="User"){
		        if (confirm('<%=IDM_JS_ClickOKCancel%>')){
			        _fireEvent('DeleteUser',val,'');}
	        }else if (tipo=="Group"){
		        if(val>10){
			        if (confirm('<%=IDM_JS_ClickOKCancel%>'))
				        _fireEvent('DeleteGroup',val,'');
		        }
		        else{
			        alert('<%=IDM_UserGroupListErr40%>');
		        }
	        }
        }


        function copyElement(val, tipo){
	        if (tipo=="User"){
		        window.open('UserCopy.asp?ID=' + val,'CPY','width=400,height=250');
	        }
        }

        function DeleteUser(IDUs){
	        if (confirm(("<%=IDM_JS_ClickOKCancel%>")))
		        _fireEvent('DeleteUser',IDUs,'')
        }


        function DeleteGroup(IDGr){
	        if (confirm(("<%=IDM_JS_ClickOKCancel%>")))
		        _fireEvent('DeleteGroup',IDGr,'')
        }

        function ResetFields(){
	        thisForm.AssignedTo.value='';
	        thisForm.SGroup.selectedIndex=0;		
	        }

        function submitenter(myfield,e,evento)
        {
	        var keycode;
	        if (window.event) keycode = window.event.keyCode;
	        else if (e) keycode = e.which;
	        else return true;

	        if (keycode == 13){
		        if (evento=="Search"){
			        _fireEvent('Search','click','')
		        }
		        else{
		           _fireEvent(evento,'','');
		          }
	           return false;
	        }
	        else
	           return true;
        }

    </script>

</head>
<body class="BODY_MAIN" style="background-image:url('images/background.jpg');background-repeat:no-repeat;">


<form action="UserGroupList.asp" method="post" name="thisForm">


    <%' TABLA TOP%>
    <TABLE width="100%" ID="TBL_TOP" cellpadding=0 cellspacing=0 class=topTable background="images/a3.gif">
        <TR>
            <TD valign=top width=380 style="padding:10px;">
                <font class=fontTitleTop>
                    <%=IDM_MAINTITLE1 %>
                </font>
                <font class=fontTitleTop2>
                    <br />
                    <%=IDM_MAINTITLE2 %>
                </font>
            </TD>
        </TR>
    </TABLE>    


	<br><br>


<table border=0 cellpadding=0 cellspacing=0 width="600px" style="border:1 solid gray;" align=center ><tr height=400px><td align=center valign=top style="border-right:2px solid black;border-bottom:2px solid black;">

    <!-- #include file = "ClassMenuAdmin.asp" -->

    <table class="TB_TITLE" cellpadding=0 cellspacing=0 width="100%"><tr><td>
	    
	    <table border=0 cellpadding=0 cellspacing=0 width="100%">
	    <tr>
		    <td align="left" style="padding-left:5">
			    <img src="images/users.png" /><FONT class="font20">&nbsp;&nbsp;<STRONG><%=IDM_UserGroupListTitle%></STRONG></FONT>
		    </td>
		    <td align=right>
			    <%'Botons de guardar, cancel·lar%>
    			
			    <%if Pestana="1" then	'estamos en la pestaña de Usuarios%>
				    <input style="height:30px;cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" alt="<%=IDM_NewUser%>" type="image" src="images/add.png" value="New User" onclick="window.open('UserNewEdit.asp?NewUser=1','User','top=200,left=300,width=400,height=400,scrollbars');return false;">
				    <Input style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" alt="<%=IDM_Search%>" type=image src="images/search.png" value="<%=IDM_Search%>" onclick="_fireEvent('Search','click','');return false;">
			    </td>
			    </tr>
    			
			    <%elseif Pestana="2" then	'estamos en la pestaña de Grupos%>
				    <input style="height:30px;cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" alt="<%=IDM_NewGroup%>" type="image" src="images/add.png" value="New Group" onclick="window.open('GroupNewEdit.asp?NewGroup=1','Group','top=200,left=300,width=400,height=300,scrollbars');return false;">
				    <Input style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" alt="<%=IDM_Search%>" type=image src="images/search.png" value="<%=IDM_Search%>" onclick="_fireEvent('Search','click','');return false;">
			    </td>
			    </tr>
			    <%end if%>

	    </tr>
	    </table>
	    
	    <hr style="height:3px;" color=black />
	    
	    <table border=0 cellpadding=0 cellspacing=0 width="100%">
	    <tr>
			    <tr>
				    <td colspan="2">
				    <table>
			    <tr>
				    <TD valign=top><font class="font20"> <%=IDM_User%>&nbsp;</TD>
			    <TD>
    			
				    <input class="textfield" type=text name="AssignedTo" VALUE="<%=AssignedTo%>" width=20 onkeypress="return submitenter(this,event,'Search');">
    			
    			
				    <%if FALSE then%>
						    <SELECT name="AssignedTo" id="AssignedTo" style="width:200">
						    <option value=""></option>
								    <%
								    if SearchUser<>"" then
									    if SearchUser="*" then
										    valor ="%"
									    else
										    valor = SearchUser
									    end if
									    strSql = "SELECT em.IDEmpleado, em.ApellidosNombre AS FullName " & _
									    " FROM Users u " & _
									    " INNER JOIN EmpleadosGlobal em ON u.IDEmpleado=em.IDEmpleado " & _
									    " WHERE em.NTUser LIKE '%" & valor & "%' OR em.ApellidosNombre LIKE '%" & valor & "%'"
									    rst.Open strSql, ObjConnectionSQL, adOpenStatic, adLockReadOnly
									    while not rst.EOF%>
										    <option value="<%=rst("IDEmpleado")%>"
										    <%if AssignedTo<>"" then
											    if cint(AssignedTo) = cint(rst("IDEmpleado")) then
												    Response.Write " selected "
											    end if
										    end if%>							
										    ><%=rst("FullName")%></option>
										    <%rst.MoveNext
									    wend
									    rst.Close
								    else%>
									    <OPTION value="<%=AssignedTo%>"><%=AssignedToName%></OPTION>
								    <%end if%>
						    </SELECT>
						    <a onclick="if (srch = prompt('Search User (Type NTUser or Name)','*')){
											    _fireEvent('SearchUser',srch,'');
										    }
										    return false;" href=""><img align=middle border=0 alt="Search" src="images/ilupa.gif"></a>
				    <%end if%>
			    </td>
			    <td>&nbsp;</td>
			    <TD valign=top><font class="font20"> <%=IDM_Group%>&nbsp;</TD>
			    <TD>
				    <select name="SGroup" style="width:150">
					    <option value=""></option>
				    <%
				    strSql = "SELECT IDGroup, Description FROM Groups ORDER BY Description"
				    rst.Open strSql, ObjConnectionSQL, adOpenStatic, adLockReadOnly
				    while not rst.EOF
					    strSelected = " "
					    if SGroup<>"" then
						    if Cint(SGroup)=rst("IDGroup") then
							    strSelected = " SELECTED"
						    end if
					    end if
					    Response.Write "<option value='" & rst("IDGroup") & "'" & strSelected & ">" & rst("Description") & "</option>"
					    rst.MoveNext
				    wend
				    rst.Close
				    %>
				    </select>
			    </td>
			    </tr>
			    </table>
			    <td>
			    </tr>
	    </table>
    </table>

    <br>


    <table width=100% cellspacing=0>
	    <tr>
		    <td width=20 CLASS="PEST_ESPACIO"><font class="font12">&nbsp;</td>
		    <td onclick="location.href='UserGroupList.asp?Pestana=1';" width=120 CLASS="<%if Pestana="1" then%>PEST_SELEC<%else%>PEST_NOSELEC<%end if%>"><font class="font12"><b><%=IDM_Users%> <label id="TotalUsers"></label></td>
		    <td width=5 CLASS="PEST_ESPACIO">&nbsp;</td>
		    <td onclick="location.href='UserGroupList.asp?Pestana=2';" width=120 CLASS="<%if Pestana="2" then%>PEST_SELEC<%else%>PEST_NOSELEC<%end if%>"><font class="font12"><b><%=IDM_Groups%> <label id="TotalGroups"></label></td>
		    <td CLASS="PEST_ESPACIO">&nbsp;</td>
	    </tr>
    </table>
    <br>



    <%
    dim RegsPorPagina, aDatos, pag
    dim vec, iEstado
    %>

    <%if Pestana="1" then	'estamos en la pestaña de Usuarios%>

	    <%	if Order<>"" then
			    strOrder = " Order By " & Order
		    else
			    strOrder = " ORDER BY e.ApellidosNombre"
		    end if
    	
		    SQL = "select e.IDEmpleado AS [" & IDM_FLD_IDEmpleado & " ], e.ApellidosNombre + CASE WHEN e.NTUser IS NULL OR e.NTUser = '' THEN ' <font color=red>(Falta NTUser)</font>' ELSE '' END AS [" & IDM_FLD_NombreEmpleado & "], " & _
		    " CONVERT(VARCHAR(100),e.EMail) AS [" & IDM_FLD_EMail & "] " & _
		    " FROM Users u " & _
		    " INNER JOIN EmpleadosGlobal e ON u.IDEmpleado=e.IDEmpleado "
		    if SearchGroup<>"" then
			    SQL = SQL & " INNER JOIN UserGroup ug ON ug.IDEmpleado=u.IDEmpleado AND ug.IDGroup=" & SearchGroup
		    end if
		    if SearchAssignedTo<>"" then
			    SQL = SQL & " WHERE (e.NTUser LIKE '%" & replace(SearchAssignedTo,"'","''") & "%' OR e.ApellidosNombre LIKE '%" & replace(SearchAssignedTo,"'","''") & "%' )"
		    end if
		    SQL = SQL & strOrder
    		
		    'PrintSQL

		    vecWidth = Array("100", "250", "", "")

		    'on error resume next
		    rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
		    if rst.RecordCount=0 then
			    'Response.Write msgRstVacio
		    else
			    'Obtengo los datos con GetRows
			    aDatos = rst.GetRows

			    'Obtengo la página a mostrar de 
			    'la querystring
			    RegsPorPagina = Application("ListItems")
			    if RegsPorPagina="" then
				    RegsPorPagina = 8
			    end if
			    pag = CInt(Pagina)
			    if pag<0 then
				    rst.PageSize = RegsPorPagina
				    pag = rst.PageCount
			    end if
			    Pagina = pag
    			
			    iEstado = PaginarGR (RegsPorPagina, pag, aDatos, rst.Fields, "110", "User", vec, "100%", vecWidth)

		    end if
		    'Cierro y limpio objetos ya
		    rst.Close
		    on error goto 0
		    %>
    		
    		
		    <INPUT TYPE="HIDDEN" NAME="IDUser" ID="IDUser" VALUE="">



    <%elseif Pestana="2" then	'estamos en la pestaña de Grupos
    	
		    if Order<>"" then
			    strOrder = " Order By " & Order
		    else
			    strOrder = " Order By g.IDGroup"
		    end if
    		
		    SQL = "SELECT g.IDGroup AS [" & IDM_FLD_IDGroup & " ], g.Description AS [" & IDM_FLD_Description & "], " & _
		    " CONVERT(VARCHAR(200), g.Observations) AS [" & IDM_FLD_Observations & "] " 
		    SQL = SQL & " FROM Groups g "  
		    if SearchAssignedTo<>"" then
			    SQL = SQL & " INNER JOIN UserGroup ug ON ug.IDGroup=g.IDGroup "
			    SQL = SQL & " INNER JOIN EmpleadosGlobal e ON ug.IDEmpleado=e.IDEmpleado AND (e.NTUser LIKE '%" & replace(SearchAssignedTo,"'","''") & "%' OR e.ApellidosNombre LIKE '%" & replace(SearchAssignedTo,"'","''") & "%' )"
		    end if
		    if SearchGroup<>"" then
			    SQL = SQL & " WHERE g.IDGroup='" & SearchGroup & "' "
		    end if
		    SQL = SQL & strOrder


		    'PrintSQL
    		

		    vecWidth = Array("100", "250", "", "", "", "")

		    rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly

		    if rst.RecordCount=0 then
			    'Response.Write msgRstVacio
		    else
			    'Obtengo los datos con GetRows
			    aDatos = rst.GetRows

			    'Obtengo la página a mostrar de 
			    'la querystring
			    RegsPorPagina = Application("ListItems")
			    if RegsPorPagina="" then
				    RegsPorPagina = 8
			    end if
			    pag = CInt(Pagina)
			    if pag<0 then
				    rst.PageSize = RegsPorPagina
				    pag = rst.PageCount
			    end if
			    Pagina = pag
    			
    			
			    iEstado = PaginarGR (RegsPorPagina, pag, aDatos, rst.Fields, "110", "Group", vec, "100%", vecWidth)

		    end if
		    'Cierro y limpio objetos ya
		    rst.Close
		    %>

		    <INPUT TYPE="HIDDEN" NAME="IDGroup" ID="IDGroup" VALUE="">

    	
    <%end if%>


</TR></TD></TABLE>


<%
'-------------------------------------------------------------
'-------------------------------------------------------------
'Manté els valors en el formulari %>
<INPUT type="hidden" id=EstadoASP name=EstadoASP value="<%=EstadoASP%>">
<INPUT type="hidden" id=Pestana name=Pestana value="<%=Pestana%>">
<INPUT type="hidden" id=NewUser name=NewUser value="<%=NewUser%>">
<INPUT type="hidden" id=NewGroup name=NewGroup value="<%=NewGroup%>">



<INPUT type="hidden" id=Order name=Order value="<%=Order%>">
<INPUT type="hidden" id=Pagina name=Pagina value="<%=Pagina%>">
<INPUT type="hidden" id=SearchUser name=SearchUser value="<%=SearchUser%>">
<INPUT type="hidden" id=SearchAssignedTo name=SearchAssignedTo value="<%=SearchAssignedTo%>">
<INPUT type="hidden" id=SearchGroup name=SearchGroup value="<%=SearchGroup%>">


<!-- #include file = "include/EventFunctions2.asp" -->



</form>


</body>
</html>

<%Response.Flush%>
