<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
	Response.Buffer=true
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">

<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/EventFunctions1.asp" -->

<!-- #include file = "ClassInclude.asp" -->

<%
dim rst, SQL
Set rst = Server.CreateObject("ADODB.Recordset")%>


<%	
RecoverSQLConnection()
RecoverSession(true)
%>
<!-- #include file = "include/Idioma.asp" -->


<%
if not isAdmin() then
	msgError "You are not allowed to view this information", true, true
end if


Sub SaveParams_click()
    dim i, r, ID, Value, SQL
	
	for each r in Request.Form
		if Mid(r, 1, 4) = "TXT_" then
			ID = Mid(r, 5)
			Value = Request.Form(r)
			
			SQL = "UPDATE TableVarApplication SET VarValue=N'" & replace(Value,"'","''") & "' WHERE ID=" & ID
			ObjConnectionSQL.Execute(SQL)
			
		end if
	next
	
	'Load Application configuration
	Application("ApplicationCharged") = ""
	RecoverApplication()
	
End Sub



Function currSeleccionado(valor, arrayValores)
dim a, curr

	curr = FALSE
	for each a in arrayValores
		if UCASE(a)=UCASE(valor) then
			curr = TRUE
		end if
	next
	
	currSeleccionado = curr
	
End Function


EventObject = Request.Form("EventObject")
EventParam1 = Request.Form("EventParam1")
select case EventObject
	case "SaveParams" call SaveParams_click()
end select
%>

<head>
<title>Variables</title>

	<LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">

	<script language="JavaScript">
		function Save(){
			_fireEvent('SaveParams','','');
		}
	</script>

</head>


<BODY class=BODY_MAIN style="background-image:url('images/background.jpg');background-repeat:no-repeat;">

<FORM method=post name="thisForm">


    <!-- #include file = "ClassTopButtonsAdmin.asp" -->


	<table border=0 cellpadding=0 cellspacing=0 width="600px" style="border:1px solid gray;" align=center><tr height="400px"><td align="center" valign="top" style="border-right:2px solid black;border-bottom:2px solid black;">
	    

	    <table border=0 cellpadding=0 cellspacing=0 width="100%">
	    <tr>
		    <td align="left" style="padding-left:5">
			    <img src="images/parameters.png" /><FONT class="font20">&nbsp;&nbsp;&nbsp;<STRONG><%=IDM_ConfigParametersTitle%></STRONG></FONT>
		    </td>
		    <td align=right>
		        <img onclick="Save();" title="<%=IDM_MenuSaveParameters %>" style="cursor:pointer;border:1px solid white;" onmouseover="this.style.border = '1px solid gray';" onmouseout="this.style.border = '1px solid white';" src="images/save.png" />
		    </td>
		</tr>
		</table>
	    
		<TABLE style="text-align:left;">

		<%SQL = "SELECT * FROM TableVarApplication WHERE CanAdmin<>0 ORDER BY NGroup, Orden"
		rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
		
		dim CurrNGroup, strGroupSeparation
		
		CurrNGroup = 0
		while not rst.EOF
			strGroupSeparation = ""
			if CurrNGroup<>rst("NGroup") then
				strGroupSeparation = " <TR><TD colspan=10><HR color=black size=3></TD></TR> "
				CurrNGroup = rst("NGroup")
			end if
			%>
			<%=strGroupSeparation%>
			
			<%if rst("VarType")="GROUPTITLE" then%>
				<TR height=40>
					<TD colspan=10><b><font class=font15><%=rst("VarValue")%></font></b></TD>
				</TR>
			<%else%>
				<TR>
					<TD width=20></TD>
					<TD width=350 valign=top style="border-bottom:1px solid gray;"><font class=font12><%=rst("Description")%></TD>
					<TD valign=top>
						<%
						dim strInputWidth
						dim ArrValues, a
						
						strInputWidth = ""
						if not isnull(rst("InputWidth")) then
							strInputWidth = rst("InputWidth")
						end if
						%>
						<%if rst("VarType")="YESNO" then
							if strInputWidth="" then
								strInputWidth = "60"
							end if
							%>
							<SELECT NAME="TXT_<%=rst("ID")%>" style="width:<%=strInputWidth%>">
								<OPTION VALUE="YES" <%if rst("VarValue")="YES" then%> selected <%end if%>>YES</OPTION>
								<OPTION VALUE="NO" <%if rst("VarValue")="NO" then%> selected <%end if%>>NO</OPTION>
							</SELECT>
						<%elseif rst("VarType")="COMBO" then
							if strInputWidth="" then
								strInputWidth = "60"
							end if
							
							
							ArrValues = Split(rst("ComboValues"),";")
							%>
							<SELECT NAME="TXT_<%=rst("ID")%>" style="width:<%=strInputWidth%>">
								<%for each a in ArrValues%>
									<OPTION VALUE="<%=a%>" <%if UCASE(a)=UCASE(rst("VarValue")) then%> selected <%end if%>><%=a%></OPTION>
								<%next%>
							</SELECT>
						<%elseif rst("VarType")="COLOR" then
							if strInputWidth="" then
								strInputWidth = "120"
							end if
							%>
							<INPUT style="width:<%=strInputWidth%>;background-color:<%=rst("VarValue")%>" Name="TXT_<%=rst("ID")%>" TYPE="TEXT" VALUE="<%=rst("VarValue")%>">
							<input type=button value="Select" style="width:60" onclick="window.open('ColorPicker.asp?campo=TXT_<%=rst("ID")%>','COL','width=520,height=550');return false;">
						
						<%elseif rst("VarType")="MULTILIST" then
							if strInputWidth="" then
								strInputWidth = "120"
							end if
							ArrValues = Split(rst("ComboValues"),";")
							
							if not isNull(rst("VarValue")) then
								ArrCurrValues = Split(rst("VarValue"), ", ")
							else
								ArrCurrValues = Array()
							end if
							
							%>
							
							<SELECT size=8 NAME="TXT_<%=rst("ID")%>" style="height=60;width:<%=strInputWidth%>" MULTIPLE>
								<%for each a in ArrValues%>
									<OPTION VALUE="<%=a%>" <%if currSeleccionado(a, ArrCurrValues) then%> selected <%end if%>><%=a%></OPTION>
								<%next%>
							</SELECT>
							
						<%elseif rst("VarType")="DATE" then
							if strInputWidth="" then
								strInputWidth = "80"
							end if
							%>
							<INPUT style="width:<%=strInputWidth%>" Name="TXT_<%=rst("ID")%>" TYPE="TEXT" VALUE="<%=rst("VarValue")%>" readonly>
							<A HREF="" onclick="window.open('include/calendario.asp?campo=TXT_<%=rst("ID")%>','CAL','width=150,height=170');return false;"><img src="images/calendar.gif" border=0></A>

						<%elseif rst("VarType")="PASSWORD" then
							if strInputWidth="" then
								strInputWidth = "120"
							end if
							%>
							<INPUT style="width:<%=strInputWidth%>" Name="TXT_<%=rst("ID")%>" TYPE="PASSWORD" TITLE="<%=rst("VarValue")%>" VALUE="<%=rst("VarValue")%>" >


						<%else
							if strInputWidth="" then
								strInputWidth = "120"
							end if
							%>
							<INPUT style="width:<%=strInputWidth%>" Name="TXT_<%=rst("ID")%>" TYPE="TEXT" VALUE="<%=rst("VarValue")%>">
						<%end if%>
					</TD>
				</TR>
			
			<%end if%>
			
			
			<%
			rst.MoveNext
		wend
		rst.Close
		%>

		</TABLE>

<br />
<br />
<br />
<br />


<!-- #include file = "include/EventFunctions2.asp" -->


</FORM>