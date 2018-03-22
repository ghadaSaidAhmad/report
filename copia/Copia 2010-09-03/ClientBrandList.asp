<%@language=VBScript%>
<%  Option Explicit
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
%>


<%

'-------------------------------------------------------------
'-------------------------------------------------------------
Sub DeleteClient_click(DelIDClient)
	
	on error resume next
	SQL = "UPDATE Client SET indBaja=1 WHERE IDClient=" & DelIDClient
	ObjConnectionSQL.Execute SQL
	if Err<>0 then
		MsgError "Error 10: " & IDM_ClientBrandListErr10, false, false
	end if
	
End Sub

Sub DeleteBrand_click(DelIDBrand)
	
	on error resume next
	SQL = "UPDATE Brand SET indBaja=1 WHERE IDBrand=" & DelIDBrand
	ObjConnectionSQL.Execute SQL
	if Err<>0 then
		MsgError "Error 20: " & IDM_ClientBrandListErr20, false, false
	end if
	
	
End Sub

Sub Search_click()
	Srch_TxtSearch = TxtSearch
	Pagina = "1"
End Sub



'-------------------------------------------------------------
'-------------------------------------------------------------
'Manté els valors en el formulari ----------------------------
dim IDClient: IDClient = request("IDClient")
dim IDBrand: IDBrand = request("IDBrand")

dim Pestana: Pestana = request("Pestana")
if Pestana="" then
	Pestana="1"
end if
dim Pagina: Pagina = request("Pagina")
dim Order: Order = request("Order")
dim verBorrados: verBorrados = request("verBorrados")




'-------------------------------------------------------------
'-------------------------------------------------------------
'Reconeix l'event --------------------------------------------
EventObject = request("EventObject")
EventParam1 = request("EventParam1")
EventParam2 = request("EventParam2")
select case EventObject

	case "Search" call Search_click()
	case "DeleteClient" call DeleteClient_click(EventParam1)
	case "DeleteBrand" call DeleteBrand_click(EventParam1)
	case "DeleteProduct" call DeleteProduct_click(EventParam1)
	case "DeleteFormat" call DeleteFormat_click(EventParam1)
	case "DeleteSize" call DeleteSize_click(EventParam1)
	case "Siguiente" call Siguiente_click()
	case "Anterior" call Anterior_click()
	case "Primero" call Primero_click()
	case "Ultimo" call Ultimo_click()

end select

%>


<%
if Pagina = "" then
	Pagina = "1"
end if

%>

<script language="JavaScript">
function editElement(val , tipo){
	if (tipo=="Client"){
		window.open('ClientNewEdit.asp?IDClient=' + val,'Client','top=200,left=300,width=350,height=200,scrollbars');
	}
	else if (tipo=="Brand"){
		window.open('BrandNewEdit.asp?IDBrand=' + val,'Brand','top=200,left=300,width=350,height=250,scrollbars');
	}
	else if (tipo=="Product"){
		window.open('ProductNewEdit.asp?IDProduct=' + val,'Product','top=200,left=300,width=350,height=230,scrollbars');
	}
	else if (tipo=="Format"){
		window.open('FormatNewEdit.asp?IDFormat=' + val,'Format','top=200,left=300,width=350,height=200,scrollbars');
	}
	else if (tipo=="Size"){
		window.open('SizeNewEdit.asp?IDSize=' + val,'Size','top=200,left=300,width=350,height=200,scrollbars');
	}
}

function DeleteElement(val, tipo){
	if (tipo=="Client"){
		if (confirm('Click OK to Continue. Click Cancel to Abort.')){
			_fireEvent('DeleteClient',val,'');}
	}else if (tipo=="Brand"){
		if (confirm('Click OK to Continue. Click Cancel to Abort.'))
			_fireEvent('DeleteBrand',val,'');
	}
	else if (tipo=="Product"){
		if (confirm('Click OK to Continue. Click Cancel to Abort.'))
			_fireEvent('DeleteProduct',val,'');
	}
	else if (tipo=="Format"){
		if (confirm('Click OK to Continue. Click Cancel to Abort.'))
			_fireEvent('DeleteFormat',val,'');
	}
	else if (tipo=="Size"){
		if (confirm('Click OK to Continue. Click Cancel to Abort.'))
			_fireEvent('DeleteSize',val,'');
	}
}

function ResetFields(){
	thisForm.TxtSearch.selectedIndex=0;
}
</script>


<html>
<head>
<title><%=IDM_ClientBrandListTitle%></title>

	<LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">

</head>
<BODY class=BODY_MAIN style="background-image:url('images/background.jpg');background-repeat:no-repeat;">


<form action="ClientBrandList.asp" method="post" name="thisForm">

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
    

	<table border=0 cellpadding=0 cellspacing=0 width="600px" style="border:1 solid gray;" align=center><tr height="400px"><td align="center" valign="top" style="border-right:2px solid black;border-bottom:2px solid black;">
	    
        <!-- #include file = "ClassMenuAdmin.asp" -->


	    <table border=0 cellpadding=0 cellspacing=0 width="100%">
	    <tr>
		    <td align="left" style="padding-left:5">
			    <img src="images/create.gif" style="width:30px;height:30px;"/><FONT class="font20">&nbsp;&nbsp;&nbsp;<STRONG><%=IDM_ClientBrandListTitle%></STRONG></FONT>
		    </td>
		    <td align=right>
		        <%'Botons de guardar, cancel·lar%>
		        <font face=Verdana size=1><%=IDM_Deleted %></font><input type="checkbox" style="border:0" onClick="thisForm.submit();" name="verBorrados" <%if verBorrados="on" then%> checked<%end if%>>
		        <%if Pestana="1" then	'estamos en la pestaña de Clientes%>
			        <a href="" title="<%=IDM_NewClient%>" onclick="window.open('ClientNewEdit.asp?NewClient=1','Client','top=200,left=300,width=350,height=200,scrollbars');return false;"><img src="images/add.png" border=0></a>
		        </td></tr>
		        <%elseif Pestana="2" then	'estamos en la pestaña de Grupos%>
			        <a href="" title="<%=IDM_NewBrand%>" onclick="window.open('BrandNewEdit.asp?NewBrand=1','Brand','top=200,left=300,width=350,height=250,scrollbars');return false;"><img src="images/add.png" border=0></a>
		        </td></tr>
		        <%end if%>
		        
		    </td>
		</tr>
		</table>
		
	    <hr style="height:3px;" color=black />

        <table width=100% cellspacing=0>
	        <tr>
		        <td width=20 CLASS="PEST_ESPACIO"><font class="fontNorm">&nbsp;</td>
		        <td onclick="location.href='ClientBrandList.asp?Pestana=1';" width=120 CLASS="<%if Pestana="1" then%>PEST_SELEC<%else%>PEST_NOSELEC<%end if%>"><font class="fontNorm"><b><%=IDM_Clients%> <label id="TotalUsers"></label></td>
		        <td width=5 CLASS="PEST_ESPACIO">&nbsp;</td>
		        <td onclick="location.href='ClientBrandList.asp?Pestana=2';" width=120 CLASS="<%if Pestana="2" then%>PEST_SELEC<%else%>PEST_NOSELEC<%end if%>"><font class="fontNorm"><b><%=IDM_Brands%> <label id="TotalGroups"></label></td>
		        <td CLASS="PEST_ESPACIO">&nbsp;</td>
	        </tr>
        </table>
        <br>


        <%
        dim rstCall: Set rstCall = Server.CreateObject("ADODB.Recordset")
        dim strBorrados, strOrder, vecWidth, aDatos, RegsPorPagina, pag, vec, strSql, iEstado
        if verBorrados="on" then
	        strBorrados = " where indBaja=1"
        else
	        strBorrados = " where indBaja=0"
        end if
        if Pestana="1" then	'estamos en la pestaña de Clientes

	        	if Order<>"" then
			        strOrder = " Order By " & Order
		        else
			        strOrder = " ORDER BY Orden"
		        end if
		        SQL = "SELECT IDClient, " & _
		        " Name AS " & IDM_Client & " , " & _
		        " SiebelCode AS [PlanTo], Orden AS [" & IDM_Orden & "], " & _
		        " CASE WHEN indBaja<>0 THEN '<font color=red><b>" & IDM_Deleted & "</font>' ELSE '' END AS [&nbsp;] FROM Client "
		        SQL = SQL & strBorrados & strOrder
        		
		        'PrintSQL

		        vecWidth = Array("80", "150", "150", "", "", "")

		        'on error resume next
		        rstCall.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
		        if rstCall.RecordCount=0 then
			        'Response.Write msgRstVacio
		        else
			        'Obtengo los datos con GetRows
			        aDatos = rstCall.GetRows

			        'Obtengo la página a mostrar de 
			        'la querystring
			        RegsPorPagina = Application("ListItems")
			        if RegsPorPagina="" then
				        RegsPorPagina = 8
			        end if
			        pag = CInt(Pagina)
			        if pag<0 then
				        rstCall.PageSize = RegsPorPagina
				        pag = rstCall.PageCount
			        end if
			        Pagina = pag
        			
			        iEstado = PaginarGR (RegsPorPagina, pag, aDatos, rstCall.Fields, "110", "Client", vec, "100%", vecWidth)

		        end if
		        'Cierro y limpio objetos ya
		        rstCall.Close
		        set rstCall = nothing
		        on error goto 0
		        
        		
        		

        elseif Pestana="2" then	'estamos en la pestaña de Brands
        	
		        if Order<>"" then
			        strOrder = " Order By [" & Order & "]"
		        else
			        strOrder = " Order By Orden"
		        end if
        		
		        SQL = "SELECT IDBrand, " & _
		        " Name AS [" & IDM_Brand & "] , " & _
		        " ShortName AS [" & IDM_ShortName & "], " & _
		        " SiebelCode AS [" & IDM_BrandCode & "], " & _
		        " Orden AS [" & IDM_Orden & "], " & _
		        " CASE WHEN indBaja<>0 THEN '<font color=red><b>" & IDM_Deleted & "</font>' ELSE '' END AS [&nbsp;] " & _
		        " FROM Brand "
		        SQL = SQL & strBorrados & strOrder

		        vecWidth = Array("70", "150", "100", "90", "","")

		        rstCall.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly

		        if rstCall.RecordCount=0 then
			        'Response.Write msgRstVacio
		        else
			        'Obtengo los datos con GetRows
			        aDatos = rstCall.GetRows

			        'Obtengo la página a mostrar de 
			        'la querystring
			        RegsPorPagina = Application("ListItems")
			        if RegsPorPagina="" then
				        RegsPorPagina = 8
			        end if
			        pag = CInt(Pagina)
			        if pag<0 then
				        rstCall.PageSize = RegsPorPagina
				        pag = rstCall.PageCount
			        end if
			        Pagina = pag
        			
			        iEstado = PaginarGR (RegsPorPagina, pag, aDatos, rstCall.Fields, "110", "Brand", vec, "100%", vecWidth)

		        end if
		        'Cierro y limpio objetos ya
		        rstCall.Close
		        set rstCall = nothing%>

		        <INPUT TYPE="HIDDEN" NAME="IDGroup" ID="IDGroup" VALUE="">


        <%end if%>


		</TABLE>

<br />
<br />
<br />
<br />



<%
'-------------------------------------------------------------
'-------------------------------------------------------------
'Manté els valors en el formulari %>
<INPUT type="hidden" id=Pestana name=Pestana value="<%=Pestana%>">


<INPUT type="hidden" id=Order name=Order value="<%=Order%>">
<INPUT type="hidden" id=Pagina name=Pagina value="<%=Pagina%>">


<!-- #include file = "include/EventFunctions2.asp" -->



</form>


</body>
</html>

<%Response.Flush%>
