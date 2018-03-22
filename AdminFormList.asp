<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
	Response.Buffer=true
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">

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


dim Pestana: Pestana = request("Pestana")
if Pestana="" then
	Pestana="1"
end if
dim Pagina: Pagina = request("Pagina")
dim Order: Order = request("Order")
dim verBorrados: verBorrados = request("verBorrados")


Sub deleteForm_click(idForm)
    dim frm
    set frm = getForm(idForm)
    if frm.Enabled then
        disableForm(idForm)
    else
        deleteForm(idForm)
    end if
    
    
End Sub


EventObject = Request.Form("EventObject")
EventParam1 = Request.Form("EventParam1")
select case EventObject
	case "Search" call Search_click()
	case "deleteForm" call deleteForm_click(EventParam1)
	
	case "Siguiente" call Siguiente_click()
	case "Anterior" call Anterior_click()
	case "Primero" call Primero_click()
	case "Ultimo" call Ultimo_click()
end select

if Pagina = "" then
	Pagina = "1"
end if
%>

<head>
<title><%=IDM_AdminFormsTitle %></title>

	<LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">

	<script language="JavaScript">
	    function editElement(val , tipo){
	        location.href = 'AdminFormNewEdit.asp?ID=' + val;
	    }
	    function DeleteElement(val, tipo){
	        var sMsg;
	        if (tipo == 'on'){
	            sMsg = "<%=IDM_JS_DeleteFormMessage %>";
	        }else{
	            sMsg = "";
	        }
            _fireConfirm('deleteForm', val, '', sMsg);
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
			    <img src="images/form.png" style="width:30px;height:30px;"/><FONT class="font20">&nbsp;&nbsp;&nbsp;<STRONG><%=IDM_AdminFormsTitle%></STRONG></FONT>
		    </td>
		    <td align=right>
		        <%'Botons de guardar, cancel·lar%>
		        <font face=Verdana size=1><%=IDM_Deleted %></font><input type="checkbox" style="border:0" onClick="thisForm.submit();" name="verBorrados" <%if verBorrados="on" then%> checked<%end if%>>
			        <a href="" title="<%=IDM_NewForm%>" onclick="location.href='AdminFormNewEdit.asp';return false;"><img src="images/add.png" border=0></a>
		        </td></tr>
		        
		    </td>
		</tr>
		</table>
		
	    <hr style="height:3px;" color="black" />
        
        
        <%if FALSE then %>
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
        <%end if %>

        <table width=100% cellspacing=0>
	        <tr>
	            <td width=50><font class="font12"><%=IDM_Brand %></font></td>
	            <td>
                    <select style="width:200px;" name="idBrand" onchange="_fireEvent('', '', '');" >
                        <option value=""><%=IDM_SelectBrand %></option>
                        <%
                        dim lstbra, sSelected
                        dim arrBrands
                        arrBrands = getBrands("ORDEN")
                        for each lstbra in arrBrands
                            sSelected = ""
                            if Request.Form("idBrand")<>"" then
                                if CInt(Request.Form("idBrand")) = lstbra.IDBrand then
                                    sSelected = "selected"
                                end if
                            end if
                            %><option value="<%=lstbra.IDBrand %>" <%=sSelected %>><%=lstbra.Name %></option><%
                        next
                        %>
                    </select>
	            </td>
            </tr>
        </table>

        <%
        dim rstCall: Set rstCall = Server.CreateObject("ADODB.Recordset")
        dim strBorrados, strOrder, vecWidth, aDatos, RegsPorPagina, pag, vec, strSql, iEstado
        dim idBrand, strWhere
        
        strWhere = " WHERE 1=1 "
        idBrand = Request.Form("idBrand")
        
        
        if verBorrados="on" then
	        strWhere = strWhere & " AND f.indBaja=1"
        else
	        strWhere = strWhere & " AND f.indBaja=0"
        end if

    	if Order<>"" then
	        strOrder = " Order By [" & Order & "]"
        else
	        strOrder = " ORDER BY f.IDForm "
        end if
        SQL = "SELECT f.IDForm, " & _
        " convert(varchar,f.Name) AS [" & IDM_FormName & "] , " & _
        " convert(varchar,f.DateFrom,103) AS [" & IDM_FromDate & "] , " & _
        " CASE WHEN f.indBaja<>0 THEN '<font color=red><b>" & IDM_Deleted & "</font>' ELSE '' END AS [&nbsp;] " & _
        " FROM Form f "

        if idBrand<>"" then
            SQL = SQL & " INNER JOIN Brand b ON f.idForm = b.idForm AND b.idBrand = " & Replace(idBrand, "''", "")
        end if
        
        SQL = SQL & strWhere & strOrder
		
        'PrintSQL

        vecWidth = Array("80", "", "90", "", "", "")

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
			
	        iEstado = PaginarGR (RegsPorPagina, pag, aDatos, rstCall.Fields, "110", verBorrados, vec, "100%", vecWidth)

        end if
        'Cierro y limpio objetos ya
        rstCall.Close
        set rstCall = nothing
        on error goto 0
        
        	
        %>

		</TABLE>

<br />
<br />
<br />
<br />

<INPUT type="hidden" id=Order name=Order value="<%=Order%>">
<INPUT type="hidden" id=Pagina name=Pagina value="<%=Pagina%>">

<!-- #include file = "include/EventFunctions2.asp" -->


</FORM>