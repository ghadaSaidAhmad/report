<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">

<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/EventFunctions1.asp" -->

<!-- #include file = "ClassInclude.asp" -->

<%
dim rst, rst2, menuType
set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")

rst.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient

RecoverSQLConnection()

RecoverSession(true) 


%>
<!-- #include file = "include/Idioma.asp" -->


<html>
<head>
    <title><%=IDM_Activity %></title>
    <link rel=StyleSheet href="include/style.css" type="text/css">

</head>
<body>
<form action="" method="post" name="thisForm">
    
    <!-- #include file = "ClassTopButtons.asp" -->
    <div id="TOPMARGIN" style="margin-top:100px;"></div>
    
    <table align="center" cellpadding="3" cellspacing="0">
        <tr>
            <td class="tableHead" width="250px">Cliente</td>
            <td class="tableHead" width="150px">Última Actividad</td>
            <td class="tableHead" width="150px">Usuario</td>
            <td class="tableHead" style="text-align:right;" >Días transcurridos</td>
        </tr>
        <%
        dim arrCli, c
        dim idAct, act
        arrCli = getClients("NOMBRE")
        for each c in arrCli
            lastUpdateClient(c.IDClient)
            idAct = lastUpdateClient(c.IDClient)
            if idAct <> 0 then 
                set act = getActivity(idAct)
            end if
            %>
            <tr>
                <td class="tableRow"><%=c.Name %></td>
                <td class="tableRow">
                    <%if idAct <> 0 then 
                        Response.Write act.LastUpdatedDate
                    end if %>
                </td>
                <td class="tableRow"><%if idAct <> 0 then 
                        Response.Write act.LastUpdatedBy
                    end if %>
                </td>
                <td class="tableRow" style="text-align:right;"><%if idAct <> 0 then
                        if NOT isNull(act.LastUpdatedDate) then 
                            Response.Write DateDiff("d", act.LastUpdatedDateDate, Date)
                        end if
                    end if%>
                </td>
            </tr>
        <%next %>
    </table>
    
    
</form>
</body>
</html>