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
dim rst, rst2, menuType, SQL
dim IDQuery

set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")

rst.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient

RecoverSQLConnection()

RecoverSession(true) 


%>
<!-- #include file = "include/Idioma.asp" -->

<%
if not isAdmin() then
	msgError "You are not allowed to view this information", true, true
end if


on error resume next
IDQuery = CInt(Request("Q"))
if Err <> 0 then
    Response.End
end if
on error goto 0
%>
 
<html>
<head>
    <title><%=IDM_Activity %></title>
    <link rel=StyleSheet href="include/style.css" type="text/css">

</head>
<body>
<form action="" method="post" name="thisForm">
    
    <!-- #include file = "ClassTopButtons.asp" -->
    <div id="TOPMARGIN" style="margin-top:100px;"></div>
    
        <%
        dim f
        
        SQL = "SELECT * FROM ReportQuery WHERE ID = " & IDQuery
        rst.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
        if not rst.EOF then
            %><p class="font20" style="text-align:center;"><%=rst("Nombre") %></p>
            <table align="center" cellpadding="3" cellspacing="0">
            <%
            SQL = rst("Query")
            rst2.Open SQL, ObjConnectionSQL, adOpenStatic, adLockReadOnly
            if not rst2.EOF then
                %><tr><%
                for each f in rst2.Fields
                    %><td class="tableHead" ><%=f.name %></td><%
                next
                %></tr><%
                while not rst2.EOF
                    %><tr><%
                    for each f in rst2.Fields
                        %><td class="tableRow"><%=rst2(f.name) %></td><%
                    next
                    %></tr><%
                    rst2.MoveNext
                wend
            end if
            rst2.Close
            %></table><%
        end if
        rst.Close

        %>
    
    
</form>
</body>
</html>