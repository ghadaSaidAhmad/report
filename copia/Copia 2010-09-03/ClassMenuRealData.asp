
<script runat=server language=vbscript>
</script>

<script language=javascript>
    
    function imprimir()
    {
        MAIN_MENU.style.display = 'none';
        MENU_REPORT_NAVIGATION.style.display = 'none';
        TBL_TOP.style.display = 'none';
        print();
        MAIN_MENU.style.display = '';
        MENU_REPORT_NAVIGATION.style.display = '';
        TBL_TOP.style.display = '';
    }
    
    function showMenu()
    {
        MAIN_MENU.style.display = '';
    }
    
</script>

<%
dim menuWidth
menuWidth = 130
if IsAdmin() then
    menuWidth = 190
end if
if IsAdmin() OR IsInputData() then
    menuWidth = menuWidth + 60
end if
%>

<div ID="MAIN_MENU" style="display:none;text-align:right;position:absolute;right:40px;top:0px;width:<%=menuWidth%>px;height:40px;background-color:White;border-top:3 solid Gray;border-left:3 solid Gray;border-right:3 solid Gray;border-bottom:3 solid Gray;">
    <%if IsAdmin() then %>
        <img onclick="location.href='AdminAppVariables.asp';return false;" alt="<%=IDM_MenuConfig %>" style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/config.png" />
        &nbsp;&nbsp;&nbsp;&nbsp;
    <%end if %>
    <%if IsAdmin() OR IsInputData() then %>
        <img onclick="location.href='RealData.asp';return false;" alt="<%=IDM_MenuInputData %>" style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" height=30 width=30 onmouseout="this.style.border = '2 solid white';" src="images/btnedit.png" />
    <%end if %>
    <img onclick="location.href='SOA.asp';return false" alt="<%=IDM_MenuFilterReport %>" style="width:30px;height:30px;cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/report.png" />
</div>


