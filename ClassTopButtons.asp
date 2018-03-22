<script language="javascript" type="text/javascript">
    function showMenu()
    {
        MAIN_MENU.style.display = '';
    }
</script>
<%
dim menuWidth
menuWidth = 80
if IsAdmin() then
    menuWidth = menuWidth + 40
end if
if IsAdmin() OR IsInputData() then
    menuWidth = menuWidth + 40
end if
if Request.Form("FILTER_REPORTTYPE") <> "" then
    menuWidth = menuWidth + 80
end if

dim TopButtonsStyle
if Application("TopButtonsStyle") = "All visible" then
    TopButtonsStyle = "topmenuicon_showall"
else
    TopButtonsStyle = "topmenuicon_hidetext"
end if
%>

<div id="DIV_TOP_BAR" style="width:100%;height:60px;background-image:url('images/a3.gif');z-index:999;position:fixed;top:0;left:0;">
    
    <div style="float:left;height:55px;position:relative;">
        <div class="TitleTop"><%=IDM_MAINTITLE1 %></div>
        <div class="SubTitleTop"><%=IDM_MAINTITLE2 %></div>
    </div>

    <div id="MAIN_MENU" style="float:right;padding-right:10px;height:55px;">

        <div class="<%=TopButtonsStyle %>">
            <img onclick="window.open('Docs/<%=Application("Users_Help_Doc") %>', 'HELP', '');return false;" src="images/help3.png" style="width:35px;height:35px;" />
            Help
        </div>

        <%if Request.Form("FILTER_REPORTTYPE") <> "" then %>
            <div class="<%=TopButtonsStyle %>" id="BTN_Export" >
                <img onclick="exportExcel();" src="images/excel3.png" style="width:35px;height:35px;" />
                Excel
            </div>
            <div class="<%=TopButtonsStyle %>" id="BTN_Print" >
                <img onclick="imprimir();" src="images/print3.png" style="width:35px;height:35px;" />
                Print
            </div>
        <%end if %>

        <div class="<%=TopButtonsStyle %>">
            <%if menuType = "SOA" then%>
                <img onclick="showFilter();" src="images/form3.png" style="width:35px;height:35px;" />
                Report
            <%else %>
                <img onclick="location.href='SOA.asp';return false"  src="images/form3.png" style="width:35px;height:35px;" />
                Report
            <%end if %>
        </div>

        <%if IsAdmin() OR IsInputData() then %>
            <div class="<%=TopButtonsStyle %>">
                <img onclick="location.href='RealData.asp';return false;" src="images/realdata3.png" style="height:35px;width:35px" />
                GPV
            </div>
        <%end if %>


        <%if IsAdmin() then %>
            <div class="<%=TopButtonsStyle %>">
                <img onclick="location.href='AdminAppVariables.asp';return false;" src="images/options3.png" style="width:35px;height:35px;" />
                Config
            </div>
        <%end if %>

            
        <%if puedeSuplantar OR session("PuedeSuplantar")<>"" then%>
            <div class="<%=TopButtonsStyle %>">
                <img onclick="document.getElementById('DIV_Suplantar').style.display='';form1.Suplantar.focus();return false;" src="images/suplantar3.png" style="width:35px;height:35px;" />
                Suplanta
            </div>
        <%end if%>

    </div>

</div>
