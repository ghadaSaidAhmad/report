<%

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
            <img onclick="window.open('Docs/<%=Application("Admin_Help_Doc") %>', 'HELP', '');return false;" src="images/help3.png" style="width:35px;height:35px;" />
            Help
        </div>

        <div class="<%=TopButtonsStyle %>">
            <img onclick="location.href='SOA.asp';" src="images/form3.png" style="width:35px;height:35px;"/>
            Report
        </div>

        <%if IsAdmin() OR IsInputData() then %>
            <div class="<%=TopButtonsStyle %>">
                <img onclick="location.href='RealData.asp';return false;" src="images/realdata3.png" style="height:35px;width:35px" />
                GPV
            </div>
        <%end if %>

        <div class="<%=TopButtonsStyle %>">
            <img onclick="location.href='AdminFormList.asp';" src="images/formquality.png" style="width:35px;height:35px;"/>
            Forms
        </div>

        <div class="<%=TopButtonsStyle %>">
            <img onclick="location.href='ClientBrandList.asp';" src="images/clientbrand3.png" style="width:35px;height:35px;"/>
            Cli/Bra
        </div>

        <div class="<%=TopButtonsStyle %>">
            <img onclick="location.href='UserGroupList.asp';" src="images/users3.png" style="width:35px;height:35px;"/>
            Users
        </div>

        <div class="<%=TopButtonsStyle %>">
            <img onclick="location.href='AdminAppVariables.asp';" src="images/parameters3.png" style="width:35px;height:35px;"/>
            Params
        </div>


    </div>


    
</div>

<div style="margin-top:80px;"></div>
