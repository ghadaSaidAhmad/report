<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
%>

<!-- #include file = "include/adovbs.asp" -->
<!-- #include file = "include/startup.asp" -->
<!-- #include file = "include/SrvrFunctions.asp" -->
<!-- #include file = "include/EventFunctions1.asp" -->

<!-- #include file = "RenderFunctions.asp" -->
<!-- #include file = "ClassInclude.asp" -->


<%
dim rst, rst2
set rst = CreateObject("ADODB.RecordSet")
set rst2 = CreateObject("ADODB.RecordSet")

rst.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient

RecoverSQLConnection()

RecoverSession(true) 


%>
<!-- #include file = "include/Idioma.asp" -->

<%
dim ID: ID = Request("ID")

dim IDClient: IDClient = Request("IDClient")
dim IDBrand: IDBrand = Request("IDBrand")
dim WYear: WYear = Request("WYear")
dim WMonth: WMonth = Request("WMonth")
dim WHalf: WHalf = Request("WHalf")
dim IDType: IDType = Request("IDType")

Sub Save_click()
    
    dim act
    set act = new Activity02
    
    act.ID = CInt(Request("ID"))
    act.IDClient = IDClient
    act.IDBrand = IDBrand
    act.WYear = WYear
    act.WMonth = WMonth
    act.WHalf = WHalf
    act.IDType = IDType
    act.Name = Request.Form("Name")

    act.NShops = Request.Form("NShops")
    act.PercentComplaint = Request.Form("PercentComplaint")
    act.IDStatus = Request.Form("IDStatus")
    
    on error resume next
    saveActivity02(act)
    if Err<>0 then
        bottomMessage = Err.Description
    else
        'Si era nuevo, ahora tiene un ID

        ID = act.ID

    end if
    on error goto 0
    
    if Request.Form("CloseWindow") <> "" then
        %><script language="JavaScript">try{window.close();}catch(e){} try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} try{window.opener.applyFilter(false);}catch(e){}</script><%
    end if

End Sub

Sub Delete_click(delID)
    
    on error resume next
    deleteActivity02(delID)

    if Err<>0 then
        bottomMessage = "Error deleting activity"
    else
        ' Cierra la ventana
        
        %><script language="JavaScript">try{window.close();}catch(e){} try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} try{window.opener.applyFilter(false);}catch(e){}</script><%
    end if
    
End Sub


Select Case EventObject
	case "Save" Save_click()
	case "Delete" Delete_click(EventParam1)
End Select


dim act

if CInt(ID) > -1 then
    set act = getActivity02(ID)
else
    ' Es un elemento nuevo
    set act = new Activity02
end if

dim aType, aCli, aBra
set aType = getActivityType(IDType, Idioma)
set aCli = getClient(IDClient)
set aBra = getBrand(IDBrand)

%>
<HTML>
<HEAD>
    <TITLE>Activity</TITLE>
    <LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">
    <script>
        var dataModified = false;

        function closeWindow()
        {
            if (dataModified)
            {
                if (confirm('<%=IDM_JS_DatosModificadosGuardar %>')){
                    thisForm.CloseWindow.value = '1';
                    Save();
                    return false;
                }
            }
            
          <%if Request("PageReloaded")<>"" then %> 
            try{window.opener.thisForm.action='#<%=IDClient & "_" & IDBrand %>';}catch(e){} 
            try{window.opener.applyFilter(false);}catch(e){} 
          <%end if %> 
          
          try{window.close();}catch(e){}
        }
        function Save()
        {
            _fireEvent('Save', '', '');
        }
        function Delete(id)
        {
            _fireConfirm('Delete', id, '', '');
        }
    </script>
</HEAD>

<BODY leftmargin=0 topmargin=0 >

<FORM action="SOAActivity02.asp?ID=<%=ID %>" method="post" name="thisForm">
    
    
    <table style="width:100%;height:40px;background-image:url('images/Grad5.gif'); ">
        <tr>
            <td valign="middle" style="padding-left:10px;">
                <font class="wopenTitle">
                    <%dim nTitleChars: nTitleChars = 25 %>
                    <%=Server.HTMLEncode(left(act.Name, nTitleChars)) %><%if len(act.Name)>nTitleChars then %>...<%end if %>
                </font>
            </td>
            <td align="right" width="180px;">
                <input type=button class="button" value="<%=IDM_Save %>" style="width:55px;" onclick="Save();" />
                <%if ID <> "-1" then %>
                    <input type=button class="button" value="<%=IDM_Delete %>" style="width:55px;" onclick="Delete(<%=ID %>);" />
                <%end if %>
                <input type=button class="button" value="<%=IDM_Close %>" style="width:55px;" onClick="closeWindow();" />
            </td>
        </tr>
    </table>
    
    <table style="width:100%;height:30px;border:2 solid gray;">
        <tr>
            <td valign=top width=100 class="fieldheader" ><%=IDM_ActivityType %></td>
            <td><font class="font12"><%=aType.Name %></font></td>
            <%
            dim sTiming
            if WHalf = 1 then
                sTiming = IDM_1aQuincena
            else
                sTiming = IDM_2aQuincena
            end if
            sTiming = sTiming & " " & locMonthName(WMonth, Idioma)
            sTiming = sTiming & " " & WYear
            %>
            <td align=right><font class="font15"><b><%=sTiming %></b></font></td>
        </tr>
        <tr>
            <td valign=top width=100 class="fieldheader" ><%=IDM_Client %></td>
            <td><font class="font12"><%=aCli.Name %></font></td>
        </tr>
        <tr>
            <td valign=top width=100 class="fieldheader" ><%=IDM_Brand %></td>
            <td><font class="font12"><%=aBra.Name %></font></td>
        </tr>

    </table>
    
    <table style="width:100%;height:30px;">
        <tr>
            <td valign=top width=100 class="fieldheader" ><%=IDM_Nombre %></td>
            <td>
                <textarea name="Name" class="textfield" onchange="dataModified=true;" style="width:100%;height:60px;"><%=act.Name%></textarea>
            </td>
        </tr>

        <tr>
            <td width=100 valign=top class="fieldheader"><%=IDM_NTiendas %></td>
            <td>
                <input name="NShops" value="<%=act.NShops %>" class="textfield" type="text" onchange="dataModified=true;" style="width:100%;" />
            </td>
        </tr>

        <tr>
            <td width=100 valign=top class="fieldheader"><%=IDM_PercentComplaint %></td>
            <td>
                <input name="PercentComplaint" value="<%=act.PercentComplaint %>" class="textfield" type="text" onchange="dataModified=true;" style="width:100%;" />
            </td>
        </tr>

        <tr>
            <td width=100 valign=top class="fieldheader"><%=IDM_Status %></td>
            <td>
                <%
                dim arrStatus
                dim sSelected
                dim s, iSt
                arrStatus = getActivityStatuses()
                iSt = 1
                for each s in arrStatus
                    sSelected = ""
                    if act.IDStatus<>-1 then
                        if s.ID = act.IDStatus then
                            sSelected = "checked"
                        end if
                    end if
                    %><input style="width:25px;height:25px;" type="radio" name="IDStatus" id="ST_<%=s.ID %>" value="<%=s.ID %>" <%=sSelected %> onchange="dataModified=true;" /><font class="font15"><%=s.Name %>&nbsp;&nbsp;&nbsp;</font><%
                    
                    if iSt mod 4 = 0 then
                        %><br /><%
                    end if
                    
                    iSt = iSt + 1
                next
                %>
            </td>
        </tr>

        <%if ID <> "-1" then %>
            <tr height=20><td></td></tr>
            <tr>
                <td valign=top width=100 class="fieldheader" style="border-top:1 solid silver;"><%=IDM_LastUpdatedBy %></td>
                <td style="border-top:1 solid silver;"><font class=font12>
                    <%=act.LastUpdatedBy %>
                    &nbsp;-&nbsp;
                    <%=act.LastUpdatedDate %>
                    </font>
                </td>
            </tr>
        <%end if %>
        
    </table>
    
    
    
    <input type=hidden name="ID" value="<%=act.ID %>" />
    <input type=hidden name="IDClient" value="<%=IDClient %>" />
    <input type=hidden name="IDBrand" value="<%=IDBrand %>" />
    <input type=hidden name="WYear" value="<%=WYear %>" />
    <input type=hidden name="WMonth" value="<%=WMonth %>" />
    <input type=hidden name="WHalf" value="<%=WHalf %>" />
    <input type=hidden name="IDType" value="<%=IDType %>" />
    
    <input type=hidden name="CloseWindow" value="" />

    <!-- #include file = "include/EventFunctions2.asp" -->

</FORM>


<!-- #include file = "include/pageBottom.asp" -->

<script language=javascript>
    thisForm.Name.focus();
</script>

</BODY>

</HTML>