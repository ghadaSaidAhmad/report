<%@language=VBScript%>
<%  Option Explicit
    Response.Expires=0
	Response.Buffer=true
%>

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


dim idForm: idForm = Request.QueryString("ID")
if idForm="" then
    idForm = Request.Form("idForm")
    if idForm="" then
        idForm = -1
    end if
end if
idForm = CLng(idForm)

dim Pestana: Pestana = request("Pestana")
if Pestana="" then
	Pestana="1"
end if


Sub SaveForm_click()
    dim frm
    dim sIndBaja
    dim sDateFrom, dDateFrom
    
    if Request.Form("indBaja") = "" then
        sIndBaja = 0
    else
        sIndBaja = 1
    end if
    
    if Request.Form("DateFrom") = "" then
        dDateFrom = Date()
    else
        sDateFrom = Request.Form("DateFrom")
        dDateFrom = DateSerial(mid(sDateFrom, 7), mid(sDateFrom,4,2), mid(sDateFrom,1,2))
    end if
    
    if idForm > -1 then
        set frm = getForm(idForm)
    else
        set frm = new Form
    end if
    frm.Name = Request.Form("FormName")
    frm.indBaja = sIndBaja
    frm.DateFrom = dDateFrom
    
    saveForm(frm)
    
    
    idForm = CLng(frm.ID)
    
End Sub

Sub deleteQuestion_click(idQuest)
    
	if idQuest<>"" then
	    idQuest = cint(idQuest)
    	deleteQuestion(idQuest)
	    
	end if
    
End Sub

Sub deleteResponse_click(idResp)
	
	if idResp<>"" then
	    idResp = cint(idResp)
    	deleteResponse(idResp)
	    
	end if
	
End Sub


Sub saveNewQuestion_click()
    dim editIdQuest, dWeight, iIDRespType
    
    if Request.Form("newQuestText")<>"" AND Request.Form("newQuestWeight")<>"" AND isNumeric(Request.Form("newQuestWeight")) then
        
            dWeight = CDbl(Request.Form("newQuestWeight"))
            iIDRespType = CInt(Request.Form("newQuestRespType"))
            editIdQuest = Request.Form("editIdQuest")
            
            dim qst
            set qst = new FormQuestion
            qst.IDForm = idForm
            qst.Text = Request.Form("newQuestText")
            qst.Weight = dWeight
            qst.IDRespType = iIDRespType
            if editIdQuest<>"" then
                qst.ID = CLng(editIdQuest)
            end if
            
            saveQuestion(qst)
            
    else
        msgError IDM_JS_QuestTextWeight, false, false
    end if
    
End Sub


Sub saveNewResponse_click(idQuest)
    dim sResponseText, sResponseValue
    
    if idQuest<>"" then
        idQuest = CLng(idQuest)
        
        sResponseText = Request.Form("newRespText_" & idQuest)
        sResponseValue = Request.Form("newRespValue_" & idQuest)
        
        if sResponseText<>"" AND sResponseValue<>"" AND isNumeric(sResponseValue) then
            
            on error resume next
            sResponseValue = CLng(sResponseValue)
            if Err<>0 then
                msgError "Value not valid", false, false
                exit sub
            end if
            on error goto 0
            
            dim rsp
            set rsp = new FormResponse
            rsp.IDQuest = idQuest
            rsp.IDForm = idForm
            rsp.Text = sResponseText
            rsp.RespValue = sResponseValue
            
            saveResponse(rsp)
            
        else
            msgError IDM_JS_RespTextValue, false, false
        end if
    end if
End Sub


Sub removeFromBrand_click(idBrand, deleteHistory)
    dim bDeleteHistory
    
    if deleteHistory = "0" then
        bDeleteHistory = false
    else
        bDeleteHistory = true
    end if
    
    removeFormFromBrand idBrand, idForm, deleteHistory
    
    
End Sub


Sub removeFromPromotions_click(idBrand)
    
    removeFromPromotions idBrand, idForm
    
End Sub


Sub assignBrand_click(idBrand)
    
    assignBrand idBrand, idForm
    
End Sub


Sub reassignBrand_click(idBrand, deleteHistory)
    dim bDeleteHistory
    
    if deleteHistory = "0" then
        bDeleteHistory = false
    else
        bDeleteHistory = true
    end if
    
    reassignBrand idBrand, idForm, bDeleteHistory
    
End Sub


EventObject = Request.Form("EventObject")
EventParam1 = Request.Form("EventParam1")
EventParam2 = Request.Form("EventParam2")
select case EventObject
    case "SaveForm" SaveForm_click()
    case "deleteQuestion" deleteQuestion_click(EventParam1)
	case "deleteResponse" deleteResponse_click(EventParam1)
	case "saveNewQuestion" saveNewQuestion_click()
	case "saveNewResponse" saveNewResponse_click(EventParam1)
	case "removeFromBrand" removeFromBrand_click EventParam1, EventParam2
	case "removeFromPromotions" removeFromPromotions_click(EventParam1)
	case "assignBrand" assignBrand_click(EventParam1)
	case "reassignBrand" reassignBrand_click Eventparam1, EventParam2
end select
%>

<head>
<title><%=IDM_AdminFormTitle %></title>

	<LINK REL=StyleSheet HREF="include/style.css" TYPE="text/css">

	<script language="JavaScript">
		function Save(){
			_fireEvent('SaveForm','','');
		}
		function addQuestion(){
		    DIV_NewQuestion.style.display = '';
		}
		function cancelAddQuestion(){
		    thisForm.editIdQuest.value = '';
		    thisForm.newQuestText.value = '';
		    thisForm.newQuestWeight.value = '';
		    thisForm.newQuestRespType.selectedIndex = 0;
		    DIV_NewQuestion.style.display = 'none';
		}
		function saveNewQuestion(){
		    if (thisForm.newQuestText.value == '' || thisForm.newQuestWeight.value == '') { alert('<%=IDM_JS_QuestTextWeight %>'); return false;}
		    if (isNaN(thisForm.newQuestWeight.value)) { alert('<%=IDM_JS_QuestWeight %>'); return false;}
		    _fireEvent('saveNewQuestion', '', '');
		}
		function editQuestion(id, text, weight, resptype){
		    thisForm.editIdQuest.value = id;
		    thisForm.newQuestText.value = text;
		    thisForm.newQuestWeight.value = weight;

            var obj;
            obj = document.getElementById('newQuestRespType');
            for(var i=0; i<obj.options.length;i++)
                if(obj.options[i].value == resptype)
                    obj.selectedIndex = i;

		    DIV_NewQuestion.style.display = '';
		}
		function addResponse(tr, oValue, oText){
		    tr.style.display = '';
		    oValue.focus();
		}
		function cancelAddResponse(tr, oValue, oText){
		    tr.style.display = 'none';
		    oValue.value = '';
		    oText.value = '';
		}
		function saveNewResponse(idQuest, oValue, oText){
		    if (oValue.value == '' || oText.value == '') {alert('<%=IDM_JS_RespTextValue %>'); return false;}
		    
		    _fireEvent('saveNewResponse', idQuest, '');
		}
		
		function removeFromBrand(idBrand){
		    var sDeleteHistory
		    sDeleteHistory = 0
		    if (confirm('<%=IDM_JS_RemoveFormFromBrand %>')){
		        if (confirm('<%=IDM_JS_DeleteHistory %>')){
		            sDeleteHistory = "1";
		        }else{
		            sDeleteHistory = "0";
		        }
		        _fireEvent('removeFromBrand', idBrand, sDeleteHistory);
		    }
		}
		
		function removeFromPromotions(idBrand){
		    _fireConfirm('removeFromPromotions', idBrand, '', '<%=IDM_JS_DeleteHistoryBrand %>');
		}
		
		function assignBrand(idBrand){
		    _fireEvent('assignBrand', idBrand, '');
		}
		
		function reassignBrand(idBrand){
		    if (confirm('<%=IDM_JS_ReassignForm %>')){
		        if (confirm('<%=IDM_JS_DeleteHistoryBrand %>')){
		            sDeleteHistory = "1";
		        }else{
		            sDeleteHistory = "0";
		        }
		        _fireEvent('reassignBrand', idBrand, sDeleteHistory);
		    }
		}
	</script>
    
    <style type="text/css">
        .form
        {
            text-align:left;
            padding-left:10px;
            margin-bottom:15px;
        }
        .inputName
        {
            width:300px;
            border:1px solid black;
        }
    </style>
    
</head>


<BODY class=BODY_MAIN style="background-image:url('images/background.jpg');background-repeat:no-repeat;">

<FORM method=post name="thisForm">


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
    

	<table border=0 cellpadding=0 cellspacing=0 width="600px" style="border:1 solid gray;" align=center><tr height="400px"><td valign="top" style="border-right:2px solid black;border-bottom:2px solid black;">
	    
        <!-- #include file = "ClassMenuAdmin.asp" -->

	    <table border=0 cellpadding=0 cellspacing=0 width="100%">
	    <tr>
		    <td align="left" style="padding-left:5">
			    <img src="images/form.png" /><FONT class="font20">&nbsp;&nbsp;&nbsp;<STRONG><%=IDM_AdminFormTitle%></STRONG></FONT>
		    </td>
		    <td align=right>
		        <img onclick="Save();" title="<%=IDM_MenuSaveForm %>" style="cursor:pointer;border:2 solid white;" onmouseover="this.style.border = '2 solid gray';" onmouseout="this.style.border = '2 solid white';" src="images/save.png" />
		    </td>
		</tr>
		</table>
	    
	    <hr style="height:3px;" color=black />
	    
	    
	    <%
	    dim frm, canModify
	    if idForm > -1 then 
	        set frm = getForm(idForm)
	        canModify = frm.canModify
	    else
	        set frm = new Form
	        canModify = true
	    end if
	    
	    %>
    	<div class="form">
    	    <font class="fieldheader"><%=IDM_FormName %></font>&nbsp;&nbsp;&nbsp;<input type="text" name="FormName" value="<%=Server.HTMLEncode(frm.Name) %>" class="textfield" style="width:300px;" />
    	    <font class="font10">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
    	    <input type="checkbox" name="indBaja" <%if NOT frm.Enabled then %>checked <%end if %>/><font class="font12"><%=IDM_Deleted %></font>
    	    <br />
    	    <font class="fieldheader"><%=IDM_FromDate %></font>&nbsp;&nbsp;&nbsp;<input type="text" name="DateFrom" value="<%=right("0" & day(frm.DateFrom), 2) & "/" & right("0" & month(frm.DateFrom), 2) & "/" & year(frm.DateFrom) %>" class="textfieldreadonly" readonly style="font-size:11px;width:80px;" />
    	    <a href="" onclick="window.open('include/calendario.asp?campo=DateFrom','cal','width=150,height=200');return false;"><font class="font10"><img src="images/calendar.gif" style="border:0px;" border=0 alt="Select Date" align=middle></font></a>
    	</div>
    	
    	<%if idForm > -1 then %>
            
            
            <table width=100% cellspacing=0>
	            <tr>
		            <td width=20 CLASS="PEST_ESPACIO"><font class="fontNorm">&nbsp;</td>
		            <td onclick="thisForm.Pestana.value='1';thisForm.submit();" width=120 CLASS="<%if Pestana="1" then%>PEST_SELEC<%else%>PEST_NOSELEC<%end if%>"><font class="fontNorm"><b><%=IDM_PestQuestions %></td>
		            <td width=5 CLASS="PEST_ESPACIO">&nbsp;</td>
		            <td onclick="thisForm.Pestana.value='2';thisForm.submit();" width=120 CLASS="<%if Pestana="2" then%>PEST_SELEC<%else%>PEST_NOSELEC<%end if%>"><font class="fontNorm"><b><%=IDM_PestAssignBrands %></td>
		            <td CLASS="PEST_ESPACIO">&nbsp;</td>
	            </tr>
            </table>
            <br>
            
            
            
            <%if Pestana = "1" then %>
                <div id="DIV_NewQuestion" style="display:none;position:absolute;left:300px;top:150px;border:1px solid gray;width:400px;height:100px;background-color:White;">
                    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:Gray;color:White;">
                        <tr>
                            <td><font class="font10">&nbsp;<%=IDM_AddQuestion %>
                            <%if frm.weightTotal = 0 then
                            elseif frm.weightTotal < 100 then
                                %>&nbsp;&nbsp;&nbsp;<%=IDM_CurrentTotalW %>: <%=frm.weightTotal%><%
                            else
                                %>&nbsp;&nbsp;&nbsp;<font style="color:White;background-color:Red;padding-left:10px;padding-right:10px;"><%=IDM_CurrentTotalW %>: <%=frm.weightTotal%>%</font><%
                            end if %>
                            </font></td>
                            <td align="right" width="100" ><input type="button" onclick="cancelAddQuestion();return false;" value="X" /></td>
                        </tr>
                    </table>
                    
                    <input type="hidden" name="editIdQuest" value="" />
                    <table width="100%" cellpadding="3" cellspacing="0">
                        <tr>
                            <td><font class="font10"><%=IDM_Question %></font></td>
                            <td><font class="font10"><%=IDM_Weight %>%</font></td>
                            <td><font class="font10"><%=IDM_RespType %></font></td>
                        </tr>
                        <tr>
                            <td><input type="text" name="newQuestText" value="" style="width:280px;" <%if NOT canModify then %>style="background-color:gainsboro;" readonly<%end if %> /> </td>
                            <td><input type="text" name="newQuestWeight" value="" maxlength="2" style="width:30px;" /></td>
                            <td><select name="newQuestRespType">
                                <option value="0">List</option>
                                <option value="1">Radio</option>
                            </select></td>
                        </tr>
                    </table>
                    
                    <input type="button" class="button" value="<%=IDM_Save %>" onclick="saveNewQuestion();" />
                    <input type="button" class="button" value="<%=IDM_Cancel %>" onclick="cancelAddQuestion();" />
                </div>
                    
                
                
                
    	        <table width="100%" cellpadding="0" cellspacing="0">
    	            <tr>
    	                <td width="45"></td>
    	                <td>
    	                    <%if canModify then %>
        	                    <input type="button" class="button" value="<%=IDM_AddQuestion %>" onclick="addQuestion();return false;" />
        	                <%else %>
        	                    <font class="font11"><%=IDM_FormReadOnly %></font><br />
        	                <%end if %>
    	                    <%if frm.weightTotal <> 100 then 
    	                        if frm.numQuestions>0 then%>
        	                        <font class="font12">&nbsp;&nbsp;&nbsp;<font style="color:White;background-color:Red;padding-left:10px;padding-right:10px;"><b><%=UCASE(IDM_TotalWeight) %>: <%=frm.weightTotal %>%</b></font></font>
    	                        <%end if
    	                    else%>
    	                        <font class="font12">&nbsp;&nbsp;&nbsp;<font style="color:white;background-color:Green;padding-left:10px;padding-right:10px;"><b><%=UCASE(IDM_TotalWeight) %>: <%=frm.weightTotal %>%</b></font></font>
    	                    <%end if %>
    	                </td>
    	                <%if frm.numQuestions>0 then %><td width="70" align="center"><font class="font12"><%=IDM_Weight %>%</font></td><%end if %>
    	                <%if frm.numQuestions>0 then %><td width="70" align="center"><font class="font12"><%=IDM_RespType %></font></td><%end if %>
    	            </tr>
    	            <%
                    dim qst, rsp
                    for each qst in frm.Questions %>
                        <tr>
                            <td valign="top">
                                <%if canModify then %>
                                    <a title="<%=IDM_DeleteQuestion %>" href="" onclick="_fireConfirm('deleteQuestion', '<%=qst.ID %>', '', '');return false;"><img src="images/delete.png" style="border:0px;" /></a>
                                    <a title="<%=IDM_ModifyQuestion %>" href="" onclick="editQuestion(<%=qst.ID %>, '<%=server.htmlencode(replace(replace(qst.Text,"\","\\"),"'","\'")) %>', '<%=qst.Weight %>', '<%=qst.IDRespType %>');return false;"><img src="images/edit.png" style="border:0px;" /></a>
                                <%end if %>
                            </td>
                            <td valign="top" style="border-top:1px solid black;"><font class="font12">
                                <%=qst.Text %>
                            </td>
                            <td valign="top"  style="border-top:1px solid black;" align="center"><font class="font12">
                                <%=qst.Weight %>
                            </td>
                            <td valign="top"  style="border-top:1px solid black;" align="center"><font class="font12">
                                <%if qst.IDRespType=0 then %>
                                    List
                                <%elseif qst.IDRespType=1 then %>
                                    Radio
                                <%end if %>
                            </td>
                        </tr>
                        
                        
                            <tr>
                                <td></td>
                                <td style="padding-left:20px;">
                                    <table width="100%" cellpadding="0" cellspacing="0">
                                        <tr height="25">
                                            <td width="35"><%if canModify then %><a title="<%=IDM_CreateResponse %>" href="" onclick="addResponse(TR_NewResp_<%=qst.ID %>, newRespValue_<%=qst.ID %>, newRespText_<%=qst.ID %>);return false;"><img src="images/nuevo.png" style="border:0px;" /></a><%end if %></td>
                                            
                                            <%if qst.numResponses > 0 then %>
                                                <td align="center"><font class="font10"><%=IDM_ResponseValue %></font></td>
                                                <td><font class="font10"><%=IDM_ResponseText %></font></td>
                                            <%else %>
                                                <td align="left" colspan=2><font class="font10"><%=IDM_NoResponseYet %></font></td>
                                            <%end if %>
                                        </tr>
                                        <%for each rsp in qst.Responses %>
                                            <tr>
                                                <td><%if canModify then %>
                                                        <a title="<%=IDM_DeleteResponse %>" href="" onclick="_fireConfirm('deleteResponse', '<%=rsp.ID %>', '', '');return false;"><img src="images/delete.png" style="border:0px;" /></a>
                                                    <%end if %>
                                                </td>
                                                <td width="50" align="center"><font class="font12"><%=rsp.RespValue %></td>
                                                <td><font class="font12"><%=rsp.Text %></td>
                                            </tr>
                                        <%next %>
                                        <%if canModify then %>
                                            <tr id="TR_NewResp_<%=qst.ID %>" style="display:none;">
                                                <td>
                                                    <a title="<%=IDM_Cancel %>" href="" onclick="cancelAddResponse(TR_NewResp_<%=qst.ID %>, newRespValue_<%=qst.ID %>, newRespText_<%=qst.ID %>);return false;"><img src="images/cancel.png" style="border:0px;" /></a>
                                                    <a title="<%=IDM_SaveResponse %>" href="" onclick="saveNewResponse('<%=qst.ID %>', newRespValue_<%=qst.ID %>, newRespText_<%=qst.ID %>);return false;"><img src="images/ok.png" style="border:0px;" /></a>
                                                </td>
                                                <td width="50" align="center"><font class="font10"><%=IDM_ResponseValue %><br /></font><input maxlength=3 style="width:30px;font-size:10px;" type=text name="newRespValue_<%=qst.ID %>" value="" /></td>
                                                <td><font class="font10"><%=IDM_ResponseText %><br /></font><input style="width:300px;font-size:10px;" type=text name="newRespText_<%=qst.ID %>" value="" /></td>
                                            </tr>
                                        <%end if %>
                                    </table>
                                </td>
                            </tr>
                    <%next %>
                </table>

            <%elseif Pestana = "2" then %>
                
                <div style="padding-left:20px;">
                    <font class="font15"><b><%=IDM_AssignedBrands %></b></font>
    	            <table width="100%" cellpadding="0" cellspacing="0">
    	                <%dim arrBrands, b, nBrands
    	                arrBrands = getFormBrands(idForm, "ASSIGNED_TO_FORM")
    	                nBrands = 0
    	                for each b in arrBrands
    	                %>
    	                    <tr>
    	                        <td width="30"><a title="<%=IDM_RemoveFromForm %>" href="" onclick="removeFromBrand('<%=b.IDBrand %>');return false;"><img src="images/remove.gif" style="border:0" /></a></td>
    	                        <td><font class="font12"><%=b.Name %></font></td>
                            </tr>
                            <%nBrands = nBrands + 1 %>
                        <%next %>
                        <%if nBrands = 0 then%>
    	                    <tr>
    	                        <td><font class="font10"><%=IDM_NoBrandsAssigned %></font></td>
                            </tr>
                        <%end if %>
                    </table>

                    <br />
                    <hr color="black" />
                    <font class="font15"><b><%=IDM_PromotionsAssigned %></b><br /></font>
                    <font class="font10"><%=IDM_PromotionsAssignedEx %>
                    </font>
    	            <table width="100%" cellpadding="0" cellspacing="0">
    	                <%
    	                arrBrands = getFormAssignedBrandPromotions(idForm)
    	                nBrands = 0
    	                dim nab
    	                for each nab in arrBrands
    	                    set b = nab.Brand
        	                %>
    	                    <tr>
    	                        <td width="30"><a title="<%=IDM_RemoveHistory %>" href="" onclick="removeFromPromotions('<%=b.IDBrand %>');return false;"><img src="images/remove.gif" style="border:0" /></a></td>
    	                        <td><font class="font12"><%=b.Name & " (" & nab.num & ")" %></font></td>
                            </tr>
                            <%nBrands = nBrands + 1 %>
                        <%next %>
                        <%if nBrands = 0 then%>
    	                    <tr>
    	                        <td><font class="font10"><%=IDM_NoPromotionsAssigned %></font></td>
                            </tr>
                        <%end if %>
                    </table>

                    <br />
                    <hr color="black" />
                    <font class="font15"><b><%=IDM_BrandsWithoutForm %></b></font>
    	            <table width="100%" cellpadding="0" cellspacing="0">
    	                <%
    	                arrBrands = getFormBrands(0, "NOT_ASSIGNED")
    	                nBrands = 0
    	                for each b in arrBrands
    	                %>
    	                    <tr>
    	                        <td width="30"><a title="<%=IDM_AddBrandToForm %>" href="" onclick="assignBrand('<%=b.IDBrand %>');return false;"><img src="images/add.gif" style="border:0" /></a></td>
    	                        <td><font class="font12"><%=b.Name %>  <%
    	                        if b.indBaja<>0 then
    	                            %><font color="red">&nbsp;&nbsp;(<%=IDM_Deleted %>)</font><%
    	                        end if
    	                         %></font></td>
                            </tr>
                            <%nBrands = nBrands + 1 %>
                        <%next %>
                        <%if nBrands = 0 then%>
    	                    <tr>
    	                        <td></td>
    	                        <td><font class="font10"><%=IDM_NoPendingBrands %></font></td>
                            </tr>
                        <%end if %>
                    </table>

                    <br />
                    <hr color="black" />
                    <font class="font15"><b><%=IDM_BrandsAssOtherForm %></b></font>
    	            <table width="100%" cellpadding="0" cellspacing="0">
    	                <%
    	                arrBrands = getFormBrands(idForm, "ASSIGNED_TO_OTHER_FORM")
    	                nBrands = 0
    	                for each b in arrBrands
    	                %>
    	                    <tr>
    	                        <td width="30"><a title="<%=IDM_AddBrandToForm %>" href="" onclick="reassignBrand('<%=b.IDBrand %>');return false;"><img src="images/add.gif" style="border:0" /></a></td>
    	                        <td><font class="font12"><%=b.Name %>  <%
    	                        if b.indBaja<>0 then
    	                            %><font color="red">&nbsp;&nbsp;(<%=IDM_Deleted %>)</font><%
    	                        end if
    	                         %></font></td>
                            </tr>
                            <%nBrands = nBrands + 1 %>
                        <%next %>
                        <%if nBrands = 0 then%>
    	                    <tr>
    	                        <td></td>
    	                        <td><font class="font10"><%=IDM_NoBrandsAssOtherForm %></font></td>
                            </tr>
                        <%end if %>
                    </table>
                    
                    <br /><br /><br />
                </div>
            <%end if %>
            
        <%end if %>
        
    </td></tr></table>
    
<br />
<br />
<br />
<br />


<input type="hidden" name="idForm" value="<%=idForm %>" />
<input type="hidden" id="Pestana" name="Pestana" value="<%=Pestana%>" />

<!-- #include file = "include/EventFunctions2.asp" -->


</FORM>