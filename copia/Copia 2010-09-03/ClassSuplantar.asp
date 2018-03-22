<%if request("Printable")="" AND request("XL")="" then%>
	<%dim Suplantar: Suplantar=Request("Suplantar")%>
	<%'Ho trec per poder fer proves durant la reunió.
	if UCASE(Session("IDUser"))="SQLMANAGER" OR UCASE(Session("IDUser"))="JJIMENEZ" OR UCASE(Session("IDUser"))="SLOPEZ" OR session("PuedeSuplantar")<>"" then 'Sólo lo puedo hacer yo y punto%>
		<div align=center valign=middle id="DIV_Suplantar1" style="top:0;right:6px;position:absolute;">
			<a title="Suplantar" href="" onclick="DIV_Suplantar.style.display='';DIV_Suplantar1.style.display='none';form1.Suplantar.focus();return false;">-</a>
		</div>
		<div align=center valign=middle id="DIV_Suplantar" style="display:none;top:20;left:550;position:absolute;background-color:black;border:4 solid black">
			<form id=form1 name=form1 LANGUAGE=javascript>
				<input style="background-color:black;font-weight:bold;font-size:10px;color:white;width:50px" type="text" name="Suplantar" value="<%'=Suplantar%>">
				<br>
				<input class="button" type="submit" value="Sup" id=submit1 name=submit1>
				<%if Suplantar<>"" then
					Session("IDUser") = Suplantar
					session("PuedeSuplantar") = 1
					session("UserFullName") = ""
				end if%>
				<input type="button" class="button" value="Cerrar" onclick="DIV_Suplantar.style.display='none';DIV_Suplantar1.style.display='';document.cookie='ViewSup=No';" id=button1 name=button1></form>
		</div>
	<%end if%>
<%end if%>
