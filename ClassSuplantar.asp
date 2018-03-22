
<%if puedeSuplantar OR session("PuedeSuplantar")<>"" then%>
    <div align="center" valign="middle" id="DIV_Suplantar" style="z-index:1000;display:none;top:65px;right:55px;position:absolute;background-color:black;border:4px solid black">
        <form id="form1" name="form1" action="" method="post">
	        <input style="background-color:black;font-weight:bold;font-size:10px;color:white;width:50px" type="text" name="Suplantar" value="" />
	        <br>
	        <input class="button" type="submit" value="Sup" id="submit1" name="submit1" />
	        <%if Request("Suplantar")<>"" then
		        Session("IDUser") = Request("Suplantar")
		        session("PuedeSuplantar") = 1
		        session("UserFullName") = ""
	        end if%>
	        <input type="button" class="button" value="Cerrar" onclick="document.getElementById('DIV_Suplantar').style.display='none';document.cookie='ViewSup=No';" id="button1" name="button1" />
	    </form>
    </div>
<%end if %>
