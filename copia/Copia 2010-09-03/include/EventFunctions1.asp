<%
dim EventObject, EventParam1, EventParam2, PageReloaded
EventObject = Request.Form("EventObject")
EventParam1 = Request.Form("EventParam1")
EventParam2 = Request.Form("EventParam2")
PageReloaded = Request.Form("PageReloaded")
%>
<script language="JavaScript">
function _fireEvent (Objeto, Param1, Param2)
{	
	thisForm.EventObject.value = Objeto;
	thisForm.EventParam1.value = Param1;
	thisForm.EventParam2.value = Param2;			
	thisForm.submit();
}

function _fireConfirm(Objeto, Param1, Param2, MSG)
{
	if (MSG!=""){
		if (confirm(MSG)){
			_fireEvent(Objeto,Param1,Param2);
		}
	}
	else if (window.confirm("Click OK to continue. Click Cancel to abort.")){
		_fireEvent(Objeto,Param1,Param2);
	}
}

function ajaxReq(tipo, dat)
{
	urlCall = "AjaxReq.asp?T=" + tipo + "&D=" + dat// + "&A=" + second(Time)
	//alert(urlCall);
	xmlhttp=null
	// code for Mozilla, etc.
	if (window.XMLHttpRequest)
	  {
	  xmlhttp=new XMLHttpRequest()
	  }
	// code for IE
	else if (window.ActiveXObject)
	  {
	  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP")
	  }
	if (xmlhttp!=null)
	  {
	  xmlhttp.open("GET",urlCall,false)
	  xmlhttp.send(null)
	  }
	else
	  {
	  alert("Your browser does not support XMLHTTP.")
	  }
	
	if (xmlhttp.status==200){
		return xmlhttp.responseText
	}else{
		return null
	}
	
}
</script>
