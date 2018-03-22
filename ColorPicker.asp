



<HEAD>
<TITLE>Color Picker</TITLE>

	<script langauge="javascript">
	// Author: Dion, biab@iinet.net.au, http://biab.howtojs.com/
	// Courtesy of SimplytheBest.net - http://simplythebest.net/scripts/

	addary=new Array(255,1,1);
	clrary=new Array(360);
	for(i=0;i<6;i++)
	 for(j=0;j<60;j++)
	  { clrary[60*i+j]=new Array(3);
	    for(k=0;k<3;k++)
	     { clrary[60*i+j][k]=addary[k];
	       addary[k]+=((Math.floor(65049010/Math.pow(3,i*3+k))%3)-1)*4; }; };

	function capture()
	 { if(document.layers)
	    { with(document.layers['imgdiv'])
	       { document.captureEvents(Event.MOUSEMOVE);
	         document.onmousemove=moved; }; }
	   else { document.all.imgdiv.onmousemove=moved; };
	 };

	function moved(e)
	 { sx=((document.layers)?e.layerY:event.offsetY)-256;
	   sy=((document.layers)?e.layerX:event.offsetX)-256;
	   quad=new Array(-180,360,180,0);
	   xa=Math.abs(sx); ya=Math.abs(sy);
	   d=ya*45/xa;
	   if(ya>xa) d=90-(xa*45/ya);
	   deg=Math.floor(Math.abs(quad[2*((sy<0)?0:1)+((sx<0)?0:1)]-d));
	   n=0; c="000000";
	   r=Math.sqrt((xa*xa)+(ya*ya));
	   if(sx!=0 || sy!=0)
	    { for(i=0;i<3;i++)
	       { r2=clrary[deg][i]*r/128;
	         if(r>128) r2+=Math.floor(r-128)*2;
	         if(r2>255) r2=255;
	         n=256*n+Math.floor(r2); };
	      c=(n.toString(16)).toUpperCase();
	      while(c.length<6) c="0"+c; };
	   if(document.layers)
	    { document.layers['clrdiv'].bgColor="#"+c; }
	   else
	    { document.all["clrdiv"].style.backgroundColor="#"+c; };
	   document.frm.txt.value="#"+c;
	   document.frm.hid.value="#"+c;
	   return false; };
	
	
	function setcolor()
	 {
		document.frm.sel.style.background=document.frm.hid.value;
	 };
	 
	 function confirmcolor(){
		window.opener.thisForm.<%=request("campo")%>.value=document.frm.sel.style.background;
		window.opener.thisForm.<%=request("campo")%>.style.background=document.frm.sel.style.background;
		window.close();
	 }
	</script>
	<style type="text/css">
	#imgdiv { position:relative; }
	#clrdiv { position:relative; background-color:white; }
	</style>

</HEAD>

<HTML>


<BODY onload="capture();" topmargin=0 leftmargin=0>
	
	<table border=1 cellpadding=0 cellspacing=0>
		<tr><td><div id=imgdiv><a href="javascript:void(0);" onclick="setcolor(); return false;">
		  <img src="images/colorwheel.jpg" width=512 height=512 border=0></a></div></td></tr>
		<tr><td align="center"><div id=clrdiv></div></td></tr>
		<tr>
			<form name="frm">
				<td align="center"><input type="hidden" name="txt" size=12>
				  <input type="text" name="sel" size=12>
				  <input type="hidden" name="hid">
				  <input type=button value="SELECT" onclick="confirmcolor();return false;">
				</td>
			</form>
		</tr>
</table>


	
</BODY>
</HTML>