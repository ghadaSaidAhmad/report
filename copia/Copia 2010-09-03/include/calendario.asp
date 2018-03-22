<%
campo=request("campo")
'ultima = request("ultimaPos")
anio = request("anio")
mes = request("contMes")
if(mes="") then
	mes = month(date())
end if
if(anio="") then
	anio=year(date())
end if
	fecha = DateSerial(anio, mes, 1)
	
	FirstDayOfWeek = 2
	if UCASE(Application("FirstDayOfWeek")) = "MONDAY" then
		FirstDayOfWeek = 2
	elseif UCASE(Application("FirstDayOfWeek")) = "SUNDAY" then
		FirstDayOfWeek = 1
	end if
	
	primera = datepart("w",fecha, FirstDayOfWeek)

%>

<HTML>
<HEAD>

<LINK REL=StyleSheet HREF="style.css" TYPE="text/css">

<title>Calendar</title>
<script language="JavaScript">


	var nDiaSemana = 0;
	var nAnio = 0;
	var nU = 0;
	var DiaHoy = 0;
	var nUltima = 0;
	var Meses = new Array(13);
	Meses [1]='January'
	Meses[2] = 'February'
	Meses [3] = 'March'
	Meses[4]= 'April'
	Meses [5] = 'May'
	Meses[6] ='June'
	Meses [7] = 'July'
	Meses [8] = 'August'
	Meses [9] = 'September'
	Meses [10]= 'October'
	Meses [11] = 'November'
	Meses [12] = 'December'
	
	dia=new Date();
	//m=dia.getMonth();
	m = 1;
	var MatMes = new Array(13);	
	for (var i = 0; i<=13; i++) {		
		MatMes[i] = new Array(3);
	}
	
		MatMes[1][1] = Meses [1]
		MatMes[1][2] = 1
		MatMes[2][1] = Meses[2]
		MatMes[2][2] = 2
		MatMes[3][1] = Meses[3]
		MatMes[3][2] = 3
		MatMes[4][1] = Meses[4]
		MatMes[4][2] = 4
		MatMes[5][1] = Meses[5]
		MatMes[5][2] = 5
		MatMes[6][1] = Meses[6]
		MatMes[6][2] = 6
		MatMes[7][1] = Meses[7]
		MatMes[7][2] = 7
		MatMes[8][1] = Meses[8]
		MatMes[8][2] = 8
		MatMes[9][1] = Meses[9]
		MatMes[9][2] = 9
		MatMes[10][1] = Meses[10]
		MatMes[10][2] = 10
		MatMes[11][1] = Meses[11]
		MatMes[11][2] = 11
		MatMes[12][1] = Meses[12]
		MatMes[12][2] = 12
		fecha = new Date();
	DiaHoy = fecha.getDate();
	
	function ImprimeCal(Pos, nCdias)
	{
		var nDia = 0;
		var mes = "";
		var dia = "";
		var sTabla = "";
		var valorTD = "";
		var Fecha = "";
		var bEscribioVacia = false;
		Pos--;
		//alert(Pos)
		sTabla = '<tr>'
		for (nFilaMes = 1;nFilaMes<=6;nFilaMes++)
		{
			for (nDiaSem = 1;nDiaSem<=7;nDiaSem++)
			{
				if (nFilaMes == 1 && !bEscribioVacia && Pos != 7)
				{
					for (nCellVacia = 1;nCellVacia <= Pos; nCellVacia++)
					{
						sTabla += '<td>&nbsp;</td>'
					}
					nDiaSem = parseInt(Pos)+1
					bEscribioVacia = true
				}
				if (nDia == nCdias)
				{
					sTabla += '<td>&nbsp;</td>'
				}
				else
				{
					nDia++;
					if(parseInt(MatMes[document.hid.contMes.value][2])<10)
						mes = '0' + MatMes[document.hid.contMes.value][2];
					else
						mes = MatMes[document.hid.contMes.value][2];
					if(nDia<10)
						dia = '0' + nDia;	
					else
						dia = nDia
					<%if Application("USDateFormatHover")="YES" then%>	
					Fecha = mes + '/' + dia + '/' + document.hid.Anio.value
					<%else%>
					Fecha = dia + '/' + mes + '/' + document.hid.Anio.value
					
					<%end if%>
					valorTD = '<td align="center"><a href="JavaScript:Copiovalor(\''+Fecha+'\')"><font class="CalDiaLink">'+nDia+'</a></a></td>'
							sTabla += valorTD
					nU = nDiaSem
				}	
			}
			if (nFilaMes == 6) 
				{
					sTabla += '</tr>'
				} 
			else 
				{
					sTabla += '</tr><tr>'
				}
		}
		document.hid.ultimaPos.value=nU
	
		
		document.hid.mes.value=MatMes[document.hid.contMes.value][2];
	
		
		return sTabla;
		
	}
	
	
	function Copiovalor(fecha)
	{
		window.opener.thisForm.<%=campo%>.value = fecha;
		window.close()
	}
	
	
	function SelMes(){
		
		document.hid.contMes.value = parseInt(document.hid.mes.value);
		document.hid.submit()
		
	}
	
	function MesNext()
	{
		document.hid.contMes.value = parseInt(document.hid.contMes.value) + 1;
		if (document.hid.contMes.value == 13)
		{
			document.hid.contMes.value = 1;
			document.hid.Anio.value = parseInt(document.hid.Anio.value) + 1
		}	
		document.hid.NextPrev.value = 'next'
		document.hid.submit()
	}
	function MesPrev()
	{
			if(document.hid.contMes.value==1){
				document.hid.contMes.value = 12;
				document.hid.Anio.value = parseInt(document.hid.Anio.value) - 1;
				}
			else{
				document.hid.contMes.value = parseInt(document.hid.contMes.value) - 1;
				}
			document.forms.hid.NextPrev.value = 'prev'
			document.hid.submit()
		//}
	}
	function GetNombreMes(mes)
	{
		return MatMes[mes][1] 
	}
	function GetCantDias(mes)
	{
		nAnio = document.hid.Anio.value
		if (mes == 4 || mes == 6 || mes == 9 || mes == 11)
		{
			nCantiDias = 30
		}
		else
		{
			if ((mes == 2) && ((nAnio % 4 == 0) || (nAnio % 100 == 0)))
			{
				nCantiDias = 29
			}
			else
			{
				if ((mes == 2) && ((nAnio % 4 != 0) || (nAnio % 100 != 0)))
				{
					nCantiDias = 28
				}
				else
				{
					nCantiDias = 31
				}
			}
		}
		return nCantiDias; 
	}


</script>

</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" bottommargin="0" valign="top">


     <form name="hid" method="post" action="Calendario.asp">
		<input type="hidden" name="contMes" value="<%=mes%>">
		<input type="hidden" name="contTot" value="1">
		
		<input type="hidden" name="ultimaPos" value="<%=ultima%>">
		<input type="hidden" name="Anio" value="<%=Anio%>">
		<input type="hidden" name="campo" value="<%=campo%>">
		<input type="hidden" name="NextPrev" value="">
		<input type="hidden" name="cualform" value="">
	
	

<table  width="148" border="0" cellspacing="0" cellpadding="1" bgcolor="#669999">
  <tr >
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#ffffff">
        <tr> 
          <td colspan="3" bgcolor="#99cccc"></td>
        </tr>
        <tr> 
          <td rowspan="8" width="1%" bgcolor="#669999"></td>
          <td width="98%"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td  bgcolor="#669999" valign="middle">
					<%if true then%>
						<a href="JavaScript:MesPrev()" STYLE="Text-Decoration:none"><b> <img width=17 height=13 src="../images/calPrev.gif" border=0 alt="Previous Month"></b></a>
					<%end if%>
				</td>
                <td width="60%"  bgcolor="#669999" align="center"> 
                    <font color="#ffffff" class="CalDiaNoLink"> 
                    <select onchange="SelMes();" id="mes" name="mes" style="width:100px;font-size:10px;">
						<%for i=1 to 12%>
							<option <%if cint(i)=cint(mes) then%>selected<%end if%> value="<%=i%>"><%=MonthName(i) & " " & Anio%></option>
						<%next%>
					</select>
                    <%if false then%>
						<script language="JavaScript">
							document.write(GetNombreMes(document.hid.contMes.value)+' ' +document.hid.Anio.value )
						</script>
					<%end if%>
                 </font>
                  </font>
                </td>
                <td  bgcolor="#669999" align=right>
					<%if true then%>
						<a href="JavaScript:MesNext()" STYLE="Text-Decoration:none"><b> <img src="../images/calNext.gif" width=17 height=13 border=0 alt="Next Month"></b></a>&nbsp;
					<%end if%>
					</td>
						
              </tr>
            </table>
          </td>
          
        </tr>
        <tr> 
          <td></td>
        </tr>
        <tr> 
          <td></td>
        </tr>
        <tr> 
          <td></td>
        </tr>
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr> 
				<%if FirstDayOfWeek=1 then%>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Sun</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Mon</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Tue</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Wed</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Thu</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Fri</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Sat</td>
				<%elseif FirstDayOfWeek=2 then%>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Mon</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Tue</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Wed</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Thu</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Fri</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Sat</td>
					<td bgcolor="#99cccc" class="CalDiaNoLink">Sun</td>
				<%end if%>
              </tr>
              <tr>
              <td colspan="7" bgcolor="#ffffff"><font size="2">
              	<script language="javascript">
					document.write(ImprimeCal(<%=primera%>,GetCantDias(MatMes[document.hid.contMes.value][2])));
				</script>  
			</font>
              </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#ffffff"></td>
        </tr>
        <tr> 
          <td align="center" class="CalDiaNoLink" bgcolor="#99cccc"><a href="JavaScript:window.close()"><B>Close</B></a></td>
        </tr>
        <tr> 
          <td bgcolor="#99cccc">
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>

	</form>

</body>
</html>


<script language="JavaScript">window.focus();</script>