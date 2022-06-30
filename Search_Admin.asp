<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, Conn, rs, CountList1Values, CountList2Values, CountList3Values, CountList4Values, CountList5Values, CountList6Values
Dim aList1Values, aList2Values, aList3Values, aList4Values, aList5Values, aList6Values, i

	GroupID = CheckNum(Request("GID")) 'Revisando que el Grupo sea 1 = Categorias, 2 = Noticias , 3 = Mensajes, 4 = Usuarios
	'Configurando La pagina de Posteo
	CountList1Values = -1
	CountList2Values = -1
	CountList3Values = -1
	CountList4Values = -1
	CountList5Values = -1
    CountList6Values = -1
	
	OpenConn Conn
	'Obteniendo listado de Carriers
	Set rs = Conn.Execute("select CarrierID, Name, Countries from Carriers where Expired=0 and Countries in " & Session("Countries") & " order by Name, Countries")
	If Not rs.EOF Then
   		aList1Values = rs.GetRows
       	CountList1Values = rs.RecordCount-1
    End If
	CloseOBJ rs
	'Obteniendo listado de Aeropuertos
	Set rs = Conn.Execute("select AirportID, Name, AirportCode from Airports where Expired=0 order by Name")
	If Not rs.EOF Then
   		aList3Values = rs.GetRows
       	CountList3Values = rs.RecordCount-1
    End If
	CloseOBJs rs, Conn

	OpenConn2 Conn
	'Obteniendo listado de Agentes
	'Set rs = Conn.Execute("select AgentID, Name, Countries from Agents where Expired=0 and Countries in " & Session("Countries") & " order by Countries, Name")
	Set rs = Conn.Execute("select agente_id, agente from agentes where activo=true order by agente")
	If Not rs.EOF Then
   		aList2Values = rs.GetRows
       	CountList2Values = rs.RecordCount-1
    End If
	CloseOBJ rs

	if GroupID = 16 then
		''Obteniendo listado de Embarcadores
        ''Set rs = Conn.Execute("select ShipperID, Name, Countries from Shippers where Expired=0 and Countries in " & Session("Countries") & " order by Countries, Name")
		'Set rs = Conn.Execute("select a.id_cliente, a.nombre_cliente, p.codigo " & _
		'					"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
		'					" where a.id_cliente = d.id_cliente" & _
		'					" and d.id_nivel_geografico = n.id_nivel" & _
		'					" and n.id_pais = p.codigo" & _
		'					" and a.es_shipper=true order by p.codigo, a.nombre_cliente")
		'If Not rs.EOF Then
		'	aList4Values = rs.GetRows
		'	CountList4Values = rs.RecordCount-1
		'End If
		'CloseOBJ rs
	End If

	'Set rs = Conn.Execute("select u.id_usuario, u.nombre, u.id_pais from usuarios u, perfiles_usuarios p where u.id_usuario = p.id_usuario and p.id_perfil=4 and u.id_pais in " & Session("Countries") & " order by u.id_pais, u.nombre")
	Set rs = Conn.Execute("select id_usuario, pw_gecos, pais from usuarios_empresas where tipo_usuario=1 and pais in " & Session("Countries") & " order by pais, pw_gecos")
	If Not rs.EOF Then
		aList5Values = rs.GetRows
		CountList5Values = rs.RecordCount-1
	End If

    CloseOBJ rs

    if GroupID = 16 then 'este catalogo es utilizado unicamente por estadistica de mediciones 2016-03-15

        Dim SQLQuery 
        SQLQuery = "SELECT pais, nombre FROM usuarios_paises WHERE activo = 't' AND pais IN " & Session("Countries") & " ORDER BY nombre LIMIT 50"
        'response.write ( SQLQuery )    
	    Set rs = Conn.Execute(SQLQuery)
	    If Not rs.EOF Then
		    aList6Values = rs.GetRows
		    CountList6Values = rs.RecordCount-1
	    End If
    
    end if

	CloseOBJs rs, Conn


%>

<HTML><HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<!--<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>-->
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function abrir(Label){
	var DateSend, Subject;
		if (parseInt(navigator.appVersion) < 5) {
			DateSend = document.forma(Label).value;
		} else {
			DateSend = document.getElementById(Label).value;
		}
		Subject = '';	
		window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
	}
	
	function SetFilter() {
		var TextFilter = "";	
		if (document.forma.AWBNumber.value != "") {
			TextFilter += ", AWB: " + document.forma.AWBNumber.value;
		}
		if (document.forma.CarrierID.value != "") {
			TextFilter += ", Transportista: " + document.forma.CarrierID.options[document.forma.CarrierID.selectedIndex].text;
		}
		if (document.forma.AgentID.value != "") {
			TextFilter += ", Agente: " + document.forma.AgentID.options[document.forma.AgentID.selectedIndex].text;
		}
		if (document.forma.SalespersonID.value != "") {
			TextFilter += ", Vendedor: " + document.forma.SalespersonID.options[document.forma.SalespersonID.selectedIndex].text;
		}
		if (document.forma.AirportDepID.value != "") {
			TextFilter += ", Salida: " + document.forma.AirportDepID.options[document.forma.AirportDepID.selectedIndex].text;
		}
		if (document.forma.AirportDesID.value != "") {
			TextFilter += ", Destino: " + document.forma.AirportDesID.options[document.forma.AirportDesID.selectedIndex].text;
		}
		<%if GroupID<>16 then%>
		if (document.forma.DateFrom.value != "") {
			TextFilter += ", Desde: " + document.forma.DateFrom.value;
		}
		if (document.forma.DateTo.value != "") {
			TextFilter += ", Hasta: " + document.forma.DateTo.value;
		}
		<%end if%>
		if (document.forma.AwbType.value == 1) {
			TextFilter = "Resultados para EXPORT" + TextFilter;
		} else {
			TextFilter = "Resultados para IMPORT" + TextFilter;
		}
		document.forma.TextFilter.value = TextFilter;
	}
	
	function validate(){
		<%select case GroupID
		Case 1, 4, 6, 15, 17, 18, 22%>
			if (!valSelec(document.forma.AwbType)){return (false)};
		<%Case 16 %>

        var elt = document.forma.MMFrom;
        document.forma.MMFromText.value = elt.options[elt.selectedIndex].text;

        elt = document.forma.MMTo;
        document.forma.MMToText.value = elt.options[elt.selectedIndex].text;
        
			if (!valSelec(document.forma.AwbType)){return (false)};
			if (document.forma.ResultType.value == 0) {
                if (!valSelec(document.forma.ReportType)){return (false)};
            }
            
            
            if (document.forma.ResultType.value == 2 || document.forma.ResultType.value == 3)
            if (document.forma.Countries.value ==  '') {
                alert('Seleccione pais');
                document.forma.Countries.focus();
                return false;
            }

	 	<%end select%>
		<%if GroupID=16 then%>
		SetFilter();
		<%end if%>
		document.forma.submit();
	}
	 	 
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<FORM name="forma" action="Search_ResultsAdmin.asp" method="post" target=_self>
	<INPUT name="MMFromText" type=hidden value="">
    <INPUT name="MMToText" type=hidden value="">
    <INPUT name="GID" type=hidden value="<%=GroupID%>">
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
	  	<%if GroupID<>16 then%>
		<TR>
		<TD width=40% class=label align=right valign=top><b>Rango de fechas:</b><br>(dd-mm-yyyy)</TD>
		<TD width=60% class=label align=left>Desde:<br><INPUT  readonly="readonly" name="DateFrom" id="DateFrom" type=text value="" size=23 maxLength=19 class=label>&nbsp;
			<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('DateFrom');" class=label><br>
			Hasta:<br><INPUT  readonly="readonly" name="DateTo" id="DateTo" type=text value="" size=23 maxLength=19 class=label>&nbsp;
			<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('DateTo');" class=label><br>
		</TD>
	 	</TR>
		<%else%>
		<TR>
		<TD width=40% class=label align=right valign=top><b>Rango de fechas:</b></TD>
		<TD width=60% class=label align=left>Desde:<br>
        <input type="hidden" name="TextFilter" value="">
		<select name="MMFrom" class="label">
		<option value=1>ENE</option>
		<option value=2>FEB</option>
		<option value=3>MAR</option>
		<option value=4>ABR</option>
		<option value=5>MAY</option>
		<option value=6>JUN</option>
		<option value=7>JUL</option>
		<option value=8>AGO</option>
		<option value=9>SEP</option>
		<option value=10>OCT</option>
		<option value=11>NOV</option>
		<option value=12>DIC</option>
		</select><select name="YYFrom" class=label>
        <% For i = 2007 To Year(date)
            response.write ( "<option value='" & i & "'>" & i & "</option>" ) 	        
        next %>
		</select><br>
			Hasta:<br><select name="MMTo" class="label">
		<option value=1>ENE</option>
		<option value=2>FEB</option>
		<option value=3>MAR</option>
		<option value=4>ABR</option>
		<option value=5>MAY</option>
		<option value=6>JUN</option>
		<option value=7>JUL</option>
		<option value=8>AGO</option>
		<option value=9>SEP</option>
		<option value=10>OCT</option>
		<option value=11>NOV</option>
		<option value=12>DIC</option>
		</select>        
        <select name="YYTo" class=label>
        <% For i = 2007 To Year(date)
            response.write ( "<option value='" & i & "'>" & i & "</option>" ) 	        
        next %>
		</select>
		</TD>
	 	</TR>
		<%end if%>
		<%if GroupID = 16 then%>
        <TR>
		<TD class=label align=right width=40%><b>Tipo de Resultados:</b></TD>
		<TD class=label align=left width=60%>
		<select name="ResultType" class=label id="Clase de Resultados">
			<option value=0>COMISIONES</option>
			<option value=1>CARGA</option>
            <option value=2>MEDICIONES</option>
            <option value=3>BITACORA</option>
		</select>
		</TD>
		</TR>
		<%end if%>
        <%select case GroupID
		Case 1, 4, 6, 15, 16, 17, 18, 22 'AWB %>
		<TR>
		<TD class=label align=right width=40%><b>Tipo de AWB:</b></TD>
		<TD class=label align=left width=60%>
		<select name="AwbType" class=label id="Tipo de AWB">
			<option value="-1">Seleccionar</option>
			<option value=1>EXPORT</option>
			<option value=2>IMPORT</option>
		</select>
		</TD>
		</TR>
        <%if GroupID = 16 then%>
		<TR>
		<TD class=label align=right width=40%><b>Tipo de REPORTE:</b></TD>
		<TD class=label align=left width=60%>
		<select name="ReportType" class=label id="Tipo de REPORTE">
			<option value="-1">Seleccionar</option>
			<option value=0>POR TRANSPORTISTA</option>
			<option value=1>POR AGENTE</option>
		</select>
		</TD>
		</TR>
        <%end if%>


		<% if GroupID = 16  then 'Mediciones %>

        <TR>
        <TD class=label align=right width=40%><b>Pais:</b></TD>
        <TD class=label align=left>
        <select name="Countries" class=label>
		<option value="">Seleccionar</option>
		<%
			For i = 0 To CountList6Values
		%>
		<option value="<%=aList6Values(0,i)%>"><%=aList6Values(1,i)%></option>
		<%
    		Next
		%>
		</select>
        </TR>
        <% end if %>

        <TR>
		<TD class=label align=right width=40%>
		<%select case GroupID
		 case 6, 18%>
			<b>No. de House AWB:</b>
		<%case 16, 17%>
			<b>No. de Master AWB:</b>
		<%case else%>
			<b>No. MAWB o HAWB:</b>
		<%end select%>
		</TD>
		<TD class=label align=left width=60%><INPUT name="AWBNumber" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Vendedor:</b></TD>
		<TD class=label align=left width=60%>
		<select name="SalespersonID" class=label>
		<option value="">Seleccionar</option>
		<%
			For i = 0 To CountList5Values
		%>
		<option value="<%=aList5Values(0,i)%>"><%=aList5Values(2,i) & " - " & aList5Values(1,i)%></option>
		<%
    		Next
		%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Transportista:</b></TD>
		<TD class=label align=left width=60%>
		<select name="CarrierID" class=label>
		<option value="">Seleccionar</option>
		<%
			For i = 0 To CountList1Values
		%>
		<option value="<%=aList1Values(0,i)%>"><%=aList1Values(1,i) & " - " & aList1Values(2,i) & " - " & aList1Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Agente:</b></TD>
		<TD class=label align=left width=60%>
		<select name="AgentID" class=label>
		<option value="">Seleccionar</option>
		<%
			For i = 0 To CountList2Values
		%>
		<option value="<%=aList2Values(0,i)%>"><%=aList2Values(1,i)%></option>
		<%
    		Next
		%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Aeropuerto Salida:</b></TD>
		<TD class=label align=left width=60%>
		<select name="AirportDepID" class=label>
		<option value="">Seleccionar</option>
		<%
			For i = 0 To CountList3Values
		%>
		<option value="<%=aList3Values(0,i)%>"><%=aList3Values(1,i) & " - " &  aList3Values(2,i) & " - " &  aList3Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Aeropuerto Destino:</b></TD>
		<TD class=label align=left width=60%>
		<select name="AirportDesID" class=label>
		<option value="">Seleccionar</option>
		<%
			For i = 0 To CountList3Values-1
		%>
		<option value="<%=aList3Values(0,i)%>"><%=aList3Values(1,i) & " - " &  aList3Values(2,i) & " - " &  aList3Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
		</TD>
		</TR>
		<% Case 2 'Transportistas - Shippers	%>
		<TR>
		<TD class=label align=right width=40%><b>Nombre:</b></TD>
		<TD class=label align=left width=60%><INPUT name="Name" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>C&oacute;digo Transportista:</b></TD>
		<TD class=label align=left width=60%><INPUT name="CarrierCode" type=text value="" size=30 maxLength=50 class=label></TD>
	  	</TR>
 		<% Case 3 'Transportistas-Salida	%>
		<TR>
		<TD class=label align=right width=40%><b>Transportista:</b></TD>
		<TD class=label align=left width=60%>
		<select name="CarrierID" class=label>
		<option value="">Seleccionar</option>
		<%
			For i = 0 To CountList1Values
		%>
		<option value="<%=aList1Values(0,i)%>"><%=aList1Values(1,i) & " - " & aList1Values(2,i) & " - " & aList1Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
	    </TR>
		<TR>
		<TD class=label align=right width=40%><b>Aeropuerto Salida:</b></TD>
		<TD class=label align=left width=60%>
		<select name="AirportID" class=label>
		<option value="">Seleccionar</option>
		<%
			For i = 0 To CountList3Values-1
		%>
		<option value="<%=aList3Values(0,i)%>"><%=aList3Values(1,i) & " - " & aList3Values(2,i) & " - " & aList3Values(0,i)%></option>
		<%
    		Next
		%>
		</select></TD>
	    </TR>		
		<% Case 5 'Transportistas-Rango %>
		<TR>
		<TD class=label align=right width=40%><b>Transportista:</b></TD>
		<TD class=label align=left width=60%>
		<select name="CarrierID" class=label>
		<option value="">Seleccionar</option>
		<%
			For i = 0 To CountList1Values
		%>
		<option value="<%=aList1Values(0,i)%>"><%=aList1Values(1,i) & " - " & aList1Values(2,i) & " - " & aList1Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
</TD>
	    </TR>
		<TR>
		<TD class=label align=right width=40%><b>Rango:</b></TD>
		<TD class=label align=left width=60%><INPUT name="RangeID" type=text value="" size=30 maxLength=50 class=label></TD>
	    </TR>		
		<% Case 7, 8, 10 'Consigners, Agents, Shippers	%>
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR> 
		<% Case 9 'Aeropuertos %>
		<TR>
		<TD class=label align=right><b>Nombre:</TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR>
		<TR>
		<TD class=label align=right><b>C&oacute;digo Aeropuerto:</TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="AirportCode" value="" size=30></TD>
		</TR>
		<% Case 11 'Commodities %>
		<TR>
		<TD class=label align=right><b>Nombre:</TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR>
		<% Case 12 'Monedas %>
		<TR>
		<TD class=label align=right><b>Nombre:</TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR>
		<TR>
		<TD class=label align=right><b>C&oacute;digo Moneda:</TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="CurrencyCode" value="" size=30></TD>
		</TR>
		<% Case 13 'Rangos %>
		<TR>
		<TD class=label align=right><b>Rango:</TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Val" value="" size=30></TD>
		</TR>
		<% Case 14 'Impuestos %>
		<TR>
		<TD class=label align=right><b>Valor:</TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Tax" value="" size=30></TD>
		</TR>
        <%
		end select
		%>
		</TABLE>
		<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
				<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validate();" value="&nbsp;&nbsp;Buscar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</TABLE>
	</FORM>
<%if GroupID=16 then%>
<SCRIPT>
    var now = new Date();
    selecciona('forma.MMFrom', now.getMonth() + 1);
    selecciona('forma.YYFrom', now.getFullYear());
    selecciona('forma.MMTo', now.getMonth() + 1);
    selecciona('forma.YYTo', now.getFullYear());
</SCRIPT>	
<%end if%>
</BODY>
</HTML>
<%
Set aList1Values = Nothing
Set aList2Values = Nothing
Set aList3Values = Nothing
Set aList4Values = Nothing
Set aList5Values = Nothing

Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
