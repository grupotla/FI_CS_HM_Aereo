<%
Checking "0|1"
if CountTableValues >= 0 then
	CarrierID = aTableValues(0, 0)
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	AirportID = aTableValues(4, 0)
end if
CountList1Values = -1
CountList2Values = -1
	if ObjectID = 0 then
		SQLQuery = "select CarrierID, Name, Countries from Carriers where Expired = 0 order by Name, Countries"
		SQLQuery2 = "select AirportID, Name, AirportCode from Airports where Expired = 0 order by Name Asc"
	else
		SQLQuery = "select CarrierID, Name, Countries from Carriers where Expired = 0 and CarrierID=" & CarrierID
		SQLQuery2 = "select AirportID, Name, AirportCode from Airports where Expired = 0 and AirportID=" & AirportID
	End If

OpenConn Conn
		'Obteniendo listado de Carriers
		Set rs = Conn.Execute(SQLQuery)
		If Not rs.EOF Then
    		aList1Values = rs.GetRows
        	CountList1Values = rs.RecordCount
	    End If
		CloseOBJ rs
		'Obteniendo listado de Aeropuertos
		Set rs = Conn.Execute(SQLQuery2)
		If Not rs.EOF Then
    		aList2Values = rs.GetRows
        	CountList2Values = rs.RecordCount
	    End If
CloseOBJs rs, Conn
%>
<HTML>
<HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
		<%
		if ObjectID = 0 then
		%>
		if (!valSelec(document.forma.CarrierID)){return (false)};
		if (!valSelec(document.forma.AirportID)){return (false)};
		<%
		end if
		%>
	    document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creación:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
        <TR><TD class=label align=right><b>Activa:</b></TD><TD class=label align=left><INPUT name=Expired TYPE=checkbox class=label <%If Expired = False Then response.write " checked"  End If%>></TD></TR>
		<TR><TD class=label align=right><b>Transportista: *</b></TD><TD class=label align=left>
		<%
		if ObjectID = 0 then
		%>
		<select name="CarrierID" id="Transportista" class=label>
		<option value="-1">Seleccionar</option>
		<%
			For i = 0 To CountList1Values-1
		%>
		<option value="<%=aList1Values(0,i)%>"><%=aList1Values(1,i) & " - " & aList1Values(2,i) & " - " & aList1Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
		<%
    	Else
		%>
		<%=aList1Values(1,0) & " - " & aList1Values(2,0) & " - " & aList1Values(0,0)%>
		<input type="hidden" name="CarrierID" id="Transportista" value="<%=aList1Values(0,0)%>">
    	<%
		End If
		%>
		</TD></TR> 

		<TR><TD class=label align=right><b>Aeropuerto:</b></TD><TD class=label align=left>
		<%
		if ObjectID = 0 then
		%>
		<select name="AirportID" id="Código Aeropuerto" class=label>
		<option value="-1">Seleccionar</option>
		<%
			For i = 0 To CountList2Values-1
		%>
		<option value="<%=aList2Values(0,i)%>"><%=aList2Values(1,i) & " - " & aList2Values(2,i) & " - " & aList2Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
				<%
    	Else
		%>
		<%=aList2Values(1,0) & "-" & aList2Values(2,0) & " - " & aList2Values(0,0)%>
		<input type="hidden" name="AirportID" id="Código Aeropuerto"  value="<%=aList2Values(0,0)%>">
    	<%
		End If
		%>
		</TD></TR> 
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
				  <TR>
							<%if CountTableValues = -1 then%>
									 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
							<%else%>
									 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(3)" value="&nbsp;&nbsp;Eliminar&nbsp;&nbsp;" class=label></TD>
							<%end if%>
					</TR>
			</TABLE>
		<TD>
		</TR>
	</TABLE>
	</FORM>
</BODY>
</HTML>
<script>
selecciona('forma.CarrierID','<%=CarrierID%>');
selecciona('forma.AirportID','<%=AirportID%>');
</script>
</script>