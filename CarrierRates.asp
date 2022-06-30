<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then

Checking "0|1"

Dim Conn, rs, CountList1Values, CountList2Values, CountList3Values, CountList4Values, CountList5Values, CountList6Values, CountTableValues
Dim aList1Values, aList2Values, aList3Values, aList4Values, aList5Values, aList6Values
Dim CarrierID, RangeID, AirportDesID, AirportDepID, i, Action, QuerySelect
CountList1Values = -1
CountList2Values = -1
CountList3Values = -1
CountList4Values = -1
CountList5Values = -1
CountList6Values = -1

CarrierID = CheckNum(Request.Form("CarrierID"))
AirportDesID = CheckNum(Request.Form("AirportDesID"))
AirportDepID = CheckNum(Request.Form("AirportDepID"))
Action = CheckNum(Request("Action"))

OpenConn Conn

	'Obteniendo listado de Carriers
	Set rs = Conn.Execute("select CarrierID, Name, Countries from Carriers where Expired = 0 and Countries in " & Session("Countries") & " order by Name, Countries")
	If Not rs.EOF Then
   		aList1Values = rs.GetRows
       	CountList1Values = rs.RecordCount-1
    End If
	CloseOBJ rs

	if CarrierID <> 0 then
		'Obteniendo listado de Aeropuertos Salida asignados al Carrier
        QuerySelect = "select b.AirportID, b.Name, b.AirportCode from CarrierDepartures a, Airports b where a.AirportID = b.AirportID and b.Expired=0 and CarrierID =" & CarrierID & " order by b.Name"
		'response.write QuerySelect & "<br>"
        Set rs = Conn.Execute(QuerySelect)
		If Not rs.EOF Then
    		aList2Values = rs.GetRows
        	CountList2Values = rs.RecordCount-1
	    End If
		CloseOBJ rs
		
		'Obteniendo listado de Aeropuertos Destino
		QuerySelect = "select AirportID, Name, AirportCode from Airports where Expired=0 order by Name"
        'response.write QuerySelect & "<br>"
        Set rs = Conn.Execute(QuerySelect)		        
		If Not rs.EOF Then
    		aList3Values = rs.GetRows
        	CountList3Values = rs.RecordCount-1
	    End If
		CloseOBJ rs
		
		'Obteniendo listado de Rangos
		Set rs = Conn.Execute("select b.Val, b.RangeID from CarrierRanges a, Ranges b where a.RangeID = b.RangeID and b.Expired = 0 and CarrierID =" & CarrierID & " order by b.Val")
		If Not rs.EOF Then
    		aList4Values = rs.GetRows
        	CountList4Values = rs.RecordCount-1
	    End If
		CloseOBJ rs
		'Obteniendo tarifas
		if CarrierID <> 0  and AirportDepID <> 0 and AirportDesID <> 0 then
			Select Case Action
        	Case 1 ' Insert
				Conn.Execute("delete from CarrierRates where CarrierID=" & CarrierID & " and AirportDepID=" & AirportDepID & " and AirportDesID=" & AirportDesID)
				Conn.Execute("insert into CarrierRates (CarrierID, AirportDepID, AirportDesID, RangeID, Val) Values (" & CarrierID & ", " & AirportDepID & ", " & AirportDesID & ", 0, " & CheckNum(request.Form("Range0")) & ")")
				For i = 0 To CountList4Values
					Conn.Execute("insert into CarrierRates (CarrierID, AirportDepID, AirportDesID, RangeID, Val) Values (" & CarrierID & ", " & AirportDepID & ", " & AirportDesID & ", " & aList4Values(1,i) & ", " & CheckNum(request.Form("Range" & aList4Values(1,i))) & ")") 
				Next
			Case 3 'Delete
				Conn.Execute("delete from CarrierRates where CarrierID=" & CarrierID & " and AirportDepID=" & AirportDepID & " and AirportDesID=" & AirportDesID)
			End Select

			'Obteniendo las tarifas por Rangos
			Set rs = Conn.Execute("select a.Val from CarrierRates a, Ranges b where a.CarrierID = " & CarrierID & " and a.AirportDepID = " & AirportDepID & " and a.AirportDesID = " & AirportDesID & " and a.RangeID = b.RangeID and b.Expired=0 order by b.Val")
			If Not rs.EOF Then
    			aList5Values = rs.GetRows
 	    	   	CountList5Values = rs.RecordCount-1
		    End If
			CloseOBJ rs
			'response.Write "select a.Val from CarrierRates a, Ranges b where a.CarrierID = " & CarrierID & " and a.AirportDepID = " & AirportDepID & " and a.AirportDesID = " & AirportDesID & " and a.RangeID = b.RangeID and b.Expired=0 order by b.Val"
			'Obteniendo la tarifa Minima
			Set rs = Conn.Execute("select Val from CarrierRates where CarrierID = " & CarrierID & " and AirportDepID = " & AirportDepID & " and AirportDesID = " & AirportDesID & " and RangeID = 0")
			If Not rs.EOF Then
    			aList6Values = rs.GetRows
 	    	   	CountList6Values = rs.RecordCount
		    End If
			CloseOBJ rs
			
			
		End If
	End If
CloseOBJ Conn
%>
<HTML>
<HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function validar(Action) {
	if (!valSelec(document.forma.CarrierID)){return (false)};
	<%
	if CarrierID <> 0 then
	%>
		if (!valSelec(document.forma.AirportDepID)){return (false)};
		if (!valSelec(document.forma.AirportDesID)){return (false)};
		<%
		if CarrierID <> 0  and AirportDepID <> 0 and AirportDesID <> 0 then
		%>
			if (!valTxt(document.forma.Range0, 1, 5)){return (false)};
			<%
			For i = 0 To CountList4Values
			%>
			if (!valTxt(document.forma.Range<%=aList4Values(1,i)%>, 1, 5)){return (false)};
			<%
			Next
		end if
	end if
	%>
	document.forma.Action.value = Action;
	document.forma.submit();			 
}
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="CarrierRates.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
		<TR><TD class=label align=right><b>Transportista:</b></TD><TD class=label align=left>
		<select name="CarrierID" id="Transportista" class=label onChange="<%if CarrierID <> 0 then%> document.forma.AirportDepID.value=0;document.forma.AirportDesID.value=0;<%end if%>document.forma.submit();">
		<option value="-1">Seleccionar</option>
		<%
			For i = 0 To CountList1Values
		%>
		<option value="<%=aList1Values(0,i)%>"><%=aList1Values(1,i) & " - " & aList1Values(2,i) & " - " & aList1Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
		</TD></TR> 
		<%
		if CarrierID <> 0 then
		%>
		<TR><TD class=label align=right><b>Aeropuerto Salida:</b></TD><TD class=label align=left>
		<select name="AirportDepID" id="Aeropuerto Salida" class=label>
		<option value="-1">Seleccionar</option>
		<%
			For i = 0 To CountList2Values
		%>
		<option value="<%=aList2Values(0,i)%>"><%=aList2Values(1,i) & " - " & aList2Values(2,i) & " - " & aList2Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
		</TD></TR> 
		<TR><TD class=label align=right><b>Aeropuerto Destino:</b></TD><TD class=label align=left>
		<select name="AirportDesID" id="Aeropuerto Destino" class=label onChange="document.forma.submit();">
		<option value="-1">Seleccionar</option>
		<%
			For i = 0 To CountList3Values
		%>
		<option value="<%=aList3Values(0,i)%>"><%=aList3Values(1,i) & " - " & aList3Values(2,i) & " - " & aList3Values(0,i)%></option>
		<%
    		Next
		%>
		</select>
		</TD></TR>
    	<%
		End If
		if CarrierID <> 0  and AirportDepID <> 0 and AirportDesID <> 0 then
		%>		
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
				<TD class=label align=center><b>MINIMO</b></TD>
			<%
			For i = 0 To CountList4Values
			%>
				<TD class=label align=center><b><%=aList4Values(0,i)%></b></TD>
			<%
    		Next
			%>
			</TR>
			<TR>
				<TD class=label align=right><input type="text" size="10" name="Range0" id="Rango Minimo" class=label value="<%if CountList6Values <> -1 then response.Write aList6Values(0,0) else response.write ""%>" onKeyUp="res(this,numb);"></TD>
			<%
			For i = 0 To CountList4Values
			%>
				<TD class=label align=right><input type="text" size="10" name="Range<%=aList4Values(1,i)%>" id="Rango <%=aList4Values(0,i)%>" class=label value="<%if CountList5Values <> -1 and CountList5Values>=CountList4Values then response.Write aList5Values(0,i) else response.write ""%>" onKeyUp="res(this,numb);"></TD>
			<%
    		Next
			'response.Write "Count4 = " & CountList4Values & "<br>"
			'response.Write "Count5 = " & CountList5Values & "<br>"
			'response.Write "Q4 = " & "select b.Val, b.RangeID from CarrierRanges a, Ranges b where a.RangeID = b.RangeID and b.Expired = 0 and CarrierID =" & CarrierID & " order by b.Val" & "<br>"
			'response.Write "Q5 = " & "select a.Val from CarrierRates a, Ranges b where a.CarrierID = " & CarrierID & " and a.AirportDepID = " & AirportDepID & " and a.AirportDesID = " & AirportDesID & " and a.RangeID = b.RangeID and b.Expired=0 order by b.Val" & "<br>"
			%>
			</TR>
			</TABLE>
		<TD>
		</TR>
		<%
   		End If
		%>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
			<%if CountList5Values = -1 then%>
				<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Ingresar&nbsp;&nbsp;" class=label></TD>
				<%else%>
				<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
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
selecciona('forma.AirportDepID','<%=AirportDepID%>');
selecciona('forma.AirportDesID','<%=AirportDesID%>');
</script>
<%Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
|