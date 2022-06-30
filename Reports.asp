<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim Action, ObjectID, QuerySelect, QuerySelect2, Conn, rs, aTableValues, TotPcs, TotWeight
Dim CountTableValues, aTableValues2, CountTableValues2, i, AwbType, TableName, iCountries, iDocID

	Action = CheckNum(Request("Action"))
	ObjectID = CheckNum(Request("OID"))
	AwbType = CheckNum(Request("AT"))
	CountTableValues = -1
	CountTableValues2 = -1
	
	if AwbType = 1 then
		TableName = "Awb"
	else
		TableName = "Awbi"
	end if

	Select case Action
	case 1
		QuerySelect = "select a.AgentData, a.ReservationDate, a.DeliveryDate, a.DepartureDate, a.AWBNumber, b.AirportCode, b.Name, c.AirportCode, c.Name, " & _
					  "a.ShipperData, a.ConsignerData, a.ChargeType, a.NoOfPieces, a.Weights, a.WeightsSymbol, a.NatureQtyGoods, a.Comment, a.Countries " & _
					  "from " & TableName & " a, Airports b, Airports c where a.AirportDepID=b.AirportID and a.AirportDesID=c.AirportID and a.AWBID = " & ObjectID
        iCountries = 17
        iDocID = "1" 'CONFIRMACION RESERVA
	case 2
		'QuerySelect = "select a.AgentData, b.Name, a.AWBDate, a.AWBNumber, a.FlightDate1, a.FlightDate2, c.AirportCode, c.Name, d.AirportCode, d.Name, a.ShipperData, a.ConsignerData, a.Comment2, a.Countries " & _
		'			  "from " & TableName & " a, Carriers b, Airports c, Airports d where a.CarrierID=b.CarrierID and a.AirportDepID=c.AirportID and a.AirportDesID=d.AirportID and a.AWBID = " & ObjectID
		QuerySelect = "select a.AgentData, b.Name, a.AWBDate, a.AWBNumber, a.FlightDate1, a.FlightDate2, c.AirportCode, c.Name, d.AirportCode, d.Name, a.ShipperData, a.ConsignerData, a.Comment2, a.Countries " & _
					  "from " & TableName & " a, Carriers b, Airports c, Airports d where a.CarrierID=b.CarrierID and a.AirportDepID=c.AirportID and a.AirportDesID=d.AirportID and a.HAWBNumber='' " & _
					  "and a.AWBNumber=(select AWBNumber from " & TableName & " where AWBID=" & ObjectID & ")"
        iCountries = 13
        iDocID = "2" 'HOUSE OF CARGO
		QuerySelect2 = "select a.HAWBNumber, a.ShipperData, a.ConsignerData, a.NoOfPieces, a.NatureQtyGoods, b.Name, a.Weights, a.WeightsSymbol " & _
					  "from " & TableName & " a, Airports b where a.HAWBNumber <> '' and a.AirportDesID=b.AirportID and a.AWBNumber = (select AWBNumber from " & TableName & " where AWBID =" & ObjectID & ")"
	case 3
		QuerySelect = "select a.AgentData, a.ArrivalAttn, a.ConsignerData, a.ShipperData, a.ArrivalFlight, c.AirportCode, c.Name, d.AirportCode, d.Name, " & _
					  "HDepartureDate, ArrivalDate, Cont, Destinity, TotalToPay, FiscalFactory, Concept, " & _		
					  "a.HAWBNumber, a.NoOfPieces, a.NatureQtyGoods, a.Weights, a.Comment3, a.AWBNumber, a.ManifestNumber, a.Countries, a.ConsignerID " & _
					  "from " & TableName & " a, Carriers b, Airports c, Airports d where a.AirportDepID=c.AirportID and a.AirportDesID=d.AirportID and a.AWBID = " & ObjectID
        iCountries = 23
        iDocID = "3" 'ARRIBO SALIDA

	case 4
		QuerySelect = "select trim(a.AWBNumber), trim(a.HAWBNumber), trim(c.AirportCode), trim(d.AirportCode), trim(a.TotNoOfPieces), trim(b.CarrierCode), trim(b.Name), a.Countries " & _
					  "from " & TableName & " a, Carriers b, Airports c, Airports d where a.CarrierID=b.CarrierID and a.AirportDepID=c.AirportID and a.AirportDesID=d.AirportID " & _
					  "and a.AwbID=" & ObjectID
        iCountries = 7

	end select
	'response.Write QuerySelect & "<br>"    
	'response.Write QuerySelect2 & "<br>"
	OpenConn Conn
	Set rs = Conn.Execute(QuerySelect)
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount
	End If  
	closeOBJ rs


    if CountTableValues = -1 then
        response.write "House Cargo solo realiza display de la Master."
        response.end
    end if


	if Action = 2 then
		Set rs = Conn.Execute(QuerySelect2)
		If Not rs.EOF Then
    		aTableValues2 = rs.GetRows
    		CountTableValues2 = rs.RecordCount-1
		End If
		closeOBJ rs
	end if
   	closeOBJ Conn

    'response.write "(" & aTableValues(iCountries, 0) & ")(" & iDocID & ")"    

    Dim iResult, iEdicion, iTitulo, iEmpresa, iDireccion, iLogo, iObservaciones

    iResult = WsGetLogo(aTableValues(iCountries, 0), "AEREO",  iDocID,  "",  "")

    iLogo = iResult(20)
    iEdicion = iResult(2)
    iTitulo = iResult(3)
    iEmpresa = iResult(4)
    iDireccion = iResult(6)
    iObservaciones = iResult(1)

    'response.write "(" & iLogo & ")(" & iEmpresa & ")"
		
	if CountTableValues <> -1 then
%>
<HTML>
<HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE cellpadding="4" cellspacing="4">
<% select case Action
case 1%>
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=right colspan="2"><%=IIf(iLogo = "", DisplayLogo(aTableValues(17, 0)), iLogo)%><br><br></TD></TR> 
	</TABLE>
</TD></TR>
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=center colspan="2"><%=FRegExp(chr(13),aTableValues(0, 0),"<br>",4)%><br><br></TD></TR> 
	</TABLE><br><br>
</TD></TR>
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=center colspan="2"><b><%=IIf(iTitulo = "", "Confirmaci&oacute;n de Reservaci&oacute;n a L&iacute;nea A&eacute;rea", iTitulo)%>:</b></TD></TR> 
		<TR><TD class=label align=left colspan="2">Agradecemos su atenci&oacute;n a realizar la siguiente reservaci&oacute;n:</TD></TR>
	</TABLE><br><br>
</TD></TR><TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left border="1">
		<TR><TD class=label align=right><b>Fecha de Reserva:</b></TD><TD class=label align=left colspan=2><%=aTableValues(1, 0)%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha&nbsp;Entrega&nbsp;Linea&nbsp;Aerea:</b></TD><TD class=label align=left colspan=2><%=aTableValues(2, 0)%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha de Salida:</b></TD><TD class=label align=left colspan=2><%=aTableValues(3, 0)%></TD></TR> 
		<TR><TD class=label align=right><b>AWB No.:</b></TD><TD class=label align=left colspan=2><b><%=aTableValues(4, 0)%></b></TD></TR> 
		<TR><TD class=label align=right><b>Origen:</b></TD><TD class=label align=left colspan=2><%=aTableValues(5, 0) & " - " & aTableValues(6, 0)%></TD></TR> 
		<TR><TD class=label align=right><b>Destino:</b></TD><TD class=label align=left colspan=2><b><%=aTableValues(7, 0) & " - " & aTableValues(8, 0)%></b></TD></TR> 
		<TR><TD class=label align=right><b>Shipper:</b></TD><TD class=label align=left colspan=2><%=FRegExp(chr(13),aTableValues(9, 0),"<br>",4)%></TD></TR> 
		<TR><TD class=label align=right><b>Consignee:</b></TD><TD class=label align=left colspan=2><%=FRegExp(chr(13),aTableValues(10, 0),"<br>",4)%></TD></TR> 
		<TR><TD class=label align=right><b>Termino:</b></TD><TD class=label align=left colspan=2><%if aTableValues(11, 0) = 1 then%>PP<%else%>CC<%end if%></TD></TR> 
	</TABLE>
</TD></TR><TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left border="1">
		<TR>
		<TD class=label><b>Piezas:</b></TD>
		<TD class=label><b>Peso(kgs):</b></TD>
		<TD class=label><b>Dimensiones:</b></TD>
		<TD class=label><b>Commodity:</b></TD>
		</TR> 
		<TR>
		<TD class=label><%=FRegExp(chr(13),aTableValues(12, 0),"<br>",4)%></TD>
		<TD class=label><%=FRegExp(chr(13),aTableValues(13, 0),"<br>",4)%></TD>
		<TD class=label><%=FRegExp(chr(13),aTableValues(14, 0),"<br>",4)%></TD>
		<TD class=label><%=FRegExp(chr(13),aTableValues(15, 0),"<br>",4)%></TD>
		</TR>
	</TABLE>
</TD></TR><TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=left><%=FRegExp(chr(13),aTableValues(16, 0),"<br>",4)%><br><br></TD></TR> 
	</TABLE>
</TD></TR>	
<% case 2%>
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label2 align=right colspan="2"><%=IIf(iLogo = "", DisplayLogo(aTableValues(13, 0)), iLogo)%><br><br></TD></TR> 
	</TABLE>
</TD></TR>
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label2 align=center colspan="2"><%=FRegExp(chr(13),aTableValues(0, 0),"<br>",4)%><br></TD></TR> 
	</TABLE><br>
</TD></TR>
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label2 align=center colspan="2"><b><%=IIf(iTitulo = "", "HOUSE CARGO MANIFEST", iTitulo)%></b></TD></TR> 
	</TABLE><br>
</TD></TR><TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left border="1">
		<TR><TD class=label2 align=right><b>AIRLINE:</b></TD><TD class=label2 align=left colspan=2><%=aTableValues(1, 0)%></TD></TR> 
		<TR><TD class=label2 align=right><b>DATE:</b></TD><TD class=label2 align=left colspan=2><%=aTableValues(2, 0)%></TD></TR> 
		<TR><TD class=label2 align=right><b>MASTER AWB:</b></TD><TD class=label2 align=left colspan=2><%=aTableValues(3, 0)%></TD></TR> 
		<TR><TD class=label2 align=right><b>FLIGTH:</b></TD><TD class=label2 align=left colspan=2><b><%=aTableValues(4, 0) & " - " & aTableValues(5, 0)%></b></TD></TR> 
		<TR><TD class=label2 align=right><b>PORT OF LOADING:</b></TD><TD class=label2 align=left colspan=2><%=aTableValues(6, 0) & " - " & aTableValues(7, 0)%></TD></TR> 
		<TR><TD class=label2 align=right><b>PORT OF DISCHARGE:</b></TD><TD class=label2 align=left colspan=2><b><%=aTableValues(8, 0) & " - " & aTableValues(9, 0)%></b></TD></TR> 
		<TR><TD class=label2 align=right><b>SHIPPER:</b></TD><TD class=label2 align=left colspan=2><%=FRegExp(chr(13),aTableValues(10, 0),"<br>",4)%></TD></TR> 
		<TR><TD class=label2 align=right><b>CONSIGNEE:</b></TD><TD class=label2 align=left colspan=2><%=FRegExp(chr(13),aTableValues(11, 0),"<br>",4)%></TD></TR> 
	</TABLE>
</TD></TR><TR><TD>
		<TABLE cellspacing=0 cellpadding=2 width=600 align=left border="1">
		<%
		if CountTableValues2 >= 0 then
	 	for i = 0 to CountTableValues2
		%>
		<TR>
		<TD class=label2><b>HAWB:</b></TD>
		<TD class=label2 colspan="4"><b><%=aTableValues2(0, i)%></b></TD>
		</TR>
		<TR>
		<TD class=label2><b>SHIPPER:</b></TD>
		<TD class=label2 colspan="4"><%=FRegExp(chr(13),aTableValues2(1, i),"<br>",4)%></TD>
		</TR>
		<TR>
		<TD class=label2><b>CONSIGNEE:</b></TD>
		<TD class=label2 colspan="4"><%=FRegExp(chr(13),aTableValues2(2, i),"<br>",4)%></TD>
		</TR>
		<TR>
		<TD class=label2><b>PCS:</b></TD>
		<TD class=label2><b>DESCRIPTION:</b></TD>
		<TD class=label2><b>FINAL DESTINATION:</b></TD>
		<TD class=label2 colspan="2"><b>GROSS WEIGHT:</b></TD>
		</TR> 
		<TR>
		<TD class=label2><%=FRegExp(chr(13),aTableValues2(3, i),"<br>",4)%></TD>
		<TD class=label2><%=FRegExp(chr(13),aTableValues2(4, i),"<br>",4)%></TD>
		<TD class=label2><%=FRegExp(chr(13),aTableValues2(5, i),"<br>",4)%></TD>
		<TD class=label2><%=FRegExp(chr(13),aTableValues2(6, i),"<br>",4)%></TD>
		<TD class=label2><%=FRegExp(chr(13),aTableValues2(7, i),"<br>",4)%></TD>
		</TR>
		<TR><TD class=label2 colspan="5">&nbsp;</TD></TR>
		<%	TotPcs = TotPcs + Round(aTableValues2(3, i),2)
		   	TotWeight = TotWeight + Round(aTableValues2(6, i),2)
			next
		%>
		<TR><TD class=label2><b>TOTAL PCS</b></TD><TD class=label2 colspan="4"><%=TotPcs%></TD></TR>
		<TR><TD class=label2><b>TOTAL WEIGHT</b></TD><TD class=label2 colspan="4"><%=TotWeight & " " & aTableValues2(7, 0) %></TD></TR>
		<%end if%>
		</TABLE>
</TD></TR><TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label2 align=left><%=FRegExp(chr(13),aTableValues(12, 0),"<br>",4)%><br><br></TD></TR> 
	</TABLE>
</TD></TR>	
<% case 3%>
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=right colspan="2"><%=IIf(iLogo = "", DisplayLogo(aTableValues(23, 0)), iLogo)%><br></TD></TR> 
	</TABLE>
</TD></TR>
<!--<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=center colspan="2"><%'=FRegExp(chr(13),aTableValues(0, 0),"<br>",4)%><br></TD></TR> 
	</TABLE><br>
</TD></TR>-->
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=center colspan="2">
		<b>        
        <% 'IIf(iTitulo = "", "HOUSE CARGO MANIFEST", iTitulo)%>        
        AVISO DE <%if AwbType=1 then%> SALIDA <%else%> LLEGADA / <%if aTableValues(7, 0)="MGA" then%>TERMINAL AEREA MANAGUA<%else%>ARRIBO<%end if end if%>        
        </b>
		</TD></TR>
		<TR><TD class=label align=center colspan="2"><br>Por medio del presente se les informa <%if AwbType=1 then%>la salida <%else%>el arribo<%end if%> de la siguiente mercader&iacute;a:</TD></TR> 
	</TABLE><br>
</TD></TR><TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left border="1">
		<%if aTableValues(1, 0) <> "" then%>
		<TR><TD class=label align=right><b>ATENCION:</b></TD><TD class=label align=left colspan=2><%=aTableValues(1, 0)%></TD></TR> 
		<%end if%>
		<%if aTableValues(2, 0) <> "" then%>
		<TR><TD class=label align=right><b>CONSIGNEE:</b></TD><TD class=label align=left colspan=2><%=FRegExp(chr(13),aTableValues(2, 0),"<br>",4)%></TD></TR>
		<%end if%>
		<%if aTableValues(3, 0) <> "" then%>
		<TR><TD class=label align=right><b>SHIPPER:</b></TD><TD class=label align=left colspan=2><%=FRegExp(chr(13),aTableValues(3, 0),"<br>",4)%></TD></TR> 
		<%end if%>
		<%if aTableValues(4, 0) <> "" then%>
		<TR><TD class=label align=right><b>FLIGTH:</b></TD><TD class=label align=left colspan=2><b>&nbsp;<%=aTableValues(4, 0)%></b></TD></TR>
		<%end if%>
		<%if aTableValues(5, 0) <> "" or aTableValues(6, 0) <> "" then%>
		<TR><TD class=label align=right><b>AIRPORT OF LOADING:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(5, 0) & " - " & aTableValues(6, 0)%></TD></TR>
		<%end if%>
		<%if aTableValues(7, 0) <> "" or aTableValues(8, 0) <> "" then%>
		<TR><TD class=label align=right><b>PORT OF DISCHARGE:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(7, 0) & " - " & aTableValues(8, 0)%></TD></TR> 
		<%end if%>
		<%if aTableValues(9, 0) <> "" then%>
		<TR><TD class=label align=right><b>DEPARTURE DATE:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(9, 0)%></TD></TR> 
		<%end if%>
		<%if aTableValues(10, 0) <> "" then%>
		<TR><TD class=label align=right><b>ARRIVAL DATE:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(10, 0)%></TD></TR> 
		<%end if%>
		<%if aTableValues(11, 0) <> "" then%>
		<TR><TD class=label align=right><b>CONTENT:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(11, 0)%></TD></TR> 
		<%end if%>
		<%if aTableValues(12, 0) <> "" then%>
		<TR><TD class=label align=right><b>DESTINITY:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(12, 0)%></TD></TR> 
		<%end if%>
		<%if aTableValues(13, 0) <> "" then%>
		<TR><TD class=label align=right><b>TOTAL TO PAY:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(13, 0)%></TD></TR> 
		<%end if%>
		<%if aTableValues(14, 0) <> "" then%>
		<TR><TD class=label align=right><b>FISCAL FACTORY:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(14, 0)%></TD></TR> 
		<%end if%>
		<%if aTableValues(15, 0) <> "" then%>
		<TR><TD class=label align=right><b>CONCEPT:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(15, 0)%></TD></TR> 
		<%end if%>
		<%if aTableValues(22, 0) <> "" then%>
		<TR><TD class=label align=right><b>MANIFEST:</b></TD><TD class=label align=left colspan=2>&nbsp;<%=aTableValues(22, 0)%></TD></TR> 
		<%end if%>
	</TABLE>
</TD></TR>
<TR><TD>
		<TABLE cellspacing=0 cellpadding=2 width=600 align=left border="1">
		<TR>
		<TD class=label><b><%if aTableValues(16, 0)<>"" then%>HAWB<%else%>AWB<%end if%>:</b></TD>
		<TD class=label><b>PCS:</b></TD>
		<TD class=label><b>DESCRIPTION:</b></TD>
		<TD class=label><b>GROSS WEIGHT:</b></TD>
		</TR> 
		<TR>
		<TD class=label><% if aTableValues(16, 0)<>"" then response.write aTableValues(16, 0) else response.write aTableValues(21, 0) end if%></TD>
		<TD class=label><%=FRegExp(chr(13),aTableValues(17, 0),"<br>",4)%></TD>
		<TD class=label><%=FRegExp(chr(13),aTableValues(18, 0),"<br>",4)%></TD>
		<TD class=label><%=FRegExp(chr(13),aTableValues(19, 0),"<br>",4)%></TD>
		</TR>
		</TABLE>
</TD></TR>
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=left><%=FRegExp(chr(13),aTableValues(20, 0),"<br>",4)%><br><br></TD></TR> 
	</TABLE>
</TD></TR>
<TR><TD>
	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=left>
        <%=CheckCreditClient(aTableValues(24,0),SetCountryBAW(aTableValues(23,0)))%><br><br></TD></TR> 
	</TABLE>
</TD></TR>
<%if aTableValues(23,0)="HN" then%>
<TR><TD>
    <TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=justify>
<% if iObservaciones = "" then %>
        Se aceptan las siguientes formas de pago:
    <br><br>
    <b>-Pago por medio de deposito en efectivo.</b>
    <br><br>
    <b>-Pago por medio de deposito de cheque:</b>
    Se espera hasta que los fondos sean reflejados en nuestra cuenta para poder libera documentos o carga.  Si el cheque es devuelto por el banco por falta de fondos, se le cobrara el recargo del banco mas un recargo nuestro de $50.00 si el cheque es en dolares y Lps500.00 si el queche es en Lempiras.
    <br><br>
    <b>-Pago por medio de Transferencia bancaria:</b>
    Se espera hasta que los fondos sean reflejados en nuestra cuenta para poder libera documentos y carga.
    <br><br>
    <b>NO se aceptan Pago en efectivo Ni en Dolares , ni en Lempiras</b>
    <br><br>
    <b>- Pago de facturas en Dolares al cambio en Lempiras</b>
    <br>
    Se devolvera lo depositado en Lempiras restando Lps 600.00 por gastos administrativos y se exigira el pago correspondiente en Dolares para poder liberar documentos o la carga.

<% else 
    response.write "**" &   iObservaciones
end if %>
        </TD></TR> 
	</TABLE>
</TD></TR>
<%end if%>

<%

'response.write "(" & aTableValues(23,0) & ")"

if aTableValues(23,0)="GT" then 'agregado 2017-07-05 segun ticket [Ticket#2017062604000483] NUEVA FORMA DE PAGO POR CONVENIO DE PAGO AIMAR%>
<TR><TD>
    <TABLE cellspacing=0 cellpadding=2 width=610 align=left>
		<TR><TD class=label align=justify>



<style>
	.styleborder { border:1px solid gray }
	.style4 { border:1px solid gray }
	.style10, td { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 7pt; }
</style>

<% if iObservaciones = "" then %>
<table>
<tr>
	    <td class="style4" align="left" width="100%">
        <b>OBSERVACIONES IMPORTANTES</b>
        <table cellpadding="2" cellspacing="0" class="styleborder">
            <tr>
            <td class="style4" align=justify>
            <span class=style10>
            Para realizar sus pagos en quetzales, puede realizarlos a través de depósitos monetarios y  a continuación detallamos los números de convenio:
            <BR><BR>
            <table cellpadding="2" cellspacing="0" class="styleborder">
                <tr>
                <td class="style4" align="left">
                <span class=style10>Banco G&T Continental (Quetzales)</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 012-0001068-6 </span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Aimar, S.A.</span>
                </td>
                <td class="style4" align="left">
                <span class=style10><b>No de convenio: 8292</b></span>
                </td>
                </tr>
                <tr>      
                <td class="style4" align="left">
                <span class=style10>Banco Industrial (Quetzales)</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 027-018962-1</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Aimar, S.A.</span>
                </td>
                <td class="style4" align="left">
                <span class=style10><b>No de convenio: 2253</b></span>
                </td>
                </tr>
                <tr>      
                <td class="style4" align="left">
                <span class=style10>Banrural (Quetzales)</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 30-3343978-4 </span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Aimar, S.A.</span>
                </td>
                <td class="style4" align="left">
                <span class=style10><b>No de convenio: ATX-249-426-1</b></span>
                </td>
                </tr>
            </table>
            <BR>Se reciben cheques de empresa local o personales, cheques de caja <b>emitidos a nombre de Aimar, S.A. o Agencia Internacional Marítima, S.A.</b>
Nota: La reincidencia de cheques rechazados conlleva a que se acepten pagos únicamente por medio de cheques de caja y tiene un costo de $35.00 y en quetzales Q168.00
        <BR><BR>
        No se recibirá: efectivo o giros bancarios
        <BR><BR>
        Favor utilizar la boleta de depósito que se adjunta en la factura electrónica en la parte inferior del documento.
        <BR><BR>Enviar copia de la boleta escaneada al correo: <b>creditosycobros-gt@aimargroup.com</b> con las instrucciones de aplicación de pago (detalle de facturas o notas de débito)
 
        <BR><BR>Para realizar sus pagos en dólares puede realizar a través de depósitos monetarios a las cuentas en dólares detalladas a continuación:

        <BR><BR>

        <table cellpadding="2" cellspacing="0" class="styleborder">
            <tr>
            <td class="style4" align="left">
            <span class=style10>Banco G&T Continental (Dólares)</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>No. 7858059517</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>Aimar, S.A.</span>
            </td>
            </tr>
            <tr>      
            <td class="style4" align="left">
            <span class=style10>Banco Industrial, S.A. (Dólares)</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>No. 027-003599-1 </span>
            </td>
            <td class="style4" align="left">
            <span class=style10>Aimar, S.A.</span>
            </td>
            </tr>
            <tr>      
            <td class="style4" align="left">
            <span class=style10>Banrural (Dólares)</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>No. 6445015801 </span>
            </td>
            <td class="style4" align="left">
            <span class=style10>Aimar, S.A.</span>
            </td>
            </tr>
        </table>

        <BR><BR>
        Enviar copia de la boleta escaneada al correo: <b>creditosycobros-gt@aimargroup.com</b> con las instrucciones de aplicación de pago (detalle de facturas o notas de débito)

        <BR><BR>O bien puede realizar transferencia bancaria a los datos detallados a continuación:
 
        <BR><BR>Transferencias Internacionales a Banco G&T Continental, S.A.:
        <BR>Banco Intermediario: BANK OF AMERICA, N.A., NEW YORK USA
        <BR><b>ABA:</b> 026009593
        <BR><b>SWIFT:</b> BOFAUS3N
        <BR><b>CUENTA:</b> 1901734945 de Banco G&T Continental, S.A., Guatemala
        <BR><b>SWIFT:</b> GTCOGTGC
        <BR>Para finalmente acreditar a:
        <BR><b>Nombre del Beneficiario :</b> Aimar, S.A.
        <BR>Cuenta: 7858059517
        <BR><BR>
    POR INSTRUCCIONES DE NUESTRO AGENTE SE SOLICITA UN BL ORIGINAL PARA REGOGER DOCUMENTOS. Favor de hacer este
    pago para poder entregar Copia de poliza de traslado y su respectivo endoso. En caso de RECLAMO debe hacerse por escrito dentro de
    los primeros 10 DIAS DEL CALENDARIO (Contando Sabado y Domingo) mismos que seran contados apartir de la FECHA DE
    DESCARGA ARRIBA DESCRITA de lo contrario EL RECLAMO NO SERA TOMADO EN CUENTA NI SE LE DARA TRAMITE ALGUNO.
    <BR><BR>
    TOMAR EN CUENTA QUE LA FACTURA SE REALIZO SEGÚN LA INFORMACION COLOCADA EN EL REQUERIMIENTO DE PARTIDAS
    O INFORMACION ANTICIPADA, POR CAMBIO DE LA MISMA TIENE UN RECARGO DE Q250.00 ó $30.00 SI SE SOLICITA CAMBIO DE
    FACTURA DE MESES ANTERIORES DEBERA CANCELAR EL VALOR DE IVA E ISR.</span>
            </td>
            </tr>
            </table>
        </td>
</tr>
</table>
<% else 
    response.write iObservaciones
 end if %>
        <!--
Para realizar sus pagos en quetzales, puede realizarlos a través de depósitos monetarios y  a continuación detallamos los números de convenio:
<br><br>
Convenio de Pago en los bancos del sistema:<br>
<ul>
<li>Banco G&T Continental, S.A. No de convenio: 8292</li>
<li>Banco Industrial, S.A. No de convenio: 2253</li>
<li>Banrural No de convenio: ATX-249-426-1</li>
</ul>
Favor utilizar la boleta de depósito que se adjunta en la factura electrónica en la parte inferior del documento.     Enviar copia de la boleta escaneada al correo: <a style="color:blue" href="mailto:Créditos y Cobros Aimar <creditosycobros-gt@aimargroup.com>">creditosycobros-gt@aimargroup.com</a> con las instrucciones de aplicación de pago (detalle de facturas o notas de débito)
<br><br> 
Para realizar su pagos en dólares puede realizar a través de depósitos monetarios a las cuentas en dólares detalladas en esta notificación.   Enviar copia de la boleta escaneada al correo: <a style="color:blue" href="mailto:Créditos y Cobros Aimar <creditosycobros-gt@aimargroup.com>">creditosycobros-gt@aimargroup.com</a> con las instrucciones de aplicación de pago (detalle de facturas o notas de débito).   O bien puede realizar transferencia bancaria a los datos detallados a continuación:
<br><br>
Transferencias Internacionales a Banco G&T Continental, S.A.:<br>
Banco Intermediario: BANK OF AMERICA, N.A., NEW YORK USA<br>
ABA: 026009593<br>
SWIFT: BOFAUS3N<br>
CUENTA: 1901734945 de Banco G&T Continental, S.A., Guatemala<br>
SWIFT: GTCOGTGC<br>
Para finalmente acreditar a:<br>
Nombre del Beneficiario, S.A. Agencia Internacional Maritima, S. A.<br>
Cuenta: 7858059517<br> -->



        </TD></TR> 
	</TABLE>
</TD></TR>
<%end if%>

<% case 4%>

<html moznomarginboxes mozdisallowselectionprint> 

<style>
/*
#tag {
    background-image:url('img/EtiquetaExport.png'); 
    background-repeat:no-repeat;    
    width:318px;
    height:543px;
    border:1px solid red;
}
*/
@page {         
    size: auto;   /* auto is the current printer page size */
    margin: 0mm;  /* this affects the margin in the printer settings */
}

.vertical-text {
	transform: rotate(-90deg);	
	font-family: Calibri;	
	font-weight:bolder;
	font-size:40px;	
}
.sml {   font-size:12px;    }
.med {   font-size:20px;    }
.lrg {   font-size:30px;    }

td { padding:0px; margin:0px;
    font-family: Calibri;	
	font-weight:bolder;
	font-size:26px;	    
	text-align:center; 
	border:1px solid black;
     }
     /*
.tit {
    position:relative; top:-2; left:-85px;display:block;
    }     
.dat {
    position:relative; top:-2; left:0px;display:block;
    }  */   
</style>


<!-- <table width="200px" style="position:absolute;top:24mm;left:-16mm;transform: rotate(-90deg);" border=1 cellpadding=0 cellspacing=0> -->
<table width="200px" style="position:absolute;;transform: rotate(-90deg);left:15px;" border=1 cellpadding=0 cellspacing=0>
<tr>
    <th colspan=3 class="sml">   
    <%=IIf(iLogo = "", "<img style='height:18mm;' src='img/EtiquetaHeadExport" & IIf(InStr(Session("Countries"),"GT"),"2","1") & ".png'>", Replace(iLogo,"<img ","<img style='width:100%;height:18mm;' "))%>    
    </th>    
</tr>                                                                                                   
<tr>
    <td colspan=2><span class="sml tit">AWB No.</span><%=aTableValues(0,0)%></td>
    <td class="med">PCS.</td>    
</tr>
<tr>
    <td colspan=2><span class="sml tit">HAWB No.</span><span class="dat"><%=aTableValues(1,0)%></span></td>
    <td><%=aTableValues(4,0)%></td>    
</tr>
<tr>
    <td width="33%" class="sml">DESTINATION</td>
    <td width="33%" class="sml">ORIGIN</td>
    <td width="33%" class="sml">VIA</td>
</tr>
<tr>
    <td class="med"><%=aTableValues(2,0)%></td>
    <td class="med"><%=aTableValues(3,0)%></td>
    <td class="med"><%=aTableValues(5,0)%></td>
</tr>
<tr>
    <td colspan=3 class="sml"><%=aTableValues(6,0)%></td>    
</tr>
</table>



<% end select%>
</TD></TR>
    
<% if Action <> 4 then %>

	<TABLE cellspacing=0 cellpadding=2 width=600 align=left>
		<TR><TD class=label align=left>Atentamente,</TD></TR>
		<TR><TD class=label align=left><%=FRegExp(chr(13),Session("Sign"),"<br>",4)%></TD></TR>
	</TABLE>

<% end if %>
</TABLE>	
</BODY>
</HTML>
<%
	end if	
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
