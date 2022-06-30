<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim ObjectID, Conn, rs, Action, aTableValues, CountTableValues, Currencies, GroupID
Dim Name, Volume, Weight, Agent, HBLNumber, BL, i, FisBillID, FinBillID, CantItems
Dim Freight, Freight2, Insurance, Insurance2, AnotherChargesCollect, AnotherChargesPrepaid

ObjectID = CheckNum(Request("OID"))
Action = CheckNum(Request("Action"))
CountTableValues = -1
CantItems = 15

OpenConn Conn
	if Action=2 then
		SaveChargeItems Conn, ObjectID, Action
	end if

	Set rs = Conn.Execute("select Clients, Volumes, Weights, Agents, HBLNumber, BLs, FisBillID, FinBillID from BLDetail where BLDetailID=" & ObjectID)
	if Not rs.EOF then
		Name = rs(0)
		Volume = rs(1)
		Weight = rs(2)
		Agent = rs(3)
		HBLNumber = rs(4)
		BL = rs(5)
		FisBillID = rs(6)
		FinBillID = rs(7)		
	end if
	CloseOBJ rs
	Set rs = Conn.Execute("select UserID, ItemName, ItemID, Currency, Value, OverSold, Local, PrepaidCollect, ServiceID, ServiceName, InvoiceID, CalcInBL from ChargeItems where Expired=0 and SBLID=" & ObjectID & " Order By InvoiceID Desc, PrepaidCollect, Local, Currency, ServiceName, ItemName")
	if Not rs.EOF then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	end if
CloseOBJs rs, Conn

OpenConn2 Conn
	'Obteniendo Monedas
	Set rs = Conn.Execute("select distinct simbolo from monedas order by simbolo")
	Do While Not rs.EOF
		Currencies = Currencies & "<option value=" & rs(0) & ">" & rs(0) & "</option>"
		rs.MoveNext
	Loop
CloseOBJs rs, Conn

%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
		var sep = '';
		CantItems=-1;
  		document.forma.ItemServIDs.value = "";
		document.forma.ItemServNames.value = "";
  		document.forma.ItemNames.value = "";
  		document.forma.ItemIDs.value = "";
  		document.forma.ItemCurrs.value = "";
  		document.forma.ItemVals.value = "";
  		document.forma.ItemOVals.value = "";
  		document.forma.ItemLocs.value = "";
  		document.forma.ItemPPCCs.value = "";
		document.forma.ItemInvoices.value = "";
		document.forma.ItemCalcInBLs.value = "";
		
		for (i=0; i<=<%=CantItems%>;i++) {
			if (document.forma.elements["N"+i].value != '') {
				if (!valSelec(document.forma.elements["N"+i])){return (false)};
				if (!valSelec(document.forma.elements["C"+i])){return (false)};
				if (!valTxt(document.forma.elements["V"+i], 1, 5)){return (false)};
				if (!valSelec(document.forma.elements["T"+i])){return (false)};
				if (!valSelec(document.forma.elements["TC"+i])){return (false)};
				if (!valSelec(document.forma.elements["CCBL"+i])){return (false)};
				if (document.forma.elements["OV"+i].value == '') {document.forma.elements["OV"+i].value = 0};
				if (document.forma.elements["SVI"+i].value!="") {
					document.forma.ItemServIDs.value = document.forma.ItemServIDs.value + sep + document.forma.elements["SVI"+i].value;
					document.forma.ItemServNames.value = document.forma.ItemServNames.value + sep + document.forma.elements["SVN"+i].value;
				} else {
					document.forma.ItemServIDs.value = "0" + sep + document.forma.elements["SVI"+i].value;
					document.forma.ItemServNames.value = " " + sep + document.forma.elements["SVN"+i].value;
				}
				document.forma.ItemNames.value = document.forma.ItemNames.value + sep + document.forma.elements["N"+i].value;
				document.forma.ItemIDs.value = document.forma.ItemIDs.value + sep + document.forma.elements["I"+i].value;
				document.forma.ItemCurrs.value = document.forma.ItemCurrs.value + sep + document.forma.elements["C"+i].value;
				document.forma.ItemVals.value = document.forma.ItemVals.value + sep + document.forma.elements["V"+i].value;
				document.forma.ItemOVals.value = document.forma.ItemOVals.value + sep + document.forma.elements["OV"+i].value;
				document.forma.ItemLocs.value = document.forma.ItemLocs.value + sep + document.forma.elements["T"+i].value;
				document.forma.ItemPPCCs.value = document.forma.ItemPPCCs.value + sep + document.forma.elements["TC"+i].value;
				document.forma.ItemInvoices.value = document.forma.ItemInvoices.value + sep + document.forma.elements["INV"+i].value;
				document.forma.ItemCalcInBLs.value = document.forma.ItemCalcInBLs.value + sep + document.forma.elements["CCBL"+i].value;
				CantItems++;
				sep = "|";
			}
		}
	    document.forma.CantItems.value = CantItems;
		document.forma.Action.value = Action;
		//alert(document.forma.ItemServIDs.value);
		//alert(document.forma.ItemServNames.value);
		//alert(document.forma.ItemNames.value);
		//alert(document.forma.ItemIDs.value);
		//alert(document.forma.ItemCurrs.value);
		//alert(document.forma.ItemVals.value);
		//alert(document.forma.ItemOVals.value);
		//alert(document.forma.ItemLocs.value);
		//alert(document.forma.ItemPPCCs.value);
		//alert(document.forma.CantItems.value);
		//alert(document.forma.ItemInvoices.value);
		document.forma.submit();			 
	 }
	 
	 function DelCharge(Pos) {
		document.forma.elements["SVI"+Pos].value='';
		document.forma.elements["SVN"+Pos].value='';
		document.forma.elements["N"+Pos].value='';
		document.forma.elements["I"+Pos].value='';
		document.forma.elements["C"+Pos].value='-1';
		document.forma.elements["V"+Pos].value='';
		document.forma.elements["OV"+Pos].value='';
		document.forma.elements["T"+Pos].value='-1';
		document.forma.elements["TC"+Pos].value='-1';
		document.forma.elements["INV"+Pos].value='0';
		document.forma.elements["CCBL"+Pos].value='-1';
		return false;	 
	 }
	 
	 function AddCharge(Pos) {
		window.open('Search_Charges.asp?GID=29&N='+Pos,'BLData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');
		return false;	 
	 }
	 
	 function ValidarDoble(Pos) {
	 	for (i=0; i<=<%=CantItems%>;i++) {
			if  (i!= Pos) {
				if ((document.forma.elements["SVI"+i].value==document.forma.elements["SVI"+Pos].value) && 
				(document.forma.elements["SVN"+i].value==document.forma.elements["SVN"+Pos].value) &&
				(document.forma.elements["N"+i].value==document.forma.elements["N"+Pos].value) &&
				(document.forma.elements["I"+i].value==document.forma.elements["I"+Pos].value) &&
				(document.forma.elements["INV"+i].value=='0')) {
					alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
					DelCharge(Pos);
					return (false);
				}
			}			
		}
	 }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<style type="text/css">
<!--
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.style4 {	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style8 {	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: bold;
	color: #999999;
}
-->
</style>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="self.focus();">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="ItineraryCharges.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="ItemServIDs" type=hidden value="">
	<INPUT name="ItemServNames" type=hidden value="">
	<INPUT name="ItemNames" type=hidden value="">
	<INPUT name="ItemIDs" type=hidden value="">
	<INPUT name="ItemCurrs" type=hidden value="">
	<INPUT name="ItemVals" type=hidden value="">
	<INPUT name="ItemOVals" type=hidden value="">
	<INPUT name="ItemLocs" type=hidden value="">
	<INPUT name="ItemPPCCs" type=hidden value="">
	<INPUT name="CantItems" type=hidden value="-1">
	<INPUT name="ItemInvoices" type=hidden value="-1">
	<INPUT name="ItemCalcInBLs" type=hidden value="-1">
		<TD colspan="2" class=label align=center>
		<table width="80%" border="0" align="center">
			<TR><TD class=label align=right><b>Carta Porte:</b></TD><TD class=label align=left><%=HBLNumber%></TD></TR> 
			<TR><TD class=label align=right><b>Consignatario:</b></TD><TD class=label align=left><%=Name%></TD></TR>
			<TR><TD class=label align=right><b>Volumen:</b></TD><TD class=label align=left><%=Volume%></TD></TR> 
			<TR><TD class=label align=right><b>Peso:</b></TD><TD class=label align=left><%=Weight%></TD></TR> 
			<TR><TD class=label align=right><b>Agente:</b></TD><TD class=label align=left><%=Agent%></TD></TR> 
			<TR><TD class=label align=right><b>BL o RO:</b></TD><TD class=label align=left><%=BL%></TD></TR> 
		</table>
		
		<table width="80%" border="0">
		  <tr><td class=submenu colspan=9></td></tr>
		  <tr>
			<td align="center" class="style4">
				Servicio
			</td>
			<td align="center" class="style4">
				Rubro
			</td>
			<td align="center" class="style4">&nbsp;
								
			</td>
			<td align="center" class="style4">
				Moneda
			</td>
			<td align="center" class="style4">
				Monto
			</td>
			<td align="center" class="style4">
				SobreVenta
			</td>
			<td align="center" class="style4">
				Tipo
			</td>
			<td align="center" class="style4">
				Pago
			</td>
			<td align="center" class="style4">
				Calcular en HBL?
			</td>
			<td align="center" class="style4">&nbsp;
								
			</td>
		  </tr>
		  <%for i=0 to CantItems%>
		  <tr>
			<td align="right" class="style4">
				<input type="text" size="30" class="style10" name="SVN<%=i%>" value="" readonly>
				<input type="hidden" name="SVI<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<input type="text" size="30" class="style10" name="N<%=i%>" value="" readonly>
				<input type="hidden" name="I<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<div id=DR<%=i%> style="VISIBILITY: visible;">
				<a href="#" onClick="Javascript:AddCharge(<%=i%>);" class="menu"><font color="FFFFFF">Buscar</font></a>
				</div>
			</td>
			<td align="right" class="style4">
				<select class='style10' name='C<%=i%>' id="Moneda">
				<option value='-1'>---</option>
				<%=Currencies%>
				</select>
			</td>
			<td align="center" class="style4">
				<input type="text" size="10" class="style10" name="V<%=i%>" value="" onKeyUp="res(this,numb);" id="Monto">
			</td>
			<td align="center" class="style4">
				<input type="text" size="10" class="style10" name="OV<%=i%>" value="" onKeyUp="res(this,numb);" id="SobreVenta">
			</td>
			<td align="right" class="style4">
				<select class='style10' name='T<%=i%>' id="Tipo de Cobro">
				<option value='-1'>---</option>
				<option value='0'>INT</option>
				<option value='1'>LOC</option>
			 	</select>
			</td>
			<td align="right" class="style4">
				<select class='style10' name='TC<%=i%>' id="Forma de Pago">
				<option value='-1'>---</option>
				<option value='0'>PREPAID</option>
				<option value='1'>COLLECT</option>
			 	</select>
			</td>
            <td align="right" class="style4">
				<select class='style10' name='CCBL<%=i%>' id="Calcular en BL">
				<option value='-1'>---</option>
				<option value='0'>NO</option>
				<option value='1'>SI</option>
			 	</select>
			</td>
			<td align="right" class="style4">
			<div id="DE<%=i%>" style="VISIBILITY: visible;">
			<a href="#" onClick="Javascript:DelCharge(<%=i%>);" class="menu"><font color="FFFFFF">X</font></a>
			</div>
			 </td>
			
			<input type="hidden" name="INV<%=i%>" value="0">
		  </tr>
		  <%next%>
		</table>

		<TABLE cellspacing=0 cellpadding=2 width=200>
		<TR>
			 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
		</TR>
		</TABLE>
		
		<TD>
		</TR>
	</FORM>
	</TABLE>
</BODY>
</HTML>
<script>
<%for i=0 to CountTableValues%>
	document.forma.N<%=i%>.value = '<%=aTableValues(1,i)%>';
	document.forma.I<%=i%>.value = '<%=aTableValues(2,i)%>';
	selecciona('forma.C<%=i%>','<%=aTableValues(3,i)%>');
	document.forma.V<%=i%>.value = '<%=aTableValues(4,i)%>';
	document.forma.OV<%=i%>.value = '<%=aTableValues(5,i)%>';
	selecciona('forma.T<%=i%>','<%=aTableValues(6,i)%>');
	selecciona('forma.TC<%=i%>','<%=aTableValues(7,i)%>');
	document.forma.SVI<%=i%>.value = '<%=aTableValues(8,i)%>';
	document.forma.SVN<%=i%>.value = '<%=aTableValues(9,i)%>';
	document.forma.INV<%=i%>.value = '<%=aTableValues(10,i)%>';
	document.forma.CCBL<%=i%>.value = '<%=aTableValues(11,i)%>';	
	<% if aTableValues(10,i) <> 0 then%>
		document.forma.N<%=i%>.disabled = 'false';
		document.forma.I<%=i%>.disabled = 'false';
		document.forma.C<%=i%>.disabled = 'false';
		document.forma.V<%=i%>.disabled = 'false';
		document.forma.OV<%=i%>.disabled = 'false';
		document.forma.T<%=i%>.disabled = 'false';
		document.forma.TC<%=i%>.disabled = 'false';
		document.forma.SVI<%=i%>.disabled = 'false';
		document.forma.SVN<%=i%>.disabled = 'false';
		document.forma.CCBL<%=i%>.disabled = 'false';
		document.getElementById("DE<%=i%>").style.visibility = "hidden";
		document.getElementById("DR<%=i%>").style.visibility = "hidden";
	<%end if%>
	
<%next
Set aTableValues = Nothing
if Action=1 or Action=2 then%>
	top.opener.location.reload();
<%end if%>
</script>