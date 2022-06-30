<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim ObjectID, Conn, rs, Action, aTableValues, CountTableValues, Currencies, Intercompanies, GroupID, Countries
Dim Name, Agent, i, CantItems, FacID, FacType, FacStatus, AwbType, TableName, AWBNumber, HAWBNumber, DocTyp, BAWResult
Dim ClientCollectID, ClientsCollect, aTableIntercompanies, CountTableIntercompanies

ObjectID = CheckNum(Request("OID"))
Action = CheckNum(Request("Action"))
AwbType = CheckNum(Request("AT"))
DocTyp = AwbType-1
CountTableValues = -1
CantItems = 30

if AwbType=1 then
    TableName= "Awb"
else
    TableName= "Awbi"
end if

OpenConn Conn
	Set rs = Conn.Execute("select AWBNumber, HAWBNumber, Countries, ClientCollectID, ClientsCollect from " & TableName & " where AWBID=" & ObjectID)
	if Not rs.EOF then
		AWBNumber = rs(0)
		HAWBNumber = rs(1)
        Countries = rs(2)
        ClientCollectID = rs(3)
        ClientsCollect = rs(4)
	end if
	CloseOBJ rs

	if Action=2 then
        'Actualizando Cliente a Colectar en Destino
        ClientCollectID = Request("ClientCollectID")
        ClientsCollect = Request("ClientsCollect")
        Conn.Execute("update " & TableName & " set ClientCollectID=" & ClientCollectID & ", ClientsCollect='" & ClientsCollect & "' where AWBID=" & ObjectID)

		'Actualizando los Rubros Intercompany        
		BAWResult = SaveInterChargeItems (Conn, ObjectID, DocTyp, Countries)
	end if

	Set rs = Conn.Execute("select UserID, ItemName, ItemID, CurrencyID, Value, OverSold, Local, PrepaidCollect, ServiceID, ServiceName, InvoiceID, CalcInBL, DocType, '', '', InterCompanyID, ItemName_Routing from ChargeItems where Expired=0 and AWBID=" & ObjectID & " and InterProviderType=5 and InterChargeType=2 and DocTyp= " & DocTyp & " Order By InvoiceID Desc, PrepaidCollect, Local, CurrencyID, ServiceName, ItemName")
	if Not rs.EOF then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	end if
CloseOBJs rs, Conn

OpenConn2 Conn
	'Obteniendo Monedas
	Currencies = Currencies & "<option value=USD>USD</option>"
    'Set rs = Conn.Execute("select distinct simbolo from monedas order by simbolo")
	'Do While Not rs.EOF
		'Currencies = Currencies & "<option value=" & rs(0) & ">" & rs(0) & "</option>"
	'	rs.MoveNext
	'Loop
    'CloseOBJ rs

    'Obteniendo Intercompanies
	Set rs = Conn.Execute("select id_intercompany, nombre_comercial from intercompanys order by nombre_comercial")
	if Not rs.EOF then
        aTableIntercompanies = rs.GetRows
        CountTableIntercompanies = rs.RecordCount-1
	end if

    Intercompanies = ""
    for i=0 to CountTableIntercompanies
        Intercompanies = Intercompanies & "<option value=" & aTableIntercompanies(0,i) & ">" & aTableIntercompanies(1,i) & "</option>"
    next 
    
	'Do While Not rs.EOF
	'	Intercompanies = Intercompanies & "<option value=" & rs(0) & ">" & rs(1) & "</option>"
	'	rs.MoveNext
	'Loop

CloseOBJs rs, Conn

'Seleccion de Serie, Correlativo y Estado de Pago de facturas/ND del BAW
openConnBAW Conn
for i=0 to CountTableValues
	FacID = CheckNum(aTableValues(10,i))
    FacType = CheckNum(aTableValues(12,i))
    FacStatus = 0

    if FacID<>0 then
	    Select case FacType
        case 1
            set rs = Conn.Execute("select tfa_serie, tfa_correlativo, tfa_ted_id from tbl_facturacion where tfa_id=" & FacID)
			    aTableValues(13,i) = "FC-" & rs(0) & "-" & rs(1)
                FacStatus = CheckNum(rs(2))
		    CloseOBJ rs
        case 4
            set rs = Conn.Execute("select tnd_serie, tnd_correlativo, tnd_ted_id from tbl_nota_debito where tnd_id=" & FacID)
			    aTableValues(13,i) = "ND-" & rs(0) & "-" & rs(1)
                FacStatus = CheckNum(rs(2))
		    CloseOBJ rs
        end Select
    End If

        'Indicando el Estado de Pago de la Factura/ND
        select Case FacStatus
        case 2
            aTableValues(14,i) = "<font color=blue>ABONADO</font>"
        case 4
            aTableValues(14,i) = "<font color=blue>PAGADO</font>"
        case Else
            aTableValues(14,i) = "<font color=red>PENDIENTE</font>"
        End Select

next
CloseOBJ Conn

%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
    <%if BAWResult <> "" then %>
    alert("<%=BAWResult%>");
    <%end if %>

	function validar(Action) {
		var sep = '';
		CantItems=-1;
        if (!valTxt(document.forma.ClientsCollect, 1, 5)){return (false)};

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
        document.forma.ItemInterCompanyIDs.value = "";
        document.forma.ItemNames_Routing.value = "";
		
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
                document.forma.ItemInterCompanyIDs.value = document.forma.ItemInterCompanyIDs.value + sep + document.forma.elements["ITCY"+i].value;

                document.forma.ItemNames_Routing.value = document.forma.ItemNames_Routing.value + sep + document.forma.elements["N_Routing"+i].value; //2016-04-01
                
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
        //alert(document.forma.ItemInterCompanyIDs.value);
		
        
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
        document.forma.elements["ITCY"+Pos].value='-1';
		return false;	 
	 }
	 
	 function AddCharge(Pos) {
        if (document.forma.elements["T"+Pos].value != -1) {
    		window.open('Search_Charges.asp?INTR=1&N='+Pos+'&IL='+(document.forma.elements["T"+Pos].value*1+1),'BLData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');
        } else {
            alert('Por favor indique el tipo de este cobro INT o LOC');
            document.forma.elements["T"+Pos].focus();
        }
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

     function GetData(GID){
		window.open('Search_AWBData.asp?GID='+GID,'BLData','height=400,width=460,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
	 };
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
	<div class=label><font color=<%if InStr(BAWResult,"Exitosamente") then %>blue<%else %>red<%end if %>><%=Replace(BAWResult,"\n","<br>")%></font></div>
    <TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InterCharges.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="AT" type=hidden value="<%=AwbType%>">
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
    <INPUT name="ItemInterCompanyIDs" type=hidden value="-1">
    <INPUT name="ItemNames_Routing" type=hidden value="">

		<TD colspan="2" class=label align=center>
		<table width="80%" border="0" align="center">
			<TR><TD class=label align=right><b>Guia Master:</b></TD><TD class=label align=left><%=AWBNumber%></TD></TR> 
			<TR><TD class=label align=right><b>Guia:</b></TD><TD class=label align=left><%=HAWBNumber%></TD></TR>
            <TR>
            <TD class=label align=right><b>Cliente a Colectar:</b></TD>
            <TD class=label align=left width="15%">
                <Input class=style10 type=text name=ClientsCollect value="<%=ClientsCollect%>" size=40 readonly id="Cliente a Colectar"/>
                <Input type=hidden name=ClientCollectID  value="<%=ClientCollectID%>" />
            </TD>
            <TD class=label align=left>
                <div id="CLICOL" style="VISIBILITY: visible;">
                <a href="#" onClick="Javascript:GetData(21);return (false);" class="menu"><font color="FFFFFF"><b>Buscar</b></font></a>
                </div>
            </TD></TR> 
		</table>

        
		
		<table width="80%" border="0">
		  <tr><td class="submenu" colspan="12"></td></tr>
          <tr><td class="style4" colspan="12" align="center">CARGOS A INTERCOMPANY</td></tr>
          <tr><td class="submenu" colspan="12"></td></tr>
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
				Intercompany
			</td>
			<td align="center" class="style4">
				Moneda
			</td>
			<td align="center" class="style4">
				Monto
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
			<td align="center" class="style4">
				Factura/ND
			</td>
            <td align="center" class="style4">
                &nbsp;
			</td>
            <td align="center" class="style4">
				Estado de Pago
			</td>
		  </tr>
		  <%
          
          dim tmp_val, j, internamesel

          for i=0 to CantItems
          
              tmp_val = "0"
              if i <= CountTableValues then 
                tmp_val = aTableValues(16,i)
              end if 
          %>
		  <tr>
			<td align="right" class="style4">
				<input type="text" size="30" class="style10" name="SVN<%=i%>" value="" readonly>
				<input type="hidden" name="SVI<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<input type="text" size="30" class="style10" name="N<%=i%>" value="" readonly>

                <input type="hidden" size="2" name="N_Routing<%=i%>" value="0" >

				<input type="hidden" name="I<%=i%>" value="">
			</td>
			<td align="right" class="style4">
                <% If tmp_val = "0" Then %>
				<div id=DR<%=i%> style="VISIBILITY: visible;">
				<a href="#" onClick="Javascript:AddCharge(<%=i%>);" class="menu"><font color="FFFFFF">Buscar</font></a>
				</div>
                <% End If %>
			</td>
            <td align="center" class="style4">

                <% If tmp_val = "0" Then %>
				    <select class='style10' name='ITCY<%=i%>' id="Intercompany">
				    <option value='-1'>---</option>
				    <%=Intercompanies%>
				    </select>
                <% Else %>
<%
                    internamesel = ""
                    if i <= CountTableValues then 
                        tmp_val = aTableValues(16,i)
                        for j=0 to CountTableIntercompanies
                            if aTableIntercompanies(0,j) =  aTableValues(15,i) then                                                                
                                internamesel = aTableIntercompanies(1,j)
                            end if
                        next 
                    end if
%>
                    <input type="text" size="40" class="style10" value="<%=internamesel%>" readonly >
                    <input type="hidden" size="2" name="ITCY<%=i%>" readonly id="Intercompany">
                <% End If %>
                <input type="hidden" size="10" class="style10" name="OV<%=i%>" value="0">
			</td>
			<td align="right" class="style4">
                <% If tmp_val = "0" Then %>
				    <select class='style10' name='C<%=i%>' id="Moneda">
				    <%=Currencies%>
				    </select>				
                <% Else %>
                    <input type="text" size="5" class="style10" name="C<%=i%>" readonly id="Moneda">
                <% End If %>

			</td>
			<td align="center" class="style4">
                <% If tmp_val = "0" Then %>
	    			<input type="text" size="10" class="style10" name="V<%=i%>" value="" onKeyUp="res(this,numb);" id="Monto">
                <% Else %>
    				<input type="text" size="10" class="style10" name="V<%=i%>" value="" id="Monto" readonly>
                <% End If %>
			</td>
			<td align="right" class="style4">
				<select class='style10' name='T<%=i%>' id="Tipo de Cobro">
				<option value='0'>INT</option>
				</select>
			</td>
			<td align="right" class="style4">
				<select class='style10' name='TC<%=i%>' id="Forma de Pago">
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
				<input type="text" size="10" class="style10" name="FAC<%=i%>" value="" readonly>
			</td>
            <td align="left" class="style4">
                <% If tmp_val = "0" Then %>
                    <div id="DE<%=i%>" style="VISIBILITY: visible;">
			        <a href="#" onClick="Javascript:DelCharge(<%=i%>);" class="menu"><font color="FFFFFF">X</font></a>
			        </div>
                <% End If %>
			</td>
            <td align="left" class="style4">
				<div id="STATFAC<%=i%>" style="VISIBILITY: visible;">
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

    document.forma.N_Routing<%=i%>.value = '<%=aTableValues(16,i)%>';

	document.forma.I<%=i%>.value = '<%=aTableValues(2,i)%>';
	
    //selecciona('forma.C<%=i%>','<%=aTableValues(3,i)%>');
    document.forma.C<%=i%>.value = '<%=aTableValues(3,i)%>';

	document.forma.V<%=i%>.value = '<%=Replace(aTableValues(4,i), ",", ".")%>';    

	document.forma.OV<%=i%>.value = '<%=aTableValues(5,i)%>';
	selecciona('forma.T<%=i%>','<%=aTableValues(6,i)%>');
	selecciona('forma.TC<%=i%>','<%=aTableValues(7,i)%>');
	document.forma.SVI<%=i%>.value = '<%=aTableValues(8,i)%>';
	document.forma.SVN<%=i%>.value = '<%=aTableValues(9,i)%>';
	document.forma.INV<%=i%>.value = '<%=aTableValues(10,i)%>';
	document.forma.CCBL<%=i%>.value = '<%=aTableValues(11,i)%>';
    document.forma.FAC<%=i%>.value = '<%=aTableValues(13,i)%>';
    document.getElementById('STATFAC<%=i%>').innerHTML = '<%=aTableValues(14,i)%>';
	
    //selecciona('forma.ITCY<%=i%>','<%=aTableValues(15,i)%>');
    document.forma.ITCY<%=i%>.value = '<%=aTableValues(15,i)%>';

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
        document.forma.FAC<%=i%>.disabled = 'false';
        document.forma.ITCY<%=i%>.disabled = 'false';
		document.getElementById("DE<%=i%>").style.visibility = "hidden";
		document.getElementById("DR<%=i%>").style.visibility = "hidden";

        document.getElementById("CLICOL").style.visibility = "hidden";
	<%end if%>
	
<%next
Set aTableValues = Nothing
%>
</script>