<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim QuerySelect, GroupID, rs, Conn, i, No, Countries, ServiceID, ItemID, AwbType, ObjectID, aList0Values, aList1Values, CountList0Values, CountList1Values, aList2Values, CountList2Values, Country, InternationalLocal, IntercompanyFilter, txtbusqueda
Dim Tarifa, tip_moneda, tip_tipo_tarifa, tip_TarifaRango, aiee_TipoAwb, TotWeight, TipoCarga 
Dim Filter

    AwbType = Request("AT")
    Countries = Request("C")
    ObjectID = CheckNum(Request("OID"))
	GroupID = CheckNum(Request("GID"))
	ItemID = CheckNum(Request("ItemID"))
	ServiceID = CheckNum(Request("ServiceID"))
	'TipoCarga = Request("TipoCarga")
	CountList0Values=-1
	CountList1Values=-1
	CountList2Values=-1
    'Obteniendo el pais de Operacion del Usuario
    Country = SetCountryBAW(Session("OperatorCountry"))
    InternationalLocal = CheckNum(Request("IL"))
    IntercompanyFilter = CheckNum(Request("INTR"))
    txtbusqueda = Trim(Request("txtbusqueda"))
    'No = CheckNum(Request("No"))

    if InternationalLocal = 0 then
        InternationalLocal = "1,2"
    end if
    
    if IntercompanyFilter = 1 then
        IntercompanyFilter = " and c.id_servicio<>14 "        
    else
        IntercompanyFilter = ""
    end if    

    Filter = ""

	'Si el usuario tiene pais de Operacion se le presentan los servicios autorizados para cobrar/pagar de ese pais
	if Country>0 then

        On Error Resume Next

            Dim myArray
            myArray = Split(TarifarioPricing (AwbType, Countries, ObjectID, ServiceID, ItemID, No),"|")
            
            'response.write "2-(" & myArray(0) & ")(" & myArray(1) & ")(" & myArray(2) & ")(" & myArray(3) & ")(" & myArray(4) & ")(" & myArray(5) & ")(" & myArray(6) & ")<br>"

            aiee_TipoAwb = myArray(0)
            TotWeight = myArray(1)
            TipoCarga = myArray(2)
            Tarifa = myArray(3)
            tip_tipo_tarifa = myArray(4)
            tip_TarifaRango = myArray(5)
            tip_moneda = myArray(6)

        If Err.Number <> 0 Then
            response.write "<br>Search Error : " & Err.Number & " - " & Err.description & "<br>"  
        end if
        
        OpenConn2 Conn
        'Obteniendo listado de Servicios
	    'tass_tsis_id ID del sistema de la tabla transporte, Aereo=1
        'tas_ttt_id es tipo de ID, 0=sistema, 1=RO


        QuerySelect = "select a.id_servicio, c.nombre_servicio " & _
        "from empresas_transportes_servicios as a " & _
        "inner join transporte as b on(b.id_transporte=a.id_transporte) " & _
        "inner join servicios as c on(c.id_servicio=a.id_servicio) " & _
        "inner join empresas as d on(d.id_empresa=a.id_empresa and d.activo=true) " & _
        "where ( d.id_empresa=" & Country & " and a.id_transporte=1 and a.activo=true ) and a.cargo_int_loc in ( "& InternationalLocal &",3 ) " & _
        IntercompanyFilter & _
        "order by c.nombre_servicio"
        'response.write ( QuerySelect )
        Set rs = Conn.Execute(QuerySelect)
	    
        If Not rs.EOF Then
		    aList1Values = rs.GetRows
		    CountList1Values = rs.RecordCount-1
	    End If
	    CloseOBJ rs
	
	    if ServiceID > 0 then

            if txtbusqueda <> "" then
                Filter = " AND UPPER(a.desc_rubro_es) LIKE '%" & UCase(txtbusqueda) & "%' "
            end if

		    'Obteniendo listado de rubros asignados al Servicio
            QuerySelect = "select a.id_rubro, a.desc_rubro_es from rubros a, rubros_servicios b where a.id_rubro=b.id_rubro and b.in_conta_baw=1 and a.id_estatus=1 and b.id_servicio=" & ServiceID & " and b.activo=1  " & Filter & " order by a.desc_rubro_es"
            'response.write ( QuerySelect )
		    Set rs = Conn.Execute(QuerySelect)
            If Not rs.EOF Then
			    aList2Values = rs.GetRows
			    CountList2Values = rs.RecordCount-1
		    End If
		    CloseOBJ rs
	    end if
        CloseOBJ Conn
    end if
%>
 
<HTML><HEAD><TITLE>Sistema Aereo</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
    function Asign() {

        var p;
        var tipocarga
        var iTarifa;
        var iTarifaRango;
        var iMoneda;
        var iTipoTarifa = '';
        var iRegimen = 'X';
        var service = '';
        var rubro = '';


try {
        if (document.forma.TipoCarga)
            tipocarga = document.forma.TipoCarga.options[document.forma.TipoCarga.selectedIndex].text;

        service = document.forma.ServiceID.options[document.forma.ServiceID.selectedIndex].text;
        p = service.indexOf("(");
        service = service.substring(0, p);

        rubro = document.forma.ItemID.options[document.forma.ItemID.selectedIndex].text;
        p = rubro.indexOf("(");
        rubro = rubro.substring(0, p);


} catch (err) {
    console.log(err)
}


        if (document.forma.TipoCarga) {

            iTarifa = document.getElementById('iTarifa').value;

            iMoneda = document.getElementById('iMoneda').value;

            iTipoTarifa = document.getElementById('iTipoTarifa').value;

            iTarifaRango = document.getElementById('iTarifaRango').value; //2022-04-21   
        }

        var NID = document.forma.NID.value;
        var CM = document.forma.CM.value; //nuevo


        if (NID != "") {
            //Cargos
            var N = document.forma.N.value;

            var iN = '';

            iN = N.replace("_Routing", "");
            iN = iN.replace("ChargeName", "ChargeVal");

            var TPI = NID.substring(0, 1);
            var Pos2 = NID.substring(1, NID.length);
            if (document.forma.TipoCarga)
            if (!valSelec(document.forma.TipoCarga)) { return (false) };
            if (!valSelec(document.forma.ServiceID)) { return (false) };
            if (!valSelec(document.forma.ItemID)) { return (false) };



try {
            top.opener.document.forma.elements["SVI" + NID].value = document.forma.ServiceID.value;
            top.opener.document.forma.elements["SVN" + NID].value = service; // document.forma.ServiceID.options[document.forma.ServiceID.selectedIndex].text;            
            top.opener.document.forma.elements[NID].value = document.forma.ItemID.value;
            top.opener.document.forma.elements[N].value = rubro; //  + ' ' + iTipoTarifa; // document.forma.ItemID.options[document.forma.ItemID.selectedIndex].text;
            top.opener.document.forma.elements["N" + NID].value = rubro; // document.forma.ItemID.options[document.forma.ItemID.selectedIndex].text;

} catch (err) {
    console.log(err)
}


if (document.forma.TipoCarga) {

    //alert(document.getElementById('iTarifaRango').value + ' ' + iTarifaRango)

    top.opener.document.forma.elements[NID + '_Tarifa'].value = iTarifaRango;       // 2022-04-29
    top.opener.document.forma.elements[NID + '_TarifaTipo'].value = iTipoTarifa;    // 2022-05-03
    top.opener.document.forma.elements[NID + '_Regimen'].value = iRegimen;

}
            //alert(N);
            //este valor se traslada unicamente para Air Freight para que realice el calculo correctamente
            if (N == 'TotCarrierRate_Routing') {

                //alert(top.opener.document.forma.elements['CarrierRates'].innerHTML + ' ' + iTarifaRango);
                //top.opener.document.forma.elements['CarrierRates'].innerHTML = iTarifaRango;

                //alert(top.opener.document.forma.elements['CarrierRates'].value + ')(' + iTarifaRango);
                top.opener.document.forma.elements['CarrierRates'].value = iTarifaRango;

                top.opener.CalcRate(top.opener.document.forms[0])            
            }

            

            if (document.forma.TipoCarga)
            if (iTarifa > 0) {

                top.opener.document.forma.elements['V' + NID].value = iTarifa;
                top.opener.document.forma.elements[iN].value = iTarifa;
                top.opener.document.forma.elements[iN].readOnly = true;

                top.opener.document.forma.elements[CM].value = iMoneda;
                top.opener.document.forma.elements[CM].readOnly = true;

            } else {

                top.opener.document.forma.elements['V' + NID].value = "";
                top.opener.document.forma.elements[iN].value = "";
                top.opener.document.forma.elements[iN].readOnly = false;

            }

            //alert(NID);
            //alert(N);
            //alert(document.forma.ServiceID.options[document.forma.ServiceID.selectedIndex].text);
            //alert(document.forma.ItemID.options[document.forma.ItemID.selectedIndex].text);



            try {

                //Revisando si el cargo que se esta agregando no esta duplicado, permite duplicidad si uno de los 2 ya esta facturado
                if (TPI == 'A') {
                    top.opener.CheckAgentDoble(Pos2);
                }
                if (TPI == 'C') {
                    top.opener.CheckCarrierDoble(Pos2);
                }
                if (TPI == 'O') {
                    top.opener.CheckOtherDoble(Pos2);
                }

            } catch (err) {
                console.log(err)
            }



        } else {
            //Costos
            var Pos = document.forma.N.value;
            if (!valSelec(document.forma.TipoCarga)) { return (false) };
            if (!valSelec(document.forma.ServiceID)) { return (false) };
            if (!valSelec(document.forma.ItemID)) { return (false) };
            top.opener.document.forma.elements["SVI" + Pos].value = document.forma.ServiceID.value;
            top.opener.document.forma.elements["SVN" + Pos].value = service; // document.forma.ServiceID.options[document.forma.ServiceID.selectedIndex].text;
            top.opener.document.forma.elements["I" + Pos].value = document.forma.ItemID.value;
            top.opener.document.forma.elements["N" + Pos].value = rubro; // document.forma.ItemID.options[document.forma.ItemID.selectedIndex].text;
            top.opener.ValidarDoble(Pos);
        }
        top.close();
    }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus()">
	<FORM name="forma" action="Search_Charges.asp" method="post" target=_self>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
    <INPUT name="OID" type=hidden value="<%=ObjectID%>">
    <INPUT name="IL" type=hidden value="<%=CheckNum(Request("IL"))%>">
    <INPUT name="AT" type=hidden value="<%=AwbType%>">
    <INPUT name="iTarifa" id="iTarifa" type=hidden value="<%=Tarifa%>">
    <INPUT name="iMoneda" id="iMoneda" type=hidden value="<%=tip_moneda%>">
    <INPUT name="iTipoTarifa" id="iTipoTarifa" type=hidden value="<%=tip_tipo_tarifa%>">
    <INPUT name="TipoCarga" id="TipoCarga" type=hidden value="<%=TipoCarga%>">
    <INPUT name="iTarifaRango" id="iTarifaRango" type=hidden value="<%=tip_TarifaRango%>"> <!-- el valor del rango correspondiente antes de la multiplicacion por el peso -->

    
    <INPUT name="No" id="No" type=hidden value="<%=No%>">
    <INPUT name="C" type=hidden value="<%=Countries%>">
    
    

	<br>
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<TR>
		<TD class=label align=center colspan="2"><h3>Rubros</h3></TD>
		</TR> 



		<TR>
		<TD class=label align=right width="120px"><b>Servicio:</b></TD>
		<TD class=label align=left>
		<select name="ServiceID" id="Servicio" class=label onChange="if (document.getElementById('Rubro')) { document.getElementById('Rubro').value=''; } document.forma.submit();">
			<option value='-1'>Seleccionar</option>
			<%for i=0 to CountList1Values%>
			<option value='<%=aList1Values(0,i)%>'><%=aList1Values(1,i) & " (" & aList1Values(0,i) & ")"%></option>
			<%next%>
		</select>
		<INPUT name="N" type=hidden value="<%=Request("N")%>">
		<INPUT name="NID" type=hidden value="<%=Request("NID")%>">
		<INPUT name="CM" type=hidden value="<%=Request("CM")%>">
		</TD>
		</TR> 
		<%if ServiceID>0 then%>
        <!--
		<TR>
		<TD class=label align=right><b>Busqueda:</b></TD>
		<TD class=label align=left>
            <input type="text" id="txtbusqueda" name="txtbusqueda" value="<%=txtbusqueda%>" />

            <input type="submit" value="Buscar" />

            <input type="submit" value="Limpiar" onclick="document.getElementById('txtbusqueda').value = '';" />
		</TD>
		</TR> 
        -->
		<TR>
		<TD class=label align=right><b>Rubro:</b></TD>
		<TD class=label align=left>
		<select name="ItemID" id="Rubro" class=label onChange="document.forma.submit();" style="width:300px">
			<option value='-1'>Seleccionar</option>
			<%for i=0 to CountList2Values%>

			    <%if GroupID <> 18 or CInt(aList2Values(0,i)) <> 486 then '2015-03-03 REBATE %>
                        <option value='<%=aList2Values(0,i)%>'><%=aList2Values(1,i) & " (" & aList2Values(0,i) & ")"%></option>
                <%end if%>			    
	
			<%next%>
		</select>
		</TD>
		</TR> 
		<%end if%>


<%
    if aiee_TipoAwb = "Master-Hija" or aiee_TipoAwb = "Hija-Directa" or aiee_TipoAwb = "Master-Master-Hija" then
%>
		<TR>
		<TD class=label align=right><b>Tipo Carga : </b></TD>
		<TD class=label align=left><%=TipoCarga%></TD>
		</TR> 

		<TR>
		<TD class=label align=right><b>Clase : </b></TD>
		<TD class=label align=left><%=tip_tipo_tarifa%></TD>
		</TR> 

		<TR>
		<TD class=label align=right><b>Tarifa:</b></TD>
		<TD class=label align=left><%=Tarifa%></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Moneda:</b></TD>
		<TD class=label align=left><%=tip_moneda%></TD>
		</TR> 

        
<%
    end if
%>


	</TABLE>
	<TABLE cellspacing=0 cellpadding=2 width=100%>
	<TR>

<%

    if aiee_TipoAwb = "Master-Hija" or aiee_TipoAwb = "Hija-Directa" or aiee_TipoAwb = "Master-Master-Hija" then
  
        if TotWeight > 0 then
%>
		 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:Asign();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
<%
        else
        
%>
		 <TD class=label align=center colspan=2><INPUT name=enviar type=button disabled value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
<%
        end if

    else

%>
		 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:Asign();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
<%

    end if
%>
	</TR>
	</TABLE>



	</FORM>
<script>   
    if (document.forma.TipoCarga)
    selecciona('forma.TipoCarga', '<%=TipoCarga%>');
    selecciona('forma.ServiceID', '<%=ServiceID%>');
    selecciona('forma.ItemID', '<%=ItemID%>');
</script>
</BODY>
</HTML>
<%
    Set aList1Values = Nothing
    Set aList2Values = Nothing
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
