<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim QuerySelect, GroupID, rs, Conn, i, No, Countries, ServiceID, ItemID, AwbType, ObjectID, aList0Values, aList1Values, CountList0Values, CountList1Values, aList2Values, CountList2Values, aList3Values, CountList3Values, Country, InternationalLocal, IntercompanyFilter, txtbusqueda, RegimenID
Dim Tarifa, tip_moneda, tip_tipo_tarifa, tip_TarifaRango, aiee_TipoAwb, TotWeight, TipoCarga, Msg, esquema, impex, SQLQuery , TPI
Dim Filter, item0, TC


    AwbType = Request("AT")
    Countries = Request("C")
    ObjectID = CheckNum(Request("OID"))
	GroupID = CheckNum(Request("GID"))
	ItemID = CheckNum(Request("ItemID"))
	ServiceID = CheckNum(Request("ServiceID"))
	TPI = Request("TPI")
	'TipoCarga = Request("TipoCarga")
	CountList0Values=-1
	CountList1Values=-1
	CountList2Values=-1
    CountList3Values=-1
    'Obteniendo el pais de Operacion del Usuario
    Country = SetCountryBAW(Session("OperatorCountry"))
    InternationalLocal = CheckNum(Request("IL"))
    IntercompanyFilter = CheckNum(Request("INTR"))
    txtbusqueda = Trim(Request("txtbusqueda"))
    'No = CheckNum(Request("No"))
    esquema = Trim(Request("esquema"))
    impex = Trim(Request("impex"))
    RegimenID = Request("RegimenID")
    TC = Request("TC")


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

            if CDbl(ServiceID) > 0 and CDBL(ItemID) > 0 then
            
                Dim myArray
                myArray = Split(TarifarioPricing (AwbType, Countries, ObjectID, ServiceID, ItemID, No, "", "", "", 0),"|")
           
                'response.write "2-(" & myArray(0) & ")(" & myArray(1) & ")(" & myArray(2) & ")(" & myArray(3) & ")(" & myArray(4) & ")(" & myArray(5) & ")(" & myArray(6) & ")<br>"

                aiee_TipoAwb = myArray(0)
                TotWeight = myArray(1)
                TipoCarga = myArray(2)
                Tarifa = myArray(3)
                tip_tipo_tarifa = myArray(4)
                tip_TarifaRango = myArray(5)
                tip_moneda = myArray(6)
                Msg = myArray(10)

                response.write Msg

            end if

        If Err.Number <> 0 Then
            response.write "<br>Search_Charges Error : " & Err.Number & " - " & Err.description & "<br>"  
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
        'response.write QuerySelect & "<br>"
        Set rs = Conn.Execute(QuerySelect)	    
        If Not rs.EOF Then
		    aList1Values = rs.GetRows
		    CountList1Values = rs.RecordCount-1
	    End If
	    CloseOBJ rs
	

        if ServiceID > 0 then

            SQLQuery = "SELECT er_regimen, er_abreviatura, er_descripcion " & _ 
            "FROM exactus_regimen " & _ 
            "WHERE er_esquema = '" & esquema & "' AND er_status = '1' "
            'response.write SQLQuery & "<br>"
            OpenConn2 Conn
            Set rs = Conn.Execute(SQLQuery)
            If Not rs.EOF Then
		        aList3Values = rs.GetRows
		        CountList3Values = rs.RecordCount-1
	        End If
	        CloseOBJ rs

            if CountList3Values = -1 then
                RegimenID = "IV"
            end if

        else
            RegimenID = ""
        end if


        if RegimenID <> "" and RegimenID <> "0" and RegimenID <> "-1" then 'or CountList3Values = 999 then

            if txtbusqueda <> "" then
                Filter = " AND UPPER(a.desc_rubro_es) LIKE '%" & UCase(txtbusqueda) & "%' "
            end if

		    'Obteniendo listado de rubros asignados al Servicio
            
            if esquema <> "" and impex <> "" then

                SQLQuery = "SELECT c.id_rubro, c.desc_rubro_es, a.codigo, COALESCE(eh_erp_codigo,''), COALESCE(eh_estado,0), COALESCE(eh_otros,'') " & vbCrLf & _ 

                "FROM rubros c " & vbCrLf & _ 

                "INNER JOIN rubros_servicios b ON (c.id_rubro=b.id_rubro AND b.activo = 1) " & vbCrLf & _ 

                "INNER JOIN vw_rubros_combinaciones a ON a.id_servicio = b.id_servicio AND a.id_rubro = c.id_rubro " & vbCrLf & _ 

                "AND a.d1 = '" & impex & "' AND (a.descripcion ILIKE '%aereo%' OR a.serv ILIKE '%aereo%') AND a.d3 = '" & RegimenID & "' AND a.d2 = 'A' " & vbCrLf & _ 

                "LEFT JOIN exactus_homologaciones ON codigo = eh_codigo AND eh_erp_categoria = '06' AND eh_estado = 1 AND eh_erp_esquema = '" & esquema & "' " & vbCrLf & _ 

                "WHERE b.in_conta_baw=1 AND c.id_estatus=1 AND b.id_servicio=" & ServiceID & " " & Filter & " ORDER BY c.desc_rubro_es" 

            else
            
                SQLQuery = "SELECT c.id_rubro, c.desc_rubro_es, '', '', 0, '' " & _ 
                "FROM rubros c  " & _ 
                "INNER JOIN rubros_servicios b ON (c.id_rubro=b.id_rubro AND b.activo = 1)  " & _ 
                "WHERE b.in_conta_baw=1 AND c.id_estatus=1 AND b.id_servicio=" & ServiceID & " " & Filter & " ORDER BY c.desc_rubro_es"

            end if
                        
            'response.write SQLQuery & "<br>"
		    Set rs = Conn.Execute(SQLQuery)
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

    //const d = new Date();
    //document.getElementById("container").innerHTML = dato + ' ' + d.toLocaleTimeString();

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
                tipocarga = document.forma.TipoCarga.value; //.options[document.forma.TipoCarga.selectedIndex].text;

            service = document.forma.ServiceID.options[document.forma.ServiceID.selectedIndex].text;
            p = service.indexOf("(");
            service = service.substring(0, p);

            rubro = document.forma.ItemID.options[document.forma.ItemID.selectedIndex].text;
            p = rubro.indexOf("[");
            rubro = rubro.substring(0, p) + '(' + document.forma.ItemID.value + ')';

            if (document.forma.TipoCarga) {
                iTarifa = document.getElementById('iTarifa').value;
                iMoneda = document.getElementById('iMoneda').value;
                iTipoTarifa = document.getElementById('iTipoTarifa').value;
                iTarifaRango = document.getElementById('iTarifaRango').value; //2022-04-21   
            }

        } catch (err) {
            console.log('1---------------')
            console.log(err)
        }

        var NID = document.forma.NID.value;
        var CM = document.forma.CM.value; //nuevo


        if (NID != "") {
            //Cargos
            var N = document.forma.N.value;

            //var iN = '';
            //iN = N.replace("_Routing", "");
            //iN = iN.replace("ChargeName", "ChargeVal");

            var TPI = document.forma.TPI.value;            

            var Pos2 = '';// NID.substring(1, NID.length);              FALTA DEFINIR
            
            if (document.forma.TipoCarga)
                if (!valSelec(document.forma.TipoCarga)) { return (false) };

            if (!valSelec(document.forma.ServiceID)) { return (false) };
            if (!valSelec(document.forma.ItemID)) { return (false) };



            

            var result = CheckDoble(document.forma.ServiceID.value, document.forma.ItemID.value, NID);

            if (!result) return false;


            try {
                top.opener.document.forma.elements["SVI" + NID].value = document.forma.ServiceID.value;
                top.opener.document.forma.elements["SVN" + NID].value = service; // document.forma.ServiceID.options[document.forma.ServiceID.selectedIndex].text;            
                top.opener.document.forma.elements["I" + NID].value = document.forma.ItemID.value;
                //top.opener.document.forma.elements[N].value = rubro;
                top.opener.document.forma.elements["N" + NID].value = rubro; 

            } catch (err) {
                console.log('2---------------')
                console.log(err)
            }


            if (document.forma.TipoCarga) {

                //alert(document.getElementById('iTarifaRango').value + ' ' + iTarifaRango)

                top.opener.document.forma.elements['TP' + NID].value = iTarifaRango;       // 2022-04-29
                top.opener.document.forma.elements['TT' + NID].value = iTipoTarifa;    // 2022-05-03
                
                //top.opener.document.forma.elements[NID + '_Regimen'].value = iRegimen;

                try {
                    if (document.forma.RegimenID)
                        top.opener.document.forma.elements['R' + NID].value = document.forma.RegimenID.options[document.forma.RegimenID.selectedIndex].value;
                } catch (err) {
                    top.opener.document.forma.elements["R" + NID].value = ''; // '<%=RegimenID%>';        
                }


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
                    //top.opener.document.forma.elements[iN].value = iTarifa;
                    //top.opener.document.forma.elements[iN].readOnly = true;

                    //top.opener.document.forma.elements[CM].value = iMoneda;
                    //top.opener.document.forma.elements[CM].readOnly = true;

                    top.opener.document.forma.elements['C' + NID].value = iMoneda;

                } else {

                    top.opener.document.forma.elements['V' + NID].value = "";
                    top.opener.document.forma.elements['C' + NID].value = "";
                    //top.opener.document.forma.elements[iN].value = "";
                    //top.opener.document.forma.elements[iN].readOnly = false;

                }



            try {

                //Revisando si el cargo que se esta agregando no esta duplicado, permite duplicidad si uno de los 2 ya esta facturado
                if (TPI == 1) {
                    top.opener.CheckAgentDoble(Pos2);
                }
                if (TPI == 0) {
                    top.opener.CheckCarrierDoble(Pos2);
                }
                if (TPI == 2) {
                    top.opener.CheckOtherDoble(Pos2);
                }

            } catch (err) {
                console.log('3---------------')
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




    function CheckDoble(ServiceID, RubID, Pos) {

        var forma = top.opener.document.forma;
        var chk = forma.elements['CHK'];
        var v1 = 0, sigue = false;
        var i;


        for (i = 0; i < chk.length; i++) {

            //alert(ServiceID + ' ' + RubID);

            if ((forma.elements["SVI" + i].value == ServiceID) && (forma.elements["I" + i].value == RubID)) {

                /*
                alert(forma.elements["SVI" + i].value + ' ' +
                    forma.elements["I" + i].value + ' ' +
                    forma.elements["DTY" + i].value + ' ' +
                    forma.elements["PID" + i].value + ' ' +
                    forma.elements["PER" + i].value);
*/

                if (forma.elements["DTY" + i].value != 10) { //10 si esta facturado

                    //if ((forma.elements["PID" + i].value != 0) && (forma.elements["PER" + i].value != "")) { //si ya tiene pedido

                    //} else {

                        if (Pos == i) { //si es el mismo si puede seleccionarse nuevmente

                        } else {

                            if (forma.elements["TC" + i].value == document.forma.TC.value) { // forma de pago

                                alert("No puede repetir el mismo Rubro y Servicio (Forma Pago) si el anterior no ha sido facturado");

                                return (false);
                            }
                        }
                    //}

                }

            }

        }

        return true;

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
    
    <INPUT name="esquema" type=hidden value="<%=esquema%>">
    <INPUT name="impex" type=hidden value="<%=impex%>">
    <INPUT name="TPI" type=hidden value="<%=TPI%>">
	<INPUT name="N" type=hidden value="<%=Request("N")%>">
	<INPUT name="NID" type=hidden value="<%=Request("NID")%>">
	<INPUT name="CM" type=hidden value="<%=Request("CM")%>">
    <INPUT name="TC" type=hidden value="<%=TC%>">

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

        <%="(" & ServiceID & ")(" & CountList3Values & ")(" & RegimenID & ")<br>" %>

		</TD>
		</TR> 


        <%if ServiceID > -1 and CountList3Values > -1 then%>
		    <TR>
		    <TD class=label align=right width="120px"><b>Regimen:</b></TD>
		    <TD class=label align=left>
		        <select name="RegimenID" id="RegimenID" class=label onChange="document.forma.submit();">
			        <option value='-1'>Seleccionar</option>
			        <%for i=0 to CountList3Values%>
			        <option value='<%=aList3Values(1,i)%>'><%=aList3Values(1,i) & " - " & aList3Values(2,i) %></option>
			        <%next%>
		        </select>
		    </TD>
		    </TR> 
        <% else %>

            <INPUT name="RegimenID" type=hidden value="<%=RegimenID%>">

        <%end if%>

		<% if RegimenID <> "" and RegimenID <> "0" and RegimenID <> "-1" then %>
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

                    <option value='<%=aList2Values(0,i)%>' 
                    <%
                    if aList2Values(3,i) <> "" then
                        response.write " style='background-color:rgb(255,241,193)' "
                    end if
                    %>
                    >
                    <%
                    response.write aList2Values(1,i) 

                    if aList2Values(2,i) <> "" then
                        response.write " [" & aList2Values(2,i) & "]"
                    end if
            
                    if aList2Values(3,i) <> "" then
                        response.write " " & aList2Values(3,i) & ""
                    else
                        response.write " NO HOMOLOGADO"
                    end if

                    if aList2Values(5,i) <> "" then
                        response.write " (" & aList2Values(5,i) & ")"
                    end if
                    %>
                    </option>

                <%end if%>			    
	
			<%next%>
		</select>

                            <!-- <option value='<%=aList2Values(0,i)%>'><%=aList2Values(1,i) & " (" & aList2Values(0,i) & ")"%></option> -->

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
    item0 = False

	if CheckNum(ItemID) > 0 and CheckNum(ServiceID) > 0 then

        if aiee_TipoAwb = "Master-Hija" or aiee_TipoAwb = "Hija-Directa" or aiee_TipoAwb = "Master-Master-Hija" then
  
            if TotWeight > 0 then
                item0 = True
            end if

        else
            item0 = True
        end if

    end if

'    response.write "(" & item0 & ")"


    if item0 = True then
%>
		 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:Asign();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
<%
    else
%>
		 <TD class=label align=center colspan=2><INPUT name=enviar type=button disabled value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
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
    selecciona('forma.RegimenID', '<%=RegimenID%>');
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
