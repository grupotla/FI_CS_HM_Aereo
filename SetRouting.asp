<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
	Checking "0|1|2"
%>
<HTML><HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<%
	Dim Conn, rs, ObjectID, Name, Address, Phone1, Phone2, Attn, AccountNo, IATANo, i
	Dim QuerySelect, GroupID, AddressID, CommodityID, Commodity, Routing, RoutingID
	Dim RoutingValues, CountRoutingValues, AWBType, RoutingCharges, CountRoutingCharges
    Dim ConsignerColoader, ShipperColoader, AgentNeutral, RoutingInterCharges, CountRoutingInterCharges
    Dim aList2Values, aList3Values, aList4Values, aList5Values, aList6Values, aList7Values, aList8Values, aList9Values
    Dim aList10Values, aList11Values, aList12Values, aList13Values, aList14Values, Val, ClientsCollect, id_coloader

	RoutingID = CheckNum(Request("RID"))
	AWBType = CheckNum(Request("AT"))
	CountRoutingCharges = -1
	CountRoutingValues = -1
    CountRoutingInterCharges = -1

	if RoutingID <> 0 then

	OpenConn2 Conn
		'Query para obtener datos del Routing
        'response.write "select id_routing, routing, id_cliente, id_shipper, agente_id, comodity_id, no_piezas, peso, vendedor_id, carrier_id, airportid_embarque, airportid_desembarque, prepaid, vendedor_id, simbolo, id_colectar, id_coloader, seguro, poliza_seguro from routings, unidad_medida where id_unidad_peso=id_unidad_medida and id_routing=" & RoutingID
        '                       0           1       2           3           4           5           6           7       8           9               10                  11              12          13          14      15          16              17      18          19            20       21          22           23                 24       25
        QuerySelect = "select id_routing, routing, id_cliente, id_shipper, agente_id, comodity_id, no_piezas, peso, vendedor_id, carrier_id, airportid_embarque, airportid_desembarque, prepaid, vendedor_id, simbolo, id_colectar, id_coloader, seguro, poliza_seguro, routing_seg, bl_id, routing_adu, routing_ter, id_cliente_order, id_pais, tarifa_minimo from routings, unidad_medida where id_unidad_peso=id_unidad_medida and id_routing=" & RoutingID        
        'response.write QuerySelect 
        Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			RoutingValues = rs.GetRows
			CountRoutingValues = rs.Recordcount -1             
            'response.write("(" & RoutingValues(25,0) & ")") ' 2016-03-29
		end if
		CloseOBJ rs		
%>
<SCRIPT LANGUAGE="JavaScript">

    var seguro_poliza = false;

    if (top.opener.document.forms[0].CarrierID.value == '<%=RoutingValues(9,0)%>' && 
    top.opener.document.forms[0].AirportDepID.value == '<%=CheckNum(RoutingValues(10,0))%>' && 
    top.opener.document.forms[0].AirportDesID.value == '<%=CheckNum(RoutingValues(11,0))%>') {

    } else { //2018-01-10    
        seguro_poliza = true;
        alert("NOTA: La Linea Aereo, ó Aeropuertos no coinciden con los ingresados en AWB");
        alert('Carrier:<%=RoutingValues(9,0)%> \n AiportDep:<%=CheckNum(RoutingValues(10,0))%> \n AiportDes:<%=CheckNum(RoutingValues(11,0))%>');
        top.close();   
    }

       
    <% if RoutingValues(17,0) = "1" then %> //seguro

        <% if CheckNum(RoutingValues(18,0)) = 0 then %> //poliza_seguro    

        seguro_poliza = true; //esto esta de mas
        alert("NOTA: \n<%=RoutingValues(1,0)%> \nEste RO tiene solicitud de seguro, pero aun no \nle han asignado su Poliza, por favor contactar \nal departamento de Ventas.");
        top.close();                        
        
        <% end if %>

    <% end if %>


    <% if RoutingValues(20,0) <> "0" then %> //bl_id
        
        seguro_poliza = true;
        alert("NOTA: \n<%=RoutingValues(1,0)%> \nEste RO esta asociado a otra Guia.");
        top.close();                        
        
    <% end if %>

    

    if (seguro_poliza == false) {
        
        ////////////////////////////////////////////////////////agregado 2016-12-07 actualiza el country
        //en Awb / Awbi ya no lee country de la aerolinea, del usuario
        

        var a = '<%=Replace(Session("Countries"),"'","")%>';

        

        var b = a.split(',');
        b = b[0];        
        b = b.replace(/\(/g, "");
        b = b.replace(/\)/g, "");

        //console.log ("(" + b + ")")
        
        var c = '<%=RoutingValues(24,0)%>';
        var d = c.length;

        if (d == 5) {
            d = c.substring(2,5);
        } else {
            d = ""
        }

        //console.log ("(" + d + ")")
        
        var country = b.substring(0,2) + d;
        //console.log(country);
    top.opener.document.forms[0].Countries.value = country;//'<%=RoutingValues(24,0)%>';
        ////////////////////////////////////////////////////////////


    //mbl_rate_comment

    if ('<%=RoutingValues(25,0)%>' != '') {
        top.opener.document.forms[0].iMinimo.checked = true;
        top.opener.document.forms[0].iMinimo.value = '<%=RoutingValues(25,0)%>';
    }

	top.opener.document.forms[0].RoutingID.value = '<%=RoutingValues(0,0)%>';
    top.opener.document.forms[0].Seguro.value = '<%=RoutingValues(17,0)%>';
    top.opener.document.forms[0].routing_seg.value = '<%=RoutingValues(19,0)%>';
	top.opener.document.forms[0].Routing.value = '<%=RoutingValues(1,0)%>';
    top.opener.document.forms[0].SalespersonID.value = <%=CheckNum(RoutingValues(8,0))%>;
	top.opener.document.forms[0].RAirportDepID.value = <%=CheckNum(RoutingValues(10,0))%>;
	top.opener.document.forms[0].RAirportDesID.value = <%=CheckNum(RoutingValues(11,0))%>;
    
	<%if RoutingValues(12,0) = true then 'true=1=prepaid, false=2=collect%>
		top.opener.document.forms[0].ChargeType.value = 1;
		top.opener.document.forms[0].ValChargeType.value = 1;
		top.opener.document.forms[0].OtherChargeType.value = 1;
	<%else%>
		top.opener.document.forms[0].ChargeType.value = 2;
		top.opener.document.forms[0].ValChargeType.value = 2;
		top.opener.document.forms[0].OtherChargeType.value = 2;
	<%end if%>
	top.opener.document.forms[0].SalespersonID.value = <%=CheckNum(RoutingValues(13,0))%>;
    top.opener.document.forms[0].ClientCollectID.value = <%=CheckNum(RoutingValues(15,0))%>;    
    
    }
    
</SCRIPT>
<%
		'TEMPORAL, obteniendo el AddressID del Consignatario
		Set rs = Conn.Execute("select id_direccion from direcciones where id_cliente=" & RoutingValues(2,0))
		if Not rs.EOF then
			AddressID = rs(0)
		end if
		CloseOBJ rs		
		'Obteniendo Datos del Consignatario
		MasterData Conn, RoutingValues(2,0), AddressID, Name, Address, Phone1, Phone2, Attn, AccountNo, IATANo, ConsignerColoader

%>
<SCRIPT LANGUAGE="JavaScript">
if (seguro_poliza == false) {
	top.opener.document.forms[0].AccountConsignerNo.value = '<%=AccountNo%>';
	top.opener.document.forms[0].ConsignerData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	top.opener.document.forms[0].ConsignerID.value=<%=RoutingValues(2,0)%>;
	top.opener.document.forms[0].ConsignerAddrID.value=<%=AddressID%>;
    top.opener.document.forms[0].ConsignerColoader.value=<%=ConsignerColoader%>;
}
</SCRIPT>
<%
        if CheckNum(RoutingValues(16,0)) > 0 then

		'TEMPORAL, obteniendo el AddressID del Coloader
		Set rs = Conn.Execute("select id_direccion from direcciones where id_cliente=" & RoutingValues(16,0))
		if Not rs.EOF then
			AddressID = rs(0)
		end if
		CloseOBJ rs		
		'Obteniendo Datos del Consignatario
		MasterData Conn, RoutingValues(16,0), AddressID, Name, Address, Phone1, Phone2, Attn, AccountNo, IATANo, ConsignerColoader

%>
<SCRIPT LANGUAGE="JavaScript">	
if (seguro_poliza == false) {
    top.opener.document.forms[0].id_coloader.value = <%=CheckNum(RoutingValues(16,0))%>;
	top.opener.document.forms[0].ColoaderData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
}
</SCRIPT>
<%

        end if


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////// 2016-10-11

        if CheckNum(RoutingValues(23,0)) > 0 then

		'TEMPORAL, obteniendo el AddressID del Notify
		Set rs = Conn.Execute("select id_direccion from direcciones where id_cliente=" & RoutingValues(23,0))
		if Not rs.EOF then
			AddressID = rs(0)
		end if
		CloseOBJ rs		
		'Obteniendo Datos del Notify
		MasterData Conn, RoutingValues(23,0), AddressID, Name, Address, Phone1, Phone2, Attn, AccountNo, IATANo, ConsignerColoader
%>
<SCRIPT LANGUAGE="JavaScript">
if (seguro_poliza == false) {	
	top.opener.document.forms[0].id_cliente_orderData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	top.opener.document.forms[0].id_cliente_order.value=<%=RoutingValues(23,0)%>;	
}
</SCRIPT>
<%
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        end if

		'TEMPORAL, obteniendo el AddressID del Shipper
		Set rs = Conn.Execute("select id_direccion from direcciones where id_cliente=" & RoutingValues(3,0))
		if Not rs.EOF then
			AddressID = rs(0)
		end if
		CloseOBJ rs		
		'Obteniendo Datos del Shipper
		MasterData Conn, RoutingValues(3,0), AddressID, Name, Address, Phone1, Phone2, Attn, AccountNo, IATANo, ShipperColoader
%>
<SCRIPT LANGUAGE="JavaScript">
if (seguro_poliza == false) {
	top.opener.document.forms[0].AccountShipperNo.value = '<%=AccountNo%>';
	top.opener.document.forms[0].ShipperData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	top.opener.document.forms[0].ShipperID.value=<%=RoutingValues(3,0)%>;
	top.opener.document.forms[0].ShipperAddrID.value=<%=AddressID%>;
    top.opener.document.forms[0].ShipperColoader.value=<%=ShipperColoader%>;
}
</SCRIPT>
<%
		'TEMPORAL, obteniendo Datos del Agente
		Set rs = Conn.Execute("select agente, direccion, telefono, fax, contacto, es_neutral from agentes where agente_id=" & RoutingValues(4,0))
		if Not rs.EOF then
			Name = rs(0)
			Address = rs(1)
			Phone1 = rs(2)
			Phone2 = rs(3)
			Attn = rs(4)
			AddressID = 0
            AgentNeutral = CheckNum(rs(5))
		end if
		CloseOBJ rs		
%>
<SCRIPT LANGUAGE="JavaScript">
if (seguro_poliza == false) {
	top.opener.document.forms[0].AgentData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	top.opener.document.forms[0].AgentID.value=<%=RoutingValues(4,0)%>;
	top.opener.document.forms[0].AgentAddrID.value=<%=AddressID%>;
    top.opener.document.forms[0].AgentNeutral.value=<%=AgentNeutral%>;
}
</SCRIPT>
<%
	 	'TEMPORAL, obteniendo nombre del cliente a Colectar en Intercompany
		Set rs = Conn.Execute("select b.nombre_cliente from direcciones a, clientes b where a.id_cliente=b.id_cliente and b.id_cliente=" & RoutingValues(15,0))
		if Not rs.EOF then
			ClientsCollect = rs(0)
		end if
		CloseOBJ rs		

		'Obteniendo los rubros Intercompany del Routing
		'response.write "select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' from cargos_routing a, rubros b, monedas c where id_routing=" & RoutingValues(0,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id"
        QuerySelect = "select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' as nada, a.inter_company, a.prepaid from cargos_routing a, rubros b, monedas c, routings d where d.id_routing=" & RoutingValues(0,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id and a.id_routing=d.id_routing and d.borrado=false and a.inter_company<>0"
        if RoutingValues(19,0) > 0 then '2016-03-29 se agrego este union de seguros
            QuerySelect = QuerySelect & " UNION select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' as nada, a.inter_company, a.prepaid from cargos_routing a, rubros b, monedas c, routings d where d.id_routing=" & RoutingValues(19,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id and a.id_routing=d.id_routing and d.borrado=false and a.inter_company<>0"
            'QuerySelect = QuerySelect & " and d.activo=true and d.borrado=false and d.routing_fac=0 "
        end if        


        if RoutingValues(21,0) > 0 then '2016-04-21 
            QuerySelect = QuerySelect & " UNION select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' as nada, a.inter_company, a.prepaid from cargos_routing a, rubros b, monedas c, routings d where d.id_routing=" & RoutingValues(21,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id and a.id_routing=d.id_routing and d.borrado=false and a.inter_company<>0"            
        end if        

        if RoutingValues(22,0) > 0 then '2016-04-21
            QuerySelect = QuerySelect & " UNION select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' as nada, a.inter_company, a.prepaid from cargos_routing a, rubros b, monedas c, routings d where d.id_routing=" & RoutingValues(22,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id and a.id_routing=d.id_routing and d.borrado=false and a.inter_company<>0"
        end if 


        'QuerySelect = "SELECT * FROM (" & QuerySelect & ") x ORDER BY local, simbolo, valor DESC"
        QuerySelect = "SELECT simbolo, id_rubro, valor, local, desc_rubro_es, id_servicio, nada, inter_company, CASE prepaid WHEN 't' THEN '0' ELSE '1' END AS prepaid FROM (" & QuerySelect & ") x ORDER BY local, simbolo, valor DESC"

        'response.write( "1.////////////////////////////////////<br>" & QuerySelect & "<br>" )
        Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			RoutingInterCharges = rs.GetRows
			CountRoutingInterCharges = rs.RecordCount-1
		end if
        CloseOBJ rs


        'response.write("por aqui va 1<br>")
		
        'c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '', a.intercompany
		Val = "" 'separador
		for i=0 to CountRoutingInterCharges
            'Obteniendo los nombres de los servicios del Routing
            Set rs = Conn.Execute("select nombre_servicio from servicios where id_servicio=" & CheckNum(RoutingInterCharges(5,i)))
			if Not rs.EOF then
				RoutingInterCharges(6,i) = rs(0)
			end if
            CloseOBJ rs

			aList2Values = aList2Values & Val & RoutingInterCharges(0,i) 'simbolo
			aList3Values = aList3Values & Val & CInt(RoutingInterCharges(1,i)) 'id_rubro
			aList4Values = aList4Values & Val & replace(RoutingInterCharges(2,i),",",".") 'valor
			aList5Values = aList5Values & Val & RoutingInterCharges(3,i) 'local
			aList6Values = aList6Values & Val & RoutingInterCharges(4,i) 'desc_rubro_es
			aList7Values = aList7Values & Val & "0" 'sobreventa: en routing no hay

			'aList8Values = aList8Values & Val & RoutingValues(12,0) 'true=1=prepaid, false=2=collect            
            aList8Values = aList8Values & Val & RoutingInterCharges(8,i)

			aList9Values = aList9Values & Val & RoutingInterCharges(5,i) 'id_servicio
			aList10Values = aList10Values & Val & RoutingInterCharges(6,i) 'nombre_servicio
			aList11Values = aList11Values & Val & "0" 'factura ID
			aList12Values = aList12Values & Val & "1" 'Si se debe calcular en el BL, el usuario puede cambiarlo luego en "Cobros y Documentos"
            aList13Values = aList13Values & Val & RoutingInterCharges(7,i) 'ID del Intercompany
            aList14Values = aList14Values & Val & "1" 'ID del Intercompany
            
			Val = "|"
		next
		Set RoutingInterCharges = Nothing

        'response.write("por aqui va 2<br>")

%>
<SCRIPT LANGUAGE="JavaScript">
    if (seguro_poliza == false) {
        top.opener.document.forms[0].ClientCollectID.value = '<%=RoutingValues(15,0)%>';
        top.opener.document.forms[0].ClientsCollect.value = '<%=ClientsCollect%>';

        top.opener.document.forms[0].ItemCurrs.value = '<%=aList2Values%>';
        top.opener.document.forms[0].ItemIDs.value = '<%=aList3Values%>';
        top.opener.document.forms[0].ItemVals.value = '<%=aList4Values%>';
        top.opener.document.forms[0].ItemLocs.value = '<%=aList5Values%>';
        top.opener.document.forms[0].ItemNames.value = '<%=aList6Values%>';
        top.opener.document.forms[0].ItemOVals.value = '<%=aList7Values%>';
        top.opener.document.forms[0].ItemPPCCs.value = '<%=aList8Values%>';
        top.opener.document.forms[0].ItemServIDs.value = '<%=aList9Values%>';
        top.opener.document.forms[0].ItemServNames.value = '<%=aList10Values%>';
        top.opener.document.forms[0].ItemInvoices.value = '<%=aList11Values%>';
        top.opener.document.forms[0].ItemCalcInBls.value = '<%=aList12Values%>';
        top.opener.document.forms[0].ItemIntercompanyIDs.value = '<%=aList13Values%>';
        top.opener.document.forms[0].ItemNames_Routing.value = '<%=aList14Values%>';        
        top.opener.document.forms[0].CantItems.value = '<%=CountRoutingInterCharges%>';
    }
</SCRIPT>
<%
		'Obteniendo Datos del Producto
		'response.write "select a.CommodityCode, a.NameES from Commodities a where commodityid=" & RoutingValues(5,0) & "<br>"
		Set rs = Conn.Execute("select a.CommodityID, a.NameES from Commodities a where commodityid=" & CheckNum(RoutingValues(5,0)))
		if Not rs.EOF then
			CommodityID = rs(0)
			Commodity = rs(1)
		end if
		CloseOBJ rs

		'Obteniendo los rubros del Routing
		'response.write "select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' from cargos_routing a, rubros b, monedas c where id_routing=" & RoutingValues(0,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id"

        QuerySelect = "select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' as nada, a.inter_company, a.prepaid from cargos_routing a, rubros b, monedas c, routings d where d.id_routing=" & RoutingValues(0,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id and a.id_routing=d.id_routing and d.borrado=false and a.inter_company=0"
        if RoutingValues(19,0) > 0 then '2016-03-29 se agrego este union de seguros
            QuerySelect = QuerySelect & " UNION select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' as nada, a.inter_company, a.prepaid from cargos_routing a, rubros b, monedas c, routings d where d.id_routing=" & RoutingValues(19,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id and a.id_routing=d.id_routing and d.borrado=false and a.inter_company=0"
            'QuerySelect = QuerySelect & " and d.activo=true and d.borrado=false and d.routing_fac=0 "
        end if

        if RoutingValues(21,0) > 0 then '2016-04-21 
            QuerySelect = QuerySelect & " UNION select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' as nada, a.inter_company, a.prepaid from cargos_routing a, rubros b, monedas c, routings d where d.id_routing=" & RoutingValues(21,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id and a.id_routing=d.id_routing and d.borrado=false and a.inter_company=0"            
        end if        

        if RoutingValues(22,0) > 0 then '2016-04-21
            QuerySelect = QuerySelect & " UNION select c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '' as nada, a.inter_company, a.prepaid from cargos_routing a, rubros b, monedas c, routings d where d.id_routing=" & RoutingValues(22,0) & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id and a.id_routing=d.id_routing and d.borrado=false and a.inter_company=0"
        end if 

        'QuerySelect = "SELECT simbolo, id_rubro, valor, local, desc_rubro_es, id_servicio, nada, inter_company, prepaid  FROM (" & QuerySelect & ") x ORDER BY local, simbolo, valor DESC"
        QuerySelect = "SELECT simbolo, id_rubro, valor, local, trim(desc_rubro_es), id_servicio, nada, inter_company, CASE prepaid WHEN 't' THEN '0' ELSE '1' END AS prepaid FROM (" & QuerySelect & ") x ORDER BY local, simbolo, valor DESC"

        'QuerySelect = "SELECT * FROM (" & QuerySelect & ") x ORDER BY local, simbolo, valor DESC"
        'response.write( "2.////////////////////////////////////<br>" & QuerySelect & "<br>" )
        Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			RoutingCharges = rs.GetRows
			CountRoutingCharges = rs.RecordCount-1
		end if
		
		'Obteniendo los nombres de los servicios del Routing
		for i=0 to CountRoutingCharges
			Set rs = Conn.Execute("select nombre_servicio from servicios where id_servicio=" & CheckNum(RoutingCharges(5,i)))
			if Not rs.EOF then
				RoutingCharges(6,i) = trim(rs(0))
			end if
		next
		CloseOBJs rs, Conn
		
		if AwbType = 1 then '////////////////////////////////////////////////EXPORT//////////////////////////////////////////%>
<SCRIPT LANGUAGE="JavaScript">
if (seguro_poliza == false) {
	top.opener.document.forms[0].NoOfPieces.value =  top.opener.document.forms[0].NoOfPieces.value + '<%=RoutingValues(6,0)%>\n';
	top.opener.document.forms[0].Weights.value =  top.opener.document.forms[0].Weights.value + '<%=RoutingValues(7,0)%>\n';
	top.opener.document.forms[0].ChargeableWeights.value =  top.opener.document.forms[0].ChargeableWeights.value + '<%=RoutingValues(7,0)%>\n';	
    top.opener.document.forms[0].WeightsSymbol.value =  top.opener.document.forms[0].WeightsSymbol.value + '<%=RoutingValues(14,0)%>';
	top.opener.document.forms[0].Commodities.value = top.opener.document.forms[0].Commodities.value + '<%=CommodityID%>\n';
	top.opener.document.forms[0].NatureQtyGoods.value =  top.opener.document.forms[0].NatureQtyGoods.value + '<%=Commodity%>\n';
	top.opener.SumVals(top.opener.document.forms[0].NoOfPieces, top.opener.document.forms[0].TotNoOfPieces);
	top.opener.SumVals(top.opener.document.forms[0].Weights, top.opener.document.forms[0].TotWeight);
	//top.opener.GetCommodityName(top.opener.document.forms[0].Commodities, top.opener.document.forms[0].NatureQtyGoods, top.opener.document.forms[0].Weights, top.opener.document.forms[0].CommoditiesTypes);
	if ((top.opener.document.forms[0].AirportDepID.value!=-1) && (top.opener.document.forms[0].AirportDesID.value!=-1)) {
		top.opener.CalcRate(top.opener.document.forms[0]);
	}

	var ItemCurrIDs = new Array();
	var ItemIDs = new Array();
	var ItemVals = new Array();
	var ItemLocs = new Array();
	var ItemNames = new Array();
	var ItemServIDs = new Array();
	var ItemServNames = new Array();

    var ItemPrePaid = new Array();

	var CarrierIDs = new Array();
    var CarrierNIDs = new Array();
	var CarrierNames = new Array();
    var CarrierVals = new Array();
	var CarrierHVals = new Array();
	var CarrierCurs = new Array();
	var CarrierTCurs = new Array();
    var CarrierPP = new Array();
	var	CarrierServNames = new Array();
	var	CarrierServIDs = new Array();

	var AgentIDs = new Array();
    var AgentNIDs = new Array();
	var AgentNames = new Array();
	var AgentVals = new Array();
	var AgentHVals = new Array();
	var AgentCurs = new Array();
	var AgentTCurs = new Array();
    var AgentPP = new Array();
	var	AgentServNames = new Array();
	var	AgentServIDs = new Array();

	for (i=0; i<=3; i++) {
		CarrierIDs[i] = "C"+(i+1);
        CarrierNIDs[i] = "NC"+(i+1);
		CarrierCurs[i] = "CC"+(i+1);
		CarrierTCurs[i] = "TCC"+(i+1);
        CarrierPP[i] = "TPC"+(i+1);
		CarrierHVals[i] = "VC"+(i+1);
		CarrierServNames[i] = "SVNC"+(i+1);
		CarrierServIDs[i] = "SVIC"+(i+1);
	}

	CarrierNames[0] = "AdditionalChargeName3";
	CarrierNames[1] = "AdditionalChargeName4";
	CarrierNames[2] = "AdditionalChargeName5";
	CarrierNames[3] = "AdditionalChargeName8";

	CarrierVals[0] = "AdditionalChargeVal3";
	CarrierVals[1] = "AdditionalChargeVal4";
	CarrierVals[2] = "AdditionalChargeVal5";
	CarrierVals[3] = "AdditionalChargeVal8";

	for (i=0; i<=10; i++) {
		AgentIDs[i] = "A"+(i+1);
        AgentNIDs[i] = "NA"+(i+1);
		AgentCurs[i] = "CA"+(i+1);
		AgentTCurs[i] = "TCA"+(i+1);
        AgentPP[i] = "TPA"+(i+1);
		AgentHVals[i] = "VA"+(i+1);
		AgentServNames[i] = "SVNA"+(i+1);
		AgentServIDs[i] = "SVIA"+(i+1);
	}

	AgentNames[0] = "AdditionalChargeName1";
	AgentNames[1] = "AdditionalChargeName2";
	AgentNames[2] = "AdditionalChargeName6";
	AgentNames[3] = "AdditionalChargeName7";
	AgentNames[4] = "AdditionalChargeName9";
	AgentNames[5] = "AdditionalChargeName10";
	AgentNames[6] = "AdditionalChargeName11";
	AgentNames[7] = "AdditionalChargeName12";
	AgentNames[8] = "AdditionalChargeName13";
	AgentNames[9] = "AdditionalChargeName14";
	AgentNames[10] = "AdditionalChargeName15";

	AgentVals[0] = "AdditionalChargeVal1";
	AgentVals[1] = "AdditionalChargeVal2";
	AgentVals[2] = "AdditionalChargeVal6";
	AgentVals[3] = "AdditionalChargeVal7";
	AgentVals[4] = "AdditionalChargeVal9";
	AgentVals[5] = "AdditionalChargeVal10";
	AgentVals[6] = "AdditionalChargeVal11";
	AgentVals[7] = "AdditionalChargeVal12";
	AgentVals[8] = "AdditionalChargeVal13";
	AgentVals[9] = "AdditionalChargeVal14";
	AgentVals[10] = "AdditionalChargeVal15";	

	<%for i=0 to CountRoutingCharges
		RoutingCharges(1,i) = CInt(RoutingCharges(1,i))
	%>
		ItemCurrIDs[<%=i%>] = "<%=RoutingCharges(0,i)%>";
		ItemIDs[<%=i%>] = <%=RoutingCharges(1,i)%>;
		ItemVals[<%=i%>] = <%=replace(RoutingCharges(2,i),",",".")%>;
		ItemLocs[<%=i%>] = <%=RoutingCharges(3,i)%>;
		ItemNames[<%=i%>] = "<%=RoutingCharges(4,i)%>";
		ItemServIDs[<%=i%>] = "<%=RoutingCharges(5,i)%>";
		ItemServNames[<%=i%>] = "<%=RoutingCharges(6,i)%>";
        ItemPrePaid[<%=i%>] = "<%=RoutingCharges(8,i)%>";
	<%
	select Case RoutingCharges(1,i)
	case 11 'Air Freight, solo se asigna la moneda y tipo de moneda, el valor se calcula en base a tarifas del sistema%>
		top.opener.document.forms[0].CAF.value = "<%=RoutingCharges(0,i)%>";
		top.opener.document.forms[0].TotCarrierRate.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].TotCarrierRate_Routing.value = "1";
		top.opener.document.forms[0].CarrierSubTot.value = <%=replace(RoutingCharges(2,i),",",".")%>;
        <%if CheckNum(RoutingValues(7,0)) <> 0 then %>
		    top.opener.document.forms[0].CarrierRates.value = <%=Round(CheckNum(replace(RoutingCharges(2,i),",","."))/CheckNum(RoutingValues(7,0)),2)%>;
        <%else %>
            alert("Aviso: El Routing que desea asignar no le colocaron Peso");
            top.opener.document.forms[0].CarrierRates.value = <%=Round(CheckNum(replace(RoutingCharges(2,i),",",".")),2)%>;
        <%end if %>
		top.opener.document.forms[0].TCAF.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPAF.value = <%=RoutingCharges(8,i)%>;
	<%case 12 'Fuel Surcharge%>
		top.opener.document.forms[0].CFS.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].FuelSurcharge.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].FuelSurcharge_Routing.value = "1";
		top.opener.document.forms[0].TCFS.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPFS.value = <%=RoutingCharges(8,i)%>;        
	<%case 13 'Security Charge%>
		top.opener.document.forms[0].CSF.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].SecurityFee.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].SecurityFee_Routing.value = "1";
		top.opener.document.forms[0].TCSF.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPSF.value = <%=RoutingCharges(8,i)%>;        
	<%case 14 'Custom Fee%>
		top.opener.document.forms[0].CCF.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].CustomFee.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].CustomFee_Routing.value = "1";
		top.opener.document.forms[0].TCCF.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPCF.value = <%=RoutingCharges(8,i)%>;        
	<%case 15 'Terminal Fee%>
		top.opener.document.forms[0].CTF.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].TerminalFee.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].TerminalFee_Routing.value = "1";
		top.opener.document.forms[0].TCTF.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPTF.value = <%=RoutingCharges(8,i)%>;        
	<%case 31 'Pick Up%>
		//top.opener.document.forms[0].CPU.value = "<%=RoutingCharges(0,i)%>";
	  	//top.opener.document.forms[0].PickUp.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        //top.opener.document.forms[0].PickUp_Routing.value = "1";
		//top.opener.document.forms[0].TCPU.value = <%=RoutingCharges(3,i)%>;
	<%case 38 'Sed (Sed Filling Fee)%>
		top.opener.document.forms[0].CFF.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].SedFilingFee.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].SedFilingFee_Routing.value = "1";
		top.opener.document.forms[0].TCFF.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPFF.value = <%=RoutingCharges(8,i)%>;        
	<%case 115 'Intermodal%>
		top.opener.document.forms[0].CIM.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].Intermodal.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].Intermodal_Routing.value = "1";
		top.opener.document.forms[0].TCIM.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPIM.value = <%=RoutingCharges(8,i)%>;        
	<%case 116 'PBA%>
		top.opener.document.forms[0].CPB.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].PBA.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].PBA_Routing.value = "1";
		top.opener.document.forms[0].TCPB.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPPB.value = <%=RoutingCharges(8,i)%>;        
	<%end select
	  next%>

    function Asignar(){ //export
	    var j = 0;
	    var k = 0;
	    var radio;
	    for (i=0; i<<%=i%>; i++) {


		    //if ((ItemIDs[i]!=11) && (ItemIDs[i]!=12) && (ItemIDs[i]!=13) && (ItemIDs[i]!=14) && (ItemIDs[i]!=15) && (ItemIDs[i]!=31) && (ItemIDs[i]!=38) && (ItemIDs[i]!=115) && (ItemIDs[i]!=116)) {
            // se quito el 31 2016-04-21 PickUp para que lo asigne donde seleccionen
            if ((ItemIDs[i]!=11) && (ItemIDs[i]!=12) && (ItemIDs[i]!=13) && (ItemIDs[i]!=14) && (ItemIDs[i]!=15) && (ItemIDs[i]!=38) && (ItemIDs[i]!=115) && (ItemIDs[i]!=116)) {
			    radio = document.frm.elements["R"+i];
			    if (radio[0].checked) { //Agente
				    if (j<=10) {

console.log('Export Agente');
console.log( AgentIDs[j] + ' ' + AgentNIDs[j] + ' ' + AgentNames[j] + ' ' + AgentVals[j] + ' ' + AgentHVals[j] + ' ' + AgentCurs[j] + ' ' + AgentTCurs[j] + ' ' + AgentPP[j] + ' ' + AgentServIDs[j] + ' ' + AgentServNames[j] );
console.log( ItemIDs[i] + ' ' + ItemNames[i] + ' ' + ItemNames[i] + ' ' + ItemVals[i] + ' ' + ItemVals[i] + ' ' + ItemCurrIDs[i] + ' ' + ItemLocs[i] + ' ' + ItemPrePaid[i] + ' ' + ItemServIDs[i] + ' ' + ItemServNames[i] );

					    top.opener.document.forms[0].elements[AgentIDs[j]].value = ItemIDs[i];
					    top.opener.document.forms[0].elements[AgentNIDs[j]].value = ItemNames[i];
                        top.opener.document.forms[0].elements[AgentNames[j]].value = ItemNames[i]; // + " (" + ItemIDs[i] + ")";
                        if (top.opener.document.forms[0].elements[AgentNames[j] + '_Routing'])
                        top.opener.document.forms[0].elements[AgentNames[j] + '_Routing'].value = "1";
					    top.opener.document.forms[0].elements[AgentVals[j]].value = ItemVals[i]; 
					    top.opener.document.forms[0].elements[AgentHVals[j]].value = ItemVals[i];
					    top.opener.document.forms[0].elements[AgentCurs[j]].value = ItemCurrIDs[i];
					    top.opener.document.forms[0].elements[AgentTCurs[j]].value = ItemLocs[i];
                        top.opener.document.forms[0].elements[AgentPP[j]].value = ItemPrePaid[i];
					    top.opener.document.forms[0].elements[AgentServIDs[j]].value = ItemServIDs[i];
					    top.opener.document.forms[0].elements[AgentServNames[j]].value = ItemServNames[i]; // + " (" + ItemServIDs[i] + ")";
					    j = j +	1;
				    } else {
					    alert("Solo Existen 11 casillas para ingresar rubros de Agente");
				    }
			    } else { //Transportista
				    if (k<=3) {

console.log('Export Transportista');
console.log( CarrierIDs[k] + ' ' + CarrierNIDs[k] + ' ' + CarrierNames[k] + ' ' + CarrierVals[k] + ' ' + CarrierHVals[k] + ' ' + CarrierCurs[k] + ' ' + CarrierTCurs[k] + ' ' + CarrierPP[k] + ' ' + CarrierServIDs[k] + ' ' + CarrierServNames[k] );
console.log( ItemIDs[i] + ' ' + ItemNames[i] + ' ' + ItemNames[i] + ' ' + ItemVals[i] + ' ' + ItemVals[i] + ' ' + ItemCurrIDs[i] + ' ' + ItemLocs[i] + ' ' + ItemPrePaid[i] + ' ' + ItemServIDs[i] + ' ' + ItemServNames[i] );

					    top.opener.document.forms[0].elements[CarrierIDs[k]].value = ItemIDs[i];
                        top.opener.document.forms[0].elements[CarrierNIDs[k]].value = ItemNames[i];
					    top.opener.document.forms[0].elements[CarrierNames[k]].value = ItemNames[i]; // + " (" + ItemIDs[i] + ")";
                        if (top.opener.document.forms[0].elements[CarrierNames[k] + '_Routing'])
                        top.opener.document.forms[0].elements[CarrierNames[k] + '_Routing'].value = "1";
					    top.opener.document.forms[0].elements[CarrierVals[k]].value = ItemVals[i]; 
					    top.opener.document.forms[0].elements[CarrierHVals[k]].value = ItemVals[i];
					    top.opener.document.forms[0].elements[CarrierCurs[k]].value = ItemCurrIDs[i];
					    top.opener.document.forms[0].elements[CarrierTCurs[k]].value = ItemLocs[i];
                        top.opener.document.forms[0].elements[CarrierPP[k]].value = ItemPrePaid[i];
					    top.opener.document.forms[0].elements[CarrierServIDs[k]].value = ItemServIDs[i];
					    top.opener.document.forms[0].elements[CarrierServNames[k]].value = ItemServNames[i]; // + " (" + ItemServIDs[i] + ")";
					    k = k +	1;
				    } else {
					    alert("Solo Existen 4 casillas para ingresar rubros de Transportista");
				    }
			    }
		    }
	    }
	    top.opener.SumOtherCharges(top.opener.document.forms[0]);top.opener.CalcTax(top.opener.document.forms[0]);top.opener.CalcTot(top.opener.document.forms[0]);
	    top.opener.document.forms[0].CarrierID.value = <%=RoutingValues(9,0)%>;
	    top.opener.document.forms[0].CallRouting.value = 1; //2018-04-19 solo para que cargue bien el routing cuando la guia ya esta guardada	
	    top.opener.document.forms[0].submit();
        top.close();
    }
}
</SCRIPT>
	<form name="frm" method="post">
	<table cellspacing=5 cellpadding=2 width="100%">
	<tr>
	<td class=titlelist><b>Servicio</b></td><td class=titlelist><b>Rubro</b></td><td class=titlelist><b>Para Agente</b></td><td class=titlelist><b>Para Transportista</b></td>
	</tr>
	<%for i=0 to CountRoutingCharges
		'if Not FRegExp("^11$|^12$|^13$|^14$|^15$|^31$|^38$|^115$|^116$",RoutingCharges(1,i),"",2) then

        ' & " " & replace(RoutingCharges(2,i),",",".")
	%>
	<tr>
	<td class=label><b><%=RoutingCharges(6,i) & " (" & RoutingCharges(5,i) & ")"%></b></td><td class=label><b><%=RoutingCharges(4,i) & " (" & RoutingCharges(1,i) & ")"%></b></td><td class=label align="center"><input type="radio" name="R<%=i%>" value="0"></td><td class=labe align="center"><input type="radio" name="R<%=i%>" value="1"></td>
	</tr>
	<%	'end if
	next%>
	<tr>
	<td class=label colspan="5" align="center"><INPUT name=enviar type=button onClick="JavaScript:Asignar();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></td>
	</tr>
	</table>
	</form>
<%		else    '///////////////////////////////////////////////IMPORT////////////////////////////////////////////////////// %>
<SCRIPT LANGUAGE="JavaScript">
if (seguro_poliza == false) {
	top.opener.document.forms[0].NoOfPieces.value =  top.opener.document.forms[0].NoOfPieces.value + '<%=RoutingValues(6,0)%>\n';
	top.opener.document.forms[0].Weights.value =  top.opener.document.forms[0].Weights.value + '<%=RoutingValues(7,0)%>\n';
	top.opener.document.forms[0].ChargeableWeights.value =  top.opener.document.forms[0].ChargeableWeights.value + '<%=RoutingValues(7,0)%>\n';
	top.opener.document.forms[0].WeightsSymbol.value =  top.opener.document.forms[0].WeightsSymbol.value + '<%=RoutingValues(14,0)%>';
	top.opener.document.forms[0].Commodities.value = top.opener.document.forms[0].Commodities.value + '<%=CommodityID%>\n';
	top.opener.document.forms[0].NatureQtyGoods.value =  top.opener.document.forms[0].NatureQtyGoods.value + '<%=Commodity%>\n';
	top.opener.SumVals(top.opener.document.forms[0].NoOfPieces, top.opener.document.forms[0].TotNoOfPieces);
	top.opener.SumVals(top.opener.document.forms[0].Weights, top.opener.document.forms[0].TotWeight);
	//top.opener.GetCommodityName(top.opener.document.forms[0].Commodities, top.opener.document.forms[0].NatureQtyGoods, top.opener.document.forms[0].Weights, top.opener.document.forms[0].CommoditiesTypes);

	var ItemCurrIDs = new Array();
	var ItemIDs = new Array();
	var ItemVals = new Array();
	var ItemLocs = new Array();
	var ItemNames = new Array();
	var ItemServIDs = new Array();
	var ItemServNames = new Array();
    var ItemPrePaid = new Array();
	var CarrierIDs = new Array();
    var CarrierNIDs = new Array();
	var CarrierNames = new Array();
	var CarrierVals = new Array();
	var CarrierHVals = new Array();
	var CarrierCurs = new Array();
	var CarrierTCurs = new Array();
    var CarrierPP = new Array();
	var	CarrierServIDs = new Array();
	var	CarrierServNames = new Array();
	var AgentIDs = new Array();
    var AgentNIDs = new Array();
	var AgentNames = new Array();
	var AgentVals = new Array();
	var AgentHVals = new Array();
	var AgentCurs = new Array();
	var AgentTCurs = new Array();
    var AgentPP = new Array();
	var	AgentServIDs = new Array();
	var	AgentServNames = new Array();
	var OtherIDs = new Array();
    var OtherNIDs = new Array();
	var OtherNames = new Array();
	var OtherVals = new Array();
	var OtherHVals = new Array();
	var OtherCurs = new Array();
	var OtherTCurs = new Array();
    var OtherPP = new Array();
	var	OtherServIDs = new Array();
	var	OtherServNames = new Array();

	for (i=0; i<=3; i++) {
		CarrierIDs[i] = "C"+(i+1);
        CarrierNIDs[i] = "NC"+(i+1);
		CarrierCurs[i] = "CC"+(i+1);
		CarrierTCurs[i] = "TCC"+(i+1);
        CarrierPP[i] = "TPC"+(i+1);
		CarrierHVals[i] = "VC"+(i+1);
		CarrierServNames[i] = "SVNC"+(i+1);
		CarrierServIDs[i] = "SVIC"+(i+1);
	}

	CarrierNames[0] = "AdditionalChargeName3";
	CarrierNames[1] = "AdditionalChargeName4";
	CarrierNames[2] = "AdditionalChargeName5";
	CarrierNames[3] = "AdditionalChargeName8";

	CarrierVals[0] = "AdditionalChargeVal3";
	CarrierVals[1] = "AdditionalChargeVal4";
	CarrierVals[2] = "AdditionalChargeVal5";
	CarrierVals[3] = "AdditionalChargeVal8";

	for (i=0; i<=5; i++) {
		OtherIDs[i] = "O"+(i+1);
        OtherNIDs[i] = "NO"+(i+1);
		OtherCurs[i] = "CO"+(i+1);
		OtherTCurs[i] = "TCO"+(i+1);
        OtherPP[i] = "TPO"+(i+1);
		OtherHVals[i] = "VO"+(i+1);
		OtherNames[i] = "OtherChargeName"+(i+1);
		OtherVals[i] = "OtherChargeVal"+(i+1);
		OtherServNames[i] = "SVNO"+(i+1);
		OtherServIDs[i] = "SVIO"+(i+1);
	}

	for (i=0; i<=10; i++) {
		AgentIDs[i] = "A"+(i+1);
        AgentNIDs[i] = "NA"+(i+1);
		AgentCurs[i] = "CA"+(i+1);
		AgentTCurs[i] = "TCA"+(i+1);
        AgentPP[i] = "TPA"+(i+1);
		AgentHVals[i] = "VA"+(i+1);
		AgentServNames[i] = "SVNA"+(i+1);
		AgentServIDs[i] = "SVIA"+(i+1);
	}

	AgentNames[0] = "AdditionalChargeName1";
	AgentNames[1] = "AdditionalChargeName2";
	AgentNames[2] = "AdditionalChargeName6";
	AgentNames[3] = "AdditionalChargeName7";
	AgentNames[4] = "AdditionalChargeName9";
	AgentNames[5] = "AdditionalChargeName10";
	AgentNames[6] = "AdditionalChargeName11";
	AgentNames[7] = "AdditionalChargeName12";
	AgentNames[8] = "AdditionalChargeName13";
	AgentNames[9] = "AdditionalChargeName14";
	AgentNames[10] = "AdditionalChargeName15";

	AgentVals[0] = "AdditionalChargeVal1";
	AgentVals[1] = "AdditionalChargeVal2";
	AgentVals[2] = "AdditionalChargeVal6";
	AgentVals[3] = "AdditionalChargeVal7";
	AgentVals[4] = "AdditionalChargeVal9";
	AgentVals[5] = "AdditionalChargeVal10";
	AgentVals[6] = "AdditionalChargeVal11";
	AgentVals[7] = "AdditionalChargeVal12";
	AgentVals[8] = "AdditionalChargeVal13";
	AgentVals[9] = "AdditionalChargeVal14";
	AgentVals[10] = "AdditionalChargeVal15";	

	<%for i=0 to CountRoutingCharges
		RoutingCharges(1,i) = CInt(RoutingCharges(1,i))
	%>
		ItemCurrIDs[<%=i%>] = "<%=RoutingCharges(0,i)%>";
		ItemIDs[<%=i%>] = <%=RoutingCharges(1,i)%>;
		ItemVals[<%=i%>] = <%=replace(RoutingCharges(2,i),",",".")%>;
		ItemLocs[<%=i%>] = <%=RoutingCharges(3,i)%>;
		ItemNames[<%=i%>] = "<%=RoutingCharges(4,i)%>";
		ItemServIDs[<%=i%>] = "<%=RoutingCharges(5,i)%>";
		ItemServNames[<%=i%>] = "<%=RoutingCharges(6,i)%>";
        ItemPrePaid[<%=i%>] = "<%=RoutingCharges(8,i)%>";                
	<%
	select Case RoutingCharges(1,i)
	case 11 'Air Freight, solo se asigna la moneda y tipo de moneda, el valor se calcula en base a tarifas del sistema%> 
		top.opener.document.forms[0].CAF.value = "<%=RoutingCharges(0,i)%>";
        top.opener.document.forms[0].TotCarrierRate.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].TotCarrierRate_Routing.value = "1";
		top.opener.document.forms[0].CarrierSubTot.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
		top.opener.document.forms[0].CarrierRates.value = "<%=Round(CheckNum(replace(RoutingCharges(2,i),",","."))/CheckNum(RoutingValues(7,0)),2)%>";
		top.opener.document.forms[0].TCAF.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPAF.value = <%=RoutingCharges(8,i)%>;        
	<%case 12 'Fuel Surcharge%>
		top.opener.document.forms[0].CFS.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].FuelSurcharge.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].FuelSurcharge_Routing.value = "1";
		top.opener.document.forms[0].TCFS.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPFS.value = <%=RoutingCharges(8,i)%>;        
	<%case 13 'Security Charge%>
		top.opener.document.forms[0].CSF.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].SecurityFee.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].SecurityFee_Routing.value = "1";
		top.opener.document.forms[0].TCSF.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPSF.value = <%=RoutingCharges(8,i)%>;        
	<%case 31 'Pick Up%>
		//top.opener.document.forms[0].CPU.value = "<%=RoutingCharges(0,i)%>";
	  	//top.opener.document.forms[0].PickUp.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        //top.opener.document.forms[0].PickUp_Routing.value = "1";
		//top.opener.document.forms[0].TCPU.value = <%=RoutingCharges(3,i)%>;
	<%case 38 'Sed (Sed Filling Fee)%>
		top.opener.document.forms[0].CFF.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].SedFilingFee.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].SedFilingFee_Routing.value = "1";
		top.opener.document.forms[0].TCFF.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPFF.value = <%=RoutingCharges(8,i)%>;        
	<%case 115 'Intermodal%>
		top.opener.document.forms[0].CIM.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].Intermodal.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].Intermodal_Routing.value = "1";
		top.opener.document.forms[0].TCIM.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPIM.value = <%=RoutingCharges(8,i)%>;        
	<%case 116 'PBA%>
		top.opener.document.forms[0].CPB.value = "<%=RoutingCharges(0,i)%>";
	  	top.opener.document.forms[0].PBA.value = "<%=replace(RoutingCharges(2,i),",",".")%>";
        top.opener.document.forms[0].PBA_Routing.value = "1";
		top.opener.document.forms[0].TCPB.value = <%=RoutingCharges(3,i)%>;
        top.opener.document.forms[0].TPPB.value = <%=RoutingCharges(8,i)%>;        
	<%end select
	next%>

    function Asignar(){  //import
	    var j = 0;
	    var k = 0;
	    var l = 0;
	    var radio;

	    for (i=0; i<<%=i%>; i++) {		    

            //alert('IMPORT : ' + ItemIDs[i]);

            //if ((ItemIDs[i]!=11) && (ItemIDs[i]!=12) && (ItemIDs[i]!=13) && (ItemIDs[i]!=31) && (ItemIDs[i]!=38) && (ItemIDs[i]!=115) && (ItemIDs[i]!=116)) {
            if ((ItemIDs[i]!=11) && (ItemIDs[i]!=12) && (ItemIDs[i]!=13) && (ItemIDs[i]!=38) && (ItemIDs[i]!=115) && (ItemIDs[i]!=116)) {			    
            // se quito el 31 2016-04-21 PickUp para que lo asigne donde seleccionen
                
                radio = document.frm.elements["R"+i];
			    if (radio[0].checked) { //Agente
				    if (j<=10) {

					    top.opener.document.forms[0].elements[AgentIDs[j]].value = ItemIDs[i];
					    top.opener.document.forms[0].elements[AgentNIDs[j]].value = ItemNames[i];
                        top.opener.document.forms[0].elements[AgentNames[j]].value = ItemNames[i]; // + " (" + ItemIDs[i] + ")";
                        if (top.opener.document.forms[0].elements[AgentNames[j] + '_Routing'])
                        top.opener.document.forms[0].elements[AgentNames[j] + '_Routing'].value = "1";
					    top.opener.document.forms[0].elements[AgentVals[j]].value = ItemVals[i]; 
					    top.opener.document.forms[0].elements[AgentHVals[j]].value = ItemVals[i];
					    top.opener.document.forms[0].elements[AgentCurs[j]].value = ItemCurrIDs[i];
					    top.opener.document.forms[0].elements[AgentTCurs[j]].value = ItemLocs[i];
                        top.opener.document.forms[0].elements[AgentPP[j]].value = ItemPrePaid[i];
					    top.opener.document.forms[0].elements[AgentServIDs[j]].value = ItemServIDs[i];
					    top.opener.document.forms[0].elements[AgentServNames[j]].value = ItemServNames[i]; // + " (" + ItemServIDs[i] + ")";
					    j = j +	1;
				    } else {
					    alert("Solo Existen 11 casillas para ingresar rubros de Agente");
				    }
			    } else { //Transportista
				    if (radio[1].checked) { 
					    if (k<=3) {
						    top.opener.document.forms[0].elements[CarrierIDs[k]].value = ItemIDs[i];
                            top.opener.document.forms[0].elements[CarrierNIDs[k]].value = ItemNames[i];
						    top.opener.document.forms[0].elements[CarrierNames[k]].value = ItemNames[i]; // + " (" + ItemIDs[i] + ")";
                            if (top.opener.document.forms[0].elements[CarrierNames[k] + '_Routing'])
                            top.opener.document.forms[0].elements[CarrierNames[k] + '_Routing'].value = "1";
						    top.opener.document.forms[0].elements[CarrierVals[k]].value = ItemVals[i]; 
						    top.opener.document.forms[0].elements[CarrierHVals[k]].value = ItemVals[i];
						    top.opener.document.forms[0].elements[CarrierCurs[k]].value = ItemCurrIDs[i];
						    top.opener.document.forms[0].elements[CarrierTCurs[k]].value = ItemLocs[i];
                            top.opener.document.forms[0].elements[CarrierPP[k]].value = ItemPrePaid[i];
						    top.opener.document.forms[0].elements[CarrierServIDs[k]].value = ItemServIDs[i];
						    top.opener.document.forms[0].elements[CarrierServNames[k]].value = ItemServNames[i]; // + " (" + ItemServIDs[i] + ")";
						    k = k +	1;
					    } else {
						    alert("Solo Existen 4 casillas para ingresar rubros de Transportista");
					    }
				    } else {
					    if (l<=5) {


						    top.opener.document.forms[0].elements[OtherIDs[l]].value = ItemIDs[i];
						    top.opener.document.forms[0].elements[OtherNIDs[l]].value = ItemNames[i];
                            top.opener.document.forms[0].elements[OtherNames[l]].value = ItemNames[i]; // + " (" + ItemIDs[i] + ")";
                            if (top.opener.document.forms[0].elements[OtherNames[l] + '_Routing'])
                            top.opener.document.forms[0].elements[OtherNames[l] + '_Routing'].value = "1";
						    top.opener.document.forms[0].elements[OtherVals[l]].value = ItemVals[i]; 
						    top.opener.document.forms[0].elements[OtherHVals[l]].value = ItemVals[i];
						    top.opener.document.forms[0].elements[OtherCurs[l]].value = ItemCurrIDs[i];
						    top.opener.document.forms[0].elements[OtherTCurs[l]].value = ItemLocs[i];
                            top.opener.document.forms[0].elements[OtherPP[l]].value = ItemPrePaid[i];
						    top.opener.document.forms[0].elements[OtherServIDs[l]].value = ItemServIDs[i];
						    top.opener.document.forms[0].elements[OtherServNames[l]].value = ItemServNames[i]; // + " (" + ItemServIDs[i] + ")";
						    l = l +	1;
					    } else {
						    alert("Solo Existen 6 casillas para ingresar Otros rubros");
					    }
				    }
			    }
		    }
	    }

        top.opener.SumOtherCharges(top.opener.document.forms[0]);
	    top.opener.document.forms[0].CarrierID.value = <%=RoutingValues(9,0)%>;	    
	    top.opener.document.forms[0].CallRouting.value = 1; //2018-04-19 solo para que cargue bien el routing cuando la guia ya esta guardada	
	    top.opener.document.forms[0].submit(); //2016-03-29

        top.close();

    }
}
</SCRIPT>
	<form name="frm" method="post">
	<table cellspacing=5 cellpadding=2 width="100%">
	<tr>
	<td class=titlelist><b>Servicio</b></td><td class=titlelist><b>Rubro</b></td><td class=titlelist><b>Para Agente</b></td><td class=titlelist><b>Para Transportista</b></td><td class=titlelist><b>Otros</b></td>
	</tr>
	<%for i=0 to CountRoutingCharges
		'if Not FRegExp("^11$|^12$|^13$|^31$|^38$|^115$|^116$",RoutingCharges(1,i),"",2) then
	%>
	<tr>
	<td class=label><b><%=RoutingCharges(6,i) & " (" & RoutingCharges(5,i) & ")"%></b></td><td class=label><b><%=RoutingCharges(4,i) & " (" & RoutingCharges(1,i) & ")"%></b></td><td class=label align="center"><input type="radio" name="R<%=i%>" value="0"></td><td class=labe align="center"><input type="radio" name="R<%=i%>" value="1"></td><td class=labe align="center"><input type="radio" name="R<%=i%>" value="2"></td>
	</tr>
	<%	'end if
	next%>
	<tr>
	<td class=label colspan="5" align="center"><INPUT name=enviar type=button onClick="JavaScript:Asignar();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></td>
	</tr>
	</table>
	</form>
<%		end if '//////////////////////////////////////// fin if import export /////////////////////////////////////////
		Set RoutingValues = Nothing
		Set RoutingCharges = Nothing
	end if
%>
</BODY>
</HTML>