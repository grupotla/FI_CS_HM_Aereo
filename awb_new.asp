<%
	AwbType = CheckNum(Request("AT"))
	
	ArrAwbType = Array("","EXPORT","IMPORT")


    'response.write  "(BtnReplica2=" & Request("BtnReplica2") & ")"

    AwbNumber = trim(Request("AwbNumber2"))
    BtnReplica2 = Request("BtnReplica2")
    BtnMaster2 = Request("BtnMaster2")
    AirportDesID2 = Request("AirportDesID2")
    AirportDepID2 = Request("AirportDepID2")
    Transportista2 = Request("Transportista2")
    Routing2 = Request("Routing2")    
    BtnHouse2 = Request("BtnHouse2")
    HAWBNumber2 = trim(Request("HAWBNumber2"))

    HAWBNumberTipo2 = Request("HAWBNumberTipo2")
    Piezas2 = CheckNum(Request("Piezas2"))    
    Peso2 = Request("Peso2")    
    TipoCarga2 = Request("TipoCarga2")       

    Country3 = Request("Country3")

    'Country2 = Request("Country2")

	couStr = Replace(Session("Countries"),"'","")
	couStr = Replace(couStr,"(","")
	couStr = Replace(couStr,")","")
	couStr = Split(couStr,",")



    if Request("Country2") <> "" then
        Country2 = Request("Country2")
    else

        if uBound(couStr) = 0 then
            Country2 = couStr(0) 
        end if

    end if




    if Request("completar") <> "" then '2018-06-25

        On Error Resume Next
        
            item0 = "Directo"

            if BtnReplica2 = "Consolidado" or BtnReplica2 = "Master-Hija" or BtnReplica2 = "Master-Master-Hija"  then 'crea la master        
        
                item0 = "Consolidado"

                'crea la master consolidado
                AWBiID2 = InsertGuia(Conn, rs, BtnReplica2, AwbType, AWBNumber, "", Country2, Request("Piezas2"), Request("Peso2"), Request("Transportista2"), Request("AirportDepID2"), Request("AirportDesID2"), Request("iAirportFromCode"), Request("iAirportToCode"), IIf(AwbType = 1,", '" & item0 & "'",""), item0, "Master", Request("TipoCarga2"))

            end if

            'crea la hija si es consolidado y master si es directa
            AWBiID2 = InsertGuia(Conn, rs, BtnReplica2, AwbType, AWBNumber, HAWBNumber2, Country2, Request("Piezas2"), Request("Peso2"), Request("Transportista2"), Request("AirportDepID2"), Request("AirportDesID2"), Request("iAirportFromCode"), Request("iAirportToCode"), IIf(AwbType = 1,", '" & item0 & "'",""), item0, Iif(item0 = "Consolidado", "Hija", "Master"), Request("TipoCarga2"))
            
            if AWBiID2 > 0 then     

                QuerySelect = "SELECT AwbID, CreatedDate, CreatedTime FROM " & IIf(AwbType = 1,"Awb","Awbi") & " WHERE AwbID = " & AWBiID2
                'response.write (QuerySelect & "<BR>")
                Set rs = Conn.Execute(QuerySelect) 'SI NO ES INSERT LEE LA REPLICA DEL IMPORT
                if Not rs.EOF then                                                        
                    QuerySelect = "InsertData.asp?OID=" & rs(0) & "&GID=1&CD=" & rs(1) & "&CT=" & rs(2) & "&AT=" & AwbType  & "&awb_frame2=2"            
                    'response.write (QuerySelect & "<BR>")
                    Response.Redirect QuerySelect
                    response.end
                end if        

            end if
        
            response.write "Problemas para almacenar guia<br>"
        
        If Err.Number <> 0 Then
          response.write  Err.Number & " - " & Err.Description
        End If

        response.end

    end if


	Select Case Request("BtnBorra")
	Case "Borra Pais"
        Country2 = ""
        Routing2 = ""
        Transportista2 = ""
        AirportDepID2 = ""
        AirportDesID2 = ""
        BtnMaster2 = ""
        AwbNumber = ""
        BtnReplica2 = ""
        HAWBNumberTipo2 = ""
        HAWBNumber2 = ""
        Piezas2 = 0


        'BtnReplica2 = ""
	Case "Borra Tipo AWB"            
        Routing2 = ""
        Transportista2 = ""
        AirportDepID2 = ""
        AirportDesID2 = ""
        BtnMaster2 = ""
        AwbNumber = ""
        BtnReplica2 = ""
        HAWBNumberTipo2 = ""
        HAWBNumber2 = ""
        Piezas2 = 0


	Case "Borra Routing"
        Routing2 = ""
        Transportista2 = ""
        AwbNumber = ""
        AirportDepID2 = ""
        AirportDesID2 = ""
        BtnMaster2 = ""
        HAWBNumberTipo2 = ""
        HAWBNumber2 = ""
        Piezas2 = 0

	Case "Borra Linea Aerea"
        Transportista2 = ""
        AirportDepID2 = ""
        AirportDesID2 = ""
        BtnMaster2 = ""
        AwbNumber = ""
        HAWBNumberTipo2 = ""
        HAWBNumber2 = ""
        Piezas2 = 0


	Case "Borra Aeropuerto Salida", "Borra Aeropuerto Destino"

        if AwbType = 1 then 'export

            if Request("BtnBorra") = "Borra Aeropuerto Salida" then
                AirportDepID2 = ""
            else
                AirportDesID2 = ""
            end if     
        else 'import
            if Request("BtnBorra") = "Borra Aeropuerto Destino" then
                AirportDesID2 = ""
            else
                AirportDepID2 = ""
            end if
        end if

        BtnMaster2 = ""
        AwbNumber = ""
        HAWBNumberTipo2 = ""
        HAWBNumber2 = ""
        Piezas2 = 0

	Case "Borra AwbNumber"
        'BtnReplica2 = ""
        AwbNumber = ""
        BtnMaster2 = ""        
        BtnHouse2 = ""
        HAWBNumber2 = ""
        HAWBNumberTipo2 = ""
        HAWBNumber2 = ""
        Piezas2 = 0
    
	Case "Borra BtnHouse"
        BtnHouse2 = ""
        HAWBNumber2 = ""    
        HAWBNumberTipo2 = ""
        HAWBNumber2 = ""
        Piezas2 = 0

	Case "Borra HAWBNumber"
        HAWBNumber2 = ""    
        HAWBNumberTipo2 = ""
        Piezas2 = 0

    Case "Borra Tipo HAWBNumber"
        HAWBNumberTipo2 = ""        
        Piezas2 = 0

    Case "Borra Piezas"
        Piezas2 = 0
        TipoCarga2 = ""
        Peso2 = 0
    
	Case "Borra TipoCarga"
        TipoCarga2 = ""
        Peso2 = 0
 
    Case "Borra Peso"
        Peso2 = 0


    End Select

    
    if Routing2 <> "" and Routing2 <> "NINGUNO" then

		QuerySelect = 	"select * from routings where id_routing = " & Request("Routing2")
		OpenConn2 Conn2
        'response.write QuerySelect & "<br>"
		Set rst = Conn2.Execute(QuerySelect)
		If Not rst.EOF Then

            if CheckNum(rst("carrier_id")) > 0 then
                Transportista2 = rst("carrier_id")
                AirportDepID2 = rst("airportid_embarque")
                AirportDesID2 = rst("airportid_desembarque")
            else
                Routing2 = ""
                response.write "Routing seleccionado " & rst("routing") & " no tiene datos basicos"
            end if

		End If
        CloseOBJs rst, Conn2

            'response.write "(" & Transportista2 & ")(" & AirportDepID2 & ")(" & AirportDesID2 & ")"

    end if



%>   
    
<script type="text/javascript">

    function move() {
        window.location = '#';
        //document.awb_frame.style.display = "none";
        document.getElementById('myBar').style.display = "block";
        var elem = document.getElementById("myBar");
        var width = 10;
        var id = setInterval(frame, 65);
        function frame() {
            if (width >= 100) {
                clearInterval(id);
            } else {
                width++;
                elem.style.width = width + '%';
                elem.innerHTML = width * 1 + '%';
            }
        }
    }

</script>

<style>
#myBar {
    width: 10%;
    height: 15px;
    background-color: #4CAF50;
    text-align: center; /* To center it horizontally (if you want) */
    line-height: 15px; /* To center it vertically */
    color: white;
    font-weight: bold;
    display: none;
}       
</style>


    <div id="myProgress">
        <div id="myBar">10%</div>
    </div>

<form name="awb_frame" action="InsertData.asp" onsubmit="move();" >

<input type="hidden" name="GID" value="1" />     
<input type="hidden" name="AT" value="<%=AwbType%>" />



<table width=70% align=center border=0 id="awb_frame">
<tr>
	<td align=center colspan=4>
        <h1>NUEVA HOUSE &nbsp::&nbsp::&nbsp; <%response.write ArrAwbType(AwbType)%> &nbsp::&nbsp::&nbsp; <%=Request("Country2")%> </h1>
    </td>
</tr>


<% If InStr(1, Session("Pricing"), "'" & Country2 & "'") > 0 and Routing2 <> "" Then 

        OpenConn3 Conn

        QuerySelect = "SELECT tpe_id_persona_fk FROM ti_pricing_list " & _ 
        "INNER JOIN ti_pricing_entidad ON tpe_tps_fk = '1' AND tpe_tpl_fk = tpl_pk AND tpe_tipo_persona_fk = 4 " & _ 
        "WHERE tpl_pais_fk = '" & Country2 & "'  AND tpl_transporte_fk = '1' AND tpl_movimiento = '" & ArrAwbType(AwbType) & "' AND tpl_tps_fk = '1'  AND tpl_tipo = 'COSTO' " 
        'response.write QuerySelect & "<br>"
        Set rs = Conn.Execute(QuerySelect)														

        item0  = ""

		If Not rs.EOF Then    
            do while Not rs.EOF
                
                if item0 <> "" then 
                    item0 = item0 & ","
                end if
                item0 = item0 & CheckNum(rs("tpe_id_persona_fk"))
                
				rs.MoveNext
			loop

		End If

		CloseOBJ rs 
		CloseOBJ Conn 

        if item0 <> "" then
            item0 = " AND CarrierID IN (" & item0 & ")"
        end if

        'response.write item0 & "<br>"
            
End If %>



<!--  ////////////////////////////////////// PAIS ///////////////////////////////////////////    -->

<tr>	
	<th style="width:25%">Pais</th>
    <td style="width:25%">    
    <%  

    QuerySelect = "SELECT pais_iso, nombre_empresa, vencimiento, COALESCE(vencimiento - CURRENT_DATE,0) as dias FROM empresas WHERE activo = true AND pais_iso IN " & Session("Countries") & " AND COALESCE(vencimiento - CURRENT_DATE,0) >= 0 ORDER BY pais_iso" 
    'response.write QuerySelect & "<br>" & uBound(couStr)

	if Country2 = "" then 
		%>
		<select name=Country2 onChange="document.awb_frame.submit();move();">
		<option value="-1">Seleccionar</option>
		<%
		'for each couStr2 in couStr		
		'	response.write("<option value='" & couStr2 & "'>" & couStr2 & "</option>")		
		'next

        if uBound(couStr) > 0 then

            OpenConn2 Conn2
	        Set rs = Conn2.Execute(QuerySelect)
	        if Not rs.EOF then
		        do while Not rs.EOF			
			        response.write("<option value='" & rs(0) & "'>" & rs(0) & " - " & rs(1) & "</option>")				
			        rs.MoveNext
		        loop
	        end if
            CloseOBJs rs, Conn2	
        end if

		%>	
		</select>		

		<input type="hidden" name="awb_frame2" value="3" />				
		
	<% else %>
		
		<input type="hidden" name="Country2" value="<%=Country2%>" />								
		<input type="text" name="" value="<%=Country2%>" disabled size=7 /> 
        
        <button type="submit" name="BtnBorra" value="Borra Pais" class="Boton2 cOrag" title="Edita Pais"><img src="img/glyphicons_150_edit.png" /></button>
        
	
	<% end if %>
	    
	</td>

<!--  ////////////////////////////////////// TIPO AWB ///////////////////////////////////////////    -->

	<th style="width:15%">Tipo Awb</th>
    <td style="width:35%">

                <%'="(" & Session("Pricing") & ")(" & Country2 & ")"%>

        <% if Country2 <> "" then %>
						
			<% if BtnReplica2 = "" then %>
			
				<input type="hidden" name="awb_frame2" value="3" />		


                <!--
                <select name="BtnReplica2" onchange="document.awb_frame.submit();move();">
                <option value="">-- Seleccione --</option>
                <option value="Consolidado">Consolidado</option>
                <option value="Directo">Directo</option>
                <option value="Master-Hija">Master-Hija</option>
                <option value="Hija-Directa">Hija-Directa</option>
                <option value="Master-Master-Hija">Master-Master-Hija</option>
                </select>
                -->

                <select name="BtnReplica2" onchange="document.awb_frame.submit();move();">
                <option value="">-- Seleccione --</option>
                <option value="<%=Iif(InStr(1, Session("Pricing"), "'" & Country2 & "'") = 0,"Consolidado","Master-Hija")%>"><%=Iif(InStr(1, Session("Pricing"), "'" & Country2 & "'") = 0,"Consolidado","Master-Hija")%></option>
                <option value="<%=Iif(InStr(1, Session("Pricing"), "'" & Country2 & "'") = 0,"Directo","Hija-Directa")%>"><%=Iif(InStr(1, Session("Pricing"), "'" & Country2 & "'") = 0,"Directo","Hija-Directa")%></option>

                <% If InStr(1, Session("Pricing"), "'" & Country2 & "'") > 0 Then %>

                    <option value="Master-Master-Hija">Master-Master-Hija</option>
            
                <% End If %>

                </select>


			
			<% else %>
				
				<input type="text" name="" value="<%                
                select case Iif(InStr(1, Session("Pricing"), "'" & Country2 & "'") = 0, BtnReplica2, "")
                case "Master-Hija"
                    response.write "Consolidado"
                case "Hija-Directa"
                    response.write "Directo"
                case "Master-Master-Hija"
                    response.write "Consolidado2"
                case else
                    response.write BtnReplica2
                end select %>                                
                " disabled size=10 />
				<input type="hidden" name="BtnReplica2" value="<%=BtnReplica2%>" />                                
                <button type="submit" name="BtnBorra" value="Borra Tipo AWB" class="Boton2 cOrag" title="Edita Tipo AWB"><img src="img/glyphicons_150_edit.png" /></button>

			<% end if %>
		
		<% end if %>		
    </td>	
</tr>





<% if 1 = 1 then %>

<!--  ////////////////////////////////////// routing ///////////////////////////////////////////    -->

<tr>	
	<th style="width:25%">Routing</th>
	<td colspan=3 align=leftstyle="width:75%">    
    <%  	
	if BtnReplica2 <> "" then 

        QuerySelect = 	"select a.id_routing, a.routing, a.carrier_id, a.airportid_embarque, a.airportid_desembarque from routings a left join seguros b on (a.id_routing=b.id_routing and b.anulado=false) "
		 
        QuerySelect = QuerySelect & " where a.id_transporte=1 and a.id_routing_type=2 and (a.activo=true or a.seguro=true) and a.borrado=false " 'Transporte Aereo, Routings Tipo Internos
         
        QuerySelect = QuerySelect & " and (a.bl_id = 0 or a.bl_id IS NULL) and a.id_cliente_order IS NOT NULL and carrier_id IS NOT NULL AND a.fecha > current_date - interval '6 month' "  
		 
        if AWBType = 1 then 'Export
            QuerySelect = QuerySelect & " AND a.import_export=false "
        else 'Import
            QuerySelect = QuerySelect & " AND a.import_export=true "
        end if
			 
        if AWBType = 1 then 'Export solo cuando es export realiza el union 			
            QuerySelect = QuerySelect & " UNION " & Replace(QuerySelect,"a.import_export=false","a.import_export=true") 
										
            QuerySelect = QuerySelect & " and a.id_pais_origen IN " & Session("Countries")			
        end if    
	
        QuerySelect = QuerySelect & " ORDER BY routing desc LIMIT 500"

        CountList4Values = -1
        
        'response.write QuerySelect & "<br>"
        OpenConn2 Conn2        		
        Set rst = Conn2.Execute(QuerySelect)
		If Not rst.EOF Then
    		aList4Values = rst.GetRows
        	CountList4Values = rst.RecordCount
		End If
        CloseOBJs rst, Conn2
         
	    if Routing2 = "" then 
		%>		    
            <input type="hidden" name="awb_frame2" value="3" />				
		    <select id="Routing2" name="Routing2" onChange="document.awb_frame.submit();move();" >
		    <option value="-1">Seleccionar</option>

            <% 'if AwbType = 2 or (AwbType = 1 and BtnReplica2 = "Directo") then 'TEMPORALMENTE DESAHABILITADO 2018-07-04 %>
		    <option value="NINGUNO">- No Routing -</option>
    	    <% 'end if %>

		    <%
			    For i = 0 To CountList4Values-1
		    %>
		    <option value="<%=aList4Values(0,i)%>"><%=aList4Values(1,i)%></option>
		    <%
			    Next
		    %>
		    </select>
	    <% else %>
            <%
				CarrierName = "*"
				For i = 0 To CountList4Values-1
					if CheckNum(aList4Values(0,i)) = CheckNum(Routing2) then
						CarrierName = aList4Values(1,i)
					end if				
				Next
            %>
		    <input type="hidden" name="Routing2" value="<%=Routing2%>" />		
            <input type="hidden" name="RoutingNo2" value="<%=CarrierName%>" />		            						
		    <input type="text" name="" value="<%=Routing2%>" disabled size=7/> 
            <% if CarrierName <> "*" then %>
            <input type="text" name="" value="<%=CarrierName%>" disabled size=50/> 
            <% end if %>        
            <button type="submit" name="BtnBorra" value="Borra Routing" class="Boton2 cOrag" title="Edita Routing"><img src="img/glyphicons_150_edit.png" /></button>        
	    <% end if %>
	<% end if %>	    
	</td>
</tr>

<% else 

    if BtnReplica2 <> "" then
        Routing2 = "NINGUNO"
    end if

 end if %>

<!--  ////////////////////////////////////// LINEA AEREA ///////////////////////////////////////////    -->

<%
OpenConn Conn
%>

<tr>
	<th style="width:25%">Linea Aerea</th>
	<td colspan=3 align=leftstyle="width:75%" style="width:75%">
		<%

		if Routing2 <> "" then 

            if Routing2 = "NINGUNO" then                 
                QuerySelect = "select CarrierID, UPPER(Name), Countries from Carriers where Expired = 0 " & item0 & " and Countries = '" & Country2 & "' order by Name, Countries"
            else
                QuerySelect = "select CarrierID, UPPER(Name), Countries from Carriers where 1 = 1 " & item0 & " order by Name, Countries"
            end if
            
            'response.write QuerySelect & "<br>"
			Set rst = Conn.Execute(QuerySelect)
			If Not rst.EOF Then
				aList1Values = rst.GetRows
				CountList1Values = rst.RecordCount
			End If
			CloseOBJ rst
				
			if Transportista2 = "" then 
				
				%>

				<select name="Transportista2" onChange="document.awb_frame.submit();move();">
				<option value="-1">Seleccionar</option>
				<%
					For i = 0 To CountList1Values-1
				%>
				<option value="<%=aList1Values(0,i)%>"><%=aList1Values(1,i) & " - " & aList1Values(2,i)  & " - " & aList1Values(0,i) %></option>
				<%
					Next
				%>
				</select>
				
				<input type="hidden" name="awb_frame2" value="3" />		
                <input type="hidden" name="linea" value="1" />		
				
			<% else %>
	
				<%
					CarrierName = "*"
					For i = 0 To CountList1Values-1
						if CheckNum(aList1Values(0,i)) = CheckNum(Transportista2) then
							CarrierName = aList1Values(2,i) & " - " & aList1Values(1,i)			
                            'Country3 = aList1Values(2,i)
						end if				
					Next

                    'if Country2 = "GT" then
                    'if Request("linea") = "1" and AwbType = 1 then 'EXPORT  
				    '    AWBNumber = NextAWBNumber(Conn, AwbType, Transportista2, "Nuevo") 
                    'end if
                    'end if
				%>	
				
				<input type="hidden" name="Transportista2" value="<%=Transportista2%>" />
				
				<input type="text" name="" value="<%=Transportista2%>" disabled size=7 />
				<input type="text" name="" value="<%=CarrierName%>" disabled size=50 />

                <input type="hidden" name="Country3" value="<%=Country3%>" />

                <button type="submit" name="BtnBorra" value="Borra Linea Aerea" class="Boton2 cOrag" title="Edita Linea Aerea"><img src="img/glyphicons_150_edit.png" /></button>

				
			<% end if %>
			
		<% 
        
        end if 
        
        
        %>

	</td>
	
</tr>




<% if AwbType = 1 then 'EXPORT %>

<!--  ////////////////////////////////////// AEROPUERTO SALIDA EXPORT ///////////////////////////////////////    -->

<tr>
	<th style="width:25%">Aeropuerto Salida</th>
	<td colspan=3 align=leftstyle="width:75%" style="width:75%">	
		<% if Transportista2 <> "" then %>

			<% if Transportista2 <> "" then 

                    result = WsAirportDisplay(Conn, rs, "EXPORT", "SALIDA", AirportDepID2, Country2, 0, Request("Transportista2"), 0)

                    Airports2 = "0"
			
            	end if
			%>
		
			<% if AirportDepID2 = "" then %>
			
				<input type="hidden" name="awb_frame2" value="3" />					
						
				<select class="style10" name="AirportDepID2" id="Aeropuerto Salida" style="width:200px" onChange="document.awb_frame.submit();move();">
				<%=result(0)%>
				</select>
			
			<% else %>

				<%
                    iAirportFromCode = result(1)
					CarrierName = result(2)	
				%>	
				
				<input type="hidden" name="AirportDepID2" value="<%=AirportDepID2%>"/>
                <input type="hidden" name="iAirportFromCode" value="<%=iAirportFromCode%>"/>
				
				<input type="text" name="" value="<%=AirportDepID2%>" disabled size=7/>
				<input type="text" name="" value="<%=CarrierName%>" disabled size=50 />


                <button type="submit" name="BtnBorra" value="Borra Aeropuerto Salida" class="Boton2 cOrag" title="Edita Aeropuerto Salida"><img src="img/glyphicons_150_edit.png" /></button>
			
			<% end if %>

		<% end if %>
	</td>
</tr>	


<!--  ////////////////////////////////////// AEROPUERTO DESTINO EXPORT //////////////////////////////////////    -->
<tr>
	<th style="width:25%">Aeropuerto Destino</th>
	<td colspan=3 align=leftstyle="width:75%" style="width:75%">	
	
		<% if AirportDepID2 <> "" then 	
			
            result = WsAirportDisplay(Conn, rs, "EXPORT", "DESTINO", AirportDesID2, Country2, Request("Routing2"), Request("Transportista2"), Request("AirportDepID2"))

			if AirportDesID2 = "" then %>
			
				<input type="hidden" name="awb_frame2" value="3" />					
			
				<select class="style10" name="AirportDesID2" onChange="document.awb_frame.submit();move();" id="Aeropuerto Destino">
				<%=result(0)%>
				</select>	
			
			<% else %>
								
				<%
                    iAirportToCode = result(1)
					CarrierName = result(2)						

                    Airports2 = "1"
				%>	
				
				<input type="hidden" name="AirportDesID2" value="<%=AirportDesID2%>" />
                <input type="hidden" name="iAirportToCode" value="<%=iAirportToCode%>" />
				
				<input type="text" name="" value="<%=AirportDesID2%>" disabled size=7/>
				<input type="text" name="" value="<%=CarrierName%>" disabled size=50 />

                <button type="submit" name="BtnBorra" value="Borra Aeropuerto Destino" class="Boton2 cOrag" title="Edita Aeropuerto Destino"><img src="img/glyphicons_150_edit.png" /></button>				
			
			<% end if %>

		<% end if %>
	</td>
	
</tr>

<% end if %>









<% if AwbType = 2 then 'IMPORT %>




<!--  ////////////////////////////////////// AEROPUERTO SALIDA IMPORT ///////////////////////////////////////    -->

<tr>
	<th style="width:25%">Aeropuerto Salida</th>
	<td colspan=3 align=leftstyle="width:75%" style="width:75%">	
			
        <% if Transportista2 <> "" then 

            result = WsAirportDisplay(Conn, rs, "IMPORT", "SALIDA", AirportDepID2, Country2, 0, Request("Transportista2"), 0)
            
            if AirportDepID2 = "" then %>
			
				<input type="hidden" name="awb_frame2" value="3" />					
			
				<select class="style10" name="AirportDepID2" onChange="document.awb_frame.submit();move();" id="Select2">
				<%=result(0)%>
				</select>	
			
			<% else %>
								
				<%					
                    iAirportFromCode = result(1)
					CarrierName = result(2)	

                    Airports2 = "1"
				%>	
				
				<input type="hidden" name="AirportDepID2" value="<%=AirportDepID2%>" />
                <input type="hidden" name="iAirportFromCode" value="<%=iAirportFromCode%>" />
				
				<input type="text" name="" value="<%=AirportDepID2%>" disabled size=7/>
				<input type="text" name="" value="<%=CarrierName%>" disabled size=50 />

                <button type="submit" name="BtnBorra" value="Borra Aeropuerto Salida" class="Boton2 cOrag" title="Edita Aeropuerto Salida"><img src="img/glyphicons_150_edit.png" /></button>
				
			
			<% end if %>

		<% end if %>
	</td>
	
</tr>

<!--  ////////////////////////////////////// AEROPUERTO DESTINO IMPORT //////////////////////////////////////    -->

<tr>
	<th style="width:25%">Aeropuerto Destino</th>
	<td colspan=3 align=leftstyle="width:75%" style="width:75%">
	
		<% if AirportDepID2 <> "" then %>		
			
			<% if Transportista2 <> "" then 

                    result = WsAirportDisplay(Conn, rs, "IMPORT", "DESTINO", AirportDesID2, Country2, Request("Routing2"), Request("Transportista2"), Request("AirportDepID2"))

                    Airports2 = "0"

				end if
			%>
		
			<% if AirportDesID2 = "" then %>
			
				<input type="hidden" name="awb_frame2" value="3" />					
						
				<select class="style10" name="AirportDesID2" id="Select1" style="width:200px" onChange="document.awb_frame.submit();move();">
				<%=result(0)%>
				</select>
			
			<% else %>

				<%					
                    iAirportToCode = result(1)
					CarrierName = result(2)	
				%>	
				
				<input type="hidden" name="AirportDesID2" value="<%=AirportDesID2%>"/>
                <input type="hidden" name="iAirportToCode" value="<%=iAirportToCode%>"/>
				
				<input type="text" name="" value="<%=AirportDesID2%>" disabled size=7/>
				<input type="text" name="" value="<%=CarrierName%>" disabled size=50 />

                <button type="submit" name="BtnBorra" value="Borra Aeropuerto Destino" class="Boton2 cOrag" title="Edita Aeropuerto Destino"><img src="img/glyphicons_150_edit.png" /></button>
			
			<% end if %>

		<% end if %>
	</td>
</tr>	


<% end if %>



<%

CloseOBJ Conn

%>



<% if 2 = 1 then %>

<!--  ////////////////////////////////////// TIPO AWB ///////////////////////////////////////////    -->

<tr>
	<th  width=25%>Tipo Awb</th>
    <td colspan=3 align=leftstyle="width:75%" style="width:75%">    
	    
        <% if AirportDesID2 <> "" and AirportDepID2 <> "" then %>		
				
			<% if BtnReplica2 = "" then %>
			
				<input type="hidden" name="awb_frame2" value="3" />		

                <%
				'<input type=submit name="BtnReplica2" value="Consolidado" class="Boton cBlue">
                ' 'if AwbType = 1 then 'EXPORT 
    		    '	<input type=submit name="BtnReplica2" value="Directo" class="Boton cBlue">
                ''end if 
                %>

                <select name="BtnReplica2" onchange="document.awb_frame.submit();move();">
                <option value="">-- Seleccione --</option>
                <option value="Consolidado">Consolidado</option>
                <option value="Directo">Directo</option>
                <option value="Master-Hija">Master-Hija</option>
                <option value="Hija-Directa">Hija-Directa</option>
                <option value="Master-Master-Hija">Master-Master-Hija</option>
                </select>

			
			<% else %>
				
				<input type="text" name="" value="<%=BtnReplica2%>" disabled size=7 />
				<input type="hidden" name="BtnReplica2" value="<%=BtnReplica2%>" />
                <button type="submit" name="BtnBorra" value="Borra Tipo AWB" class="Boton2 cOrag" title="Edita Tipo AWB"><img src="img/glyphicons_150_edit.png" /></button>

			<% end if %>
		
		<% end if %>
		
    </td>	
</tr>

<% end if %>


<!--  ////////////////////////////////////// CORRELATIVO AWB ///////////////////////////////////////////    
style="background:white;text-align:left;border:0px;border-bottom:0px solid gray;"
-->

	
<tr>
	<th width=25%>
        <% if AwbType = 1 then 'EXPORT  %>
			AWBNumber
		<% else %>
			AWBNumber 
        <% end if %>    
    </th>
    <td colspan=3 align=leftstyle="width:75%" style="width:75%">    	
        <% if AirportDesID2 <> "" and AirportDepID2 <> "" then %>
        								
			<% if AWBNumber = "" then %>
			
                    <%                                      
                    if Country3 = "" then
                        Country3 = Country2
                    end if
                                          
                    if Country3 = "GT" and AwbType = 1 then 'EXPORT      
                        'response.write "(" & Request("linea") & ")(" & Country2 & ")(" & AwbType & ")(" & Transportista2 & ")"
				        AWBNumber = NextAWBNumber(Conn, CheckNum(AwbType), CheckNum(Transportista2), "Nuevo")                     
                    %>                    
				        <input type="hidden" name="AWBNumber2" value="<%=AWBNumber%>" />								
				        <input type="text" name="" value="<%=AWBNumber%>" disabled size=25 />
			        <% else %>
				        <input type="text" id="AWBNumber2" name="AWBNumber2" value="" size=25 autocomplete=off minlength=3 />    				    
                        <!-- <input type=submit name="BtnMaster2" value="Nuevo" class="Boton cBlue"> -->
                        <button type="submit" type="submit" name="BtnMaster2" value="Nuevo" class="Boton2 cBlue2" title="Acepta AWBNumber" onclick="if (document.getElementById('AWBNumber2').value=='') { alert('Debe digitar Master'); return false; }" ><img src="img/glyphicons_193_circle_ok.png" /></button>
    				    <input type="hidden" name="awb_frame2" value="3" />		
                    <% end if %>
			
			<% else %>

				<input type="hidden" name="BtnMaster2" value="<%=BtnMaster2%>" >
                <input type="text" name="AWBNumber2" value="<%=AWBNumber%>" readonly size="25" >  
                <button type="submit" name="BtnBorra" value="Borra AwbNumber" class="Boton2 cOrag" title="Edita AwbNumber"><img src="img/glyphicons_150_edit.png" /></button>							
							
			<% end if %>
		
		<% end if %>		

                    <% if AwbType = 1 AND Country2 = "GT" then 'EXPORT  %>
				        Correlativo de Linea Aerea
			        <% else %>
				        
                    <% end if %>   


    </th>	
</tr>



<% 

'response.write "(" & BtnHouse2 & ")"

if AWBNumber <> "" and (BtnReplica2 = "Consolidado" or BtnReplica2 = "Master-Hija" or BtnReplica2 = "Master-Master-Hija") then

%>

<!-- /////////////////////////////////////////////////HAWBNumber////////////////////////////////////////////////////////// -->
<tr>
	<th style="width:25%">
        HAWBNumber
    </th>
    <td style="width:25%">
    <% if BtnHouse2 = "" then %>
        <% if AWBNumber = "" then %>
            <input type=button value="Asignar" title="Nuevo HAWB" class="Read" disabled style="margin-bottom:3px" >
            <input type=button value="Manual" title="Nuevo HAWB" class="Read" disabled >   		
        <% else %>

            <input type="hidden" name="awb_frame2" value="3" size="2"/>	

            <% if AwbType = 1 then 'EXPORT %>
                <input type=submit name="BtnHouse2" value="Asignar" title="Asignar Correlativo HAWB" <% if Request("BtnEdita") = "" then %> class="Boton cBlue" <% else %> class="cGray" disabled <% end if %> style="margin-bottom:3px" > 
            <% end if %>   
            <input type=submit name="BtnHouse2" value="Manual" title="Digitar HAWB" <% if Request("BtnEdita") = "" then %> class="Boton cBlue" <% else %> class="cGray" disabled <% end if %> >
        <% end if %>       
              
    <% else %>
            
            <input type="text" name="BtnHouse2" value="<%=BtnHouse2%>" title="Nuevo HAWB" class="Read" readonly size=7> 
            <button type="submit" name="BtnBorra" value="Borra BtnHouse" class="Boton2 cOrag" title="Edita BtnHouse"><img src="img/glyphicons_150_edit.png" /></button>

			<% if BtnHouse2 = "Asignar" then %>            
				<% HAWBNumber2 = NextHAWBNumber(HAWBNumber2, Conn, AwbType, Country2, "", "Asignar", AWBNumber) %>

                <th style="width:10%">
                    HAWBNumber 
                </th>
                <td style="width:40%">                 
				<input type=text name=HAWBNumber2 value="<%=HAWBNumber2%>" readonly>
                <button type="submit" name="BtnBorra" value="Borra HAWBNumber" class="Boton2 cOrag" title="Edita HAWBNumber"><img src="img/glyphicons_150_edit.png" /></button>
                </td>

			<% else %>			
                
                <th style="width:10%">
                    HAWBNumber 
                </th>
                <td style="width:40%">   
                <% if HAWBNumber2 <> "" then %>
				
                    <input type=text name=HAWBNumber2 value="<%=HAWBNumber2%>" readonly />
                    <button type="submit" name="BtnBorra" value="Borra HAWBNumber" class="Boton2 cOrag" title="Edita HAWBNumber"><img src="img/glyphicons_150_edit.png" /></button>

                <% else %>                    

                    <input type="hidden" name="awb_frame2" value="3" />                 
				    <input type=text id=HAWBNumber2 name=HAWBNumber2 autocomplete=off minlength=3>                                                
                    <button type=submit name="BtnAcepta" value="Acepta HawbNumber" class="Boton2 cBlue2"  onclick="if (document.getElementById('HAWBNumber2').value=='') { alert('Debe digitar House'); return false; }" title="Acepta HawbNumber"><img src="img/glyphicons_193_circle_ok.png" /></button>

                <% end if %>

                </td>

			<% end if %> 
	<% end if %> 

    </td>
</tr>

<% else %>      

    <% if AWBNumber <> "" and BtnReplica2 <> "" then %>

        <% BtnHouse2 = "Manual" %>
        <% HAWBNumber2 = AWBNumber %>
          
        <input type=hidden name=HAWBNumber2 value="<%=HAWBNumber2%>"  />
        <input type=hidden name=BtnHouse2 value="<%=BtnHouse2%>"  />

    <% end if %> 

<% end if %> 



<% if HAWBNumber2 <> "" and (BtnReplica2 = "Consolidado" or BtnReplica2 = "Master-Hija" or BtnReplica2 = "Master-Master-Hija") then %>


<% if 2 = 1 then %>

<!-- /////////////////////////////////////////////////Tipo HAWBNumber////////////////////////////////////////////////////////// -->
<tr>
	<th style="width:25%">
        Tipo HAWBNumber
    </th>
    <td style="width:25%"> 

        

    <% if HAWBNumberTipo2 = "" then %>

        <% if HAWBNumber2 = "" then %>
            <input type=button value="Completar" title="Nuevo HAWB" class="Read" disabled style="margin-bottom:3px" >
            <input type=button value="Express" title="Nuevo HAWB" class="Read" disabled >   		
        <% else %>
            <input type="hidden" name="awb_frame2" value="3" />	            
            <input type=submit name="HAWBNumberTipo2" value="Completar" title="Completar House" <% if Request("BtnEdita") = "" then %> class="Boton cBlue" <% else %> class="cGray" disabled <% end if %> >             
            <input type=submit name="HAWBNumberTipo2" value="Express" title="Crear House Express" <% if Request("BtnEdita") = "" then %> class="Boton cBlue" <% else %> class="cGray" disabled <% end if %> >
        <% end if %>       
              
    <% else %>
        
            <input type="text" name="HAWBNumberTipo2" value="<%=HAWBNumberTipo2%>" title="Tip HAWBNumbero" class="Read" readonly size=7> 
            <button type="submit" name="BtnBorra" value="Borra Tipo HAWBNumber" class="Boton2 cOrag" title="Edita Tipo HAWBNumber"><img src="img/glyphicons_150_edit.png" /></button>

			<% if HAWBNumberTipo2 = "Completar" then %>            

                <input type=hidden name=Piezas2 value="1" readonly />
                 <% Piezas2 = "1" %>

			<% else %>			
                
                <th style="width:25%">Piezas</th>
                <td style="width:25%">                

                <% if Piezas2 <> 0 then %>
				    
                    <input type=text name=Piezas2 value="<%=Piezas2%>" readonly  style="width:80px" />
                    <button type="submit" name="BtnBorra" value="Borra Piezas" class="Boton2 cOrag" title="Edita Piezas"><img src="img/glyphicons_150_edit.png" /></button>

                <% else %>                    
                    <input type="hidden" name="awb_frame2" value="3" />	            
				    <input type=number id=Piezas2 name=Piezas2 autocomplete=off style="width:80px" >
                    <button type=submit name="BtnAcepta" value="Acepta Piezas" class="Boton2 cBlue2"  onclick="if (document.getElementById('Piezas2').value=='') { alert('Debe digitar Piezas'); return false; }"  title="Acepta Piezas"><img src="img/glyphicons_193_circle_ok.png" /></button>

                <% end if %>

                </td>

			<% end if %> 
	<% end if %> 

    </td>
</tr>

<% else '2 = 1 %>          

<tr>
<th>Piezas</th>
<td>

        <% if Piezas2 <> 0 then %>
				    
            <input type=text name=Piezas2 value="<%=Piezas2%>" readonly  style="width:80px" />
            <button type="submit" name="BtnBorra" value="Borra Piezas" class="Boton2 cOrag" title="Edita Piezas"><img src="img/glyphicons_150_edit.png" /></button>

        <% else %>                    
            <input type="hidden" name="awb_frame2" value="3" />	            
			<input type=number id=Piezas2 name=Piezas2 autocomplete=off style="width:80px" >
            <button type=submit name="BtnAcepta" value="Acepta Piezas" class="Boton2 cBlue2"  onclick="if (document.getElementById('Piezas2').value=='') { alert('Debe digitar Piezas'); return false; }"  title="Acepta Piezas"><img src="img/glyphicons_193_circle_ok.png" /></button>

        <% end if %>

</td>
</tr>

<% end if %>

<% else %>          

    <% if  HAWBNumber2 <> "" and BtnReplica2 <> "" then %>

        <% HAWBNumberTipo2 = "Completar" %>
        <% Piezas2 = "1" %>
          
        <input type=hidden name=HAWBNumberTipo2 value="<%=HAWBNumberTipo2%>"  />
        <input type=hidden name=Piezas2 value="<%=Piezas2%>"  />

    <% end if %> 

<% end if %> 



<%
    if AWBNumber <> "" and AWBNumber = HAWBNumber2 and (BtnReplica2 = "Consolidado" or BtnReplica2 = "Master-Hija" or BtnReplica2 = "Master-Master-Hija") then
        HAWBNumber2 = ""
        Piezas2 = 0
        response.write "<font color=red><blink>House debe ser distinto a la master</blink></font>"
    end if
				 
    item0 = True
    if Piezas2 = 0 then 
        item0 = False
    end if

    item1 = False
    if BtnReplica2 = "Master-Hija" or BtnReplica2 = "Hija-Directa" or BtnReplica2 = "Master-Master-Hija" then        
        item1 = True    
    end if



	if CheckNum(Piezas2) > 0 and item1 = True then  'si ya ingreso piezas y es modalidad nueva
%>
<!--  ////////////////////////////////////// TipoCarga ///////////////////////////////////////////    -->
<tr>
    <th style="width:25%">Tipo Cargo</th>
	<td align=left style="width:25%">    
    <%  	
	    if TipoCarga2 = "" then 

            CountList5Values = -1        
            QuerySelect = "SELECT ""tpt_codigo"", ""tpt_descripcion"", ""tpt_pk"" FROM ""ti_pricing_tipo"" WHERE ""tpt_tipo"" = 'TIPO_CARGA' AND ""tpt_tps_fk"" = '1' ORDER BY ""tpt_descripcion"""
            'response.write QuerySelect & "<br>"
            OpenConn3 Conn2        		
            Set rst = Conn2.Execute(QuerySelect)
		    If Not rst.EOF Then
    		    aList5Values = rst.GetRows
        	    CountList5Values = rst.RecordCount
		    End If
            CloseOBJs rst, Conn2
		%>		    
            <input type="hidden" name="awb_frame2" value="3" />				
		    <select name="TipoCarga2" onChange="document.awb_frame.submit();move();">
		    <option value="-1">Seleccionar</option>
		    <%
	        For i = 0 To CountList5Values-1
		    %>
		    <option value="<%=aList5Values(0,i)%>"><%=aList5Values(1,i)%></option>
		    <%
			Next
		    %>
		    </select>
	    <% else %>

		    <input type="text" name="TipoCarga2" value="<%=TipoCarga2%>" readonly />		
     
            <button type="submit" name="BtnBorra" value="Borra TipoCarga" class="Boton2 cOrag" title="Borra TipoCarga"><img src="img/glyphicons_150_edit.png" /></button>        

	    <% end if %>
	    
	</td>


    <% 
        if TipoCarga2 = "" then 
                
            item0 = False

        else

    %>

    <th>Peso</th>
    <td>

        <% if CheckNum(Peso2) = 0 then 

                item0 = False        
        %>

            <input type="hidden" name="awb_frame2" value="3" />	           
             
			<input type=text id=Peso2 name=Peso2 autocomplete=off style="width:80px" >

            <button type=submit name="BtnAcepta" value="Acepta Peso" class="Boton2 cBlue2"  onclick="if (document.getElementById('Peso2').value=='') { alert('Debe digitar Peso'); return false; }"  title="Acepta Peso"><img src="img/glyphicons_193_circle_ok.png" /></button>
				   
        <% else 

                item0 = True
        %>                    

            <input type=text name=Peso2 value="<%=Peso2%>" readonly  style="width:80px" />

            <button type="submit" name="BtnBorra" value="Borra Peso" class="Boton2 cOrag" title="Edita Peso"><img src="img/glyphicons_150_edit.png" /></button>

        <% end if %>

    </td>

    <% end if %>

</tr>

<% 
    end if 
%>

<tr>
	<th  width=25%></th>
    <td colspan=3 align=leftstyle="width:75%" style="width:75%">            
  
		<% if item0 = False then %>
			
            <input type=button value="Completar datos Master AWB" title="Nuevo HAWB" class="Read" disabled >

		<% else %>			

            <%   
            
            OpenConn Conn
                         
            SQLQuery = "SELECT AwbID, CreatedDate, CreatedTime, AwbNumber, HAwbNumber FROM " & IIf(AwbType = 1,"Awb","Awbi") & " WHERE Countries IN ('" & Country2 & "') AND AwbNumber = '" & AWBNumber & "' AND " & IIf(BtnReplica2 = "Directo","HAwbNumber = '" & AWBNumber & "'","HAwbNumber = ''")                 
            'response.write (SQLQuery & "<BR>" + Country2)
            Set rs = Conn.Execute(SQLQuery) 
            if Not rs.EOF then '2018-03-07 esto es util, se debe descomentar en produccion                                                                                                           
            %>            
            <!-- <input type=button value="Completar datos Master AWB" class="Read" disabled> -->
            <font color=red>Correlativo AWB ya fue utilizado</font>
            <%
            else            
            %>				
            <input type="hidden" name="awb_frame2" value="3" />						
			<input type=submit name="completar" value="Completar datos Master AWB" title="Nuevo Master" class="Boton cBlue">               
            <%
            end if
            
            closeOBJ rs

            closeOBJ Conn

            %>
			
		<% end if %>
		
    </td>	
</tr>
</table>

</form>


<%'="(" & BtnHouse2 & ")(" & AWBNumber & ")(" & HAWBNumber2 & ")(" & BtnReplica2 & ")(" & Piezas2 & ")" %>

<% Response.End %>



