<%
    On Error Resume Next
    
    AwbType = CheckNum(Request("AT"))

    AWBNumber = aTableValues(4, 0)
    Country2 = aTableValues(78, 0)

    replica = Request("replica")

    'response.write "(" & Request("replica") & ")(" & replica & ")<br>"

    OpenConn Conn

    if Request("Peso2") <> "" then               
        if Request("BtnCancela") = "Cancela Peso" then            
            Peso2 = ""
        else            
            Peso2 = Request("Peso2")            
        end if            
    end if


    if Request("Piezas2") <> "" then               
        if Request("BtnCancela") = "Cancela Piezas" then            
            Piezas2 = ""
            Peso2 = ""
        else            
            Piezas2 = Request("Piezas2")            
        end if            
    end if

    
    if Request("HAWB_new") <> "" then '2018-04-18

        ObjectID = InsertGuia(Conn, rs, replica, AwbType, AWBNumber, Request("HAWBNumber2"), Country2, Request("Piezas2"), Request("Peso2"), Request("Transportista2"), Request("AirportDepID2"), Request("AirportDesID2"), Request("iAirportFromCode"), Request("iAirportToCode"), IIf(AwbType = 1,", 'Consolidado'",""), Iif(replica = "Consolidado" or replica = "Master-Hija" or replica = "Master-Master-Hija", "Consolidado", "Directo"), "Hija", Request("TipoCarga2"))

        if ObjectID > 0 then

            if replica = "Master-Master-Hija" then

                ReplicarHeaderRubros Conn, rs, False, AwbType, AWBNumber, Request("HAWBNumber2"), ObjectID, ObjectIDtmp, ClientCollectID_tmp, True, Request("Peso2")
                        
            end if
            
        end if

        Piezas2 = ""
        Peso2 = ""
        HAWBNumber = ""
        HAWBNumber2 = ""


    end if  


    if Request("AWB_edit") <> "" then '2018-05-02
        QuerySelect = "UPDATE " & IIf(AwbType = 1,"Awb","Awbi") & " SET CarrierID = '" & Request("Transportista2") & "', AirportDepID = '" & Request("AirportDepID2") & "', AirportDesID = '" & Request("AirportDesID2") & "', AirportToCode1 = '" & Request("iAirportFromCode") & "', AirportToCode2 = '" & Request("iAirportToCode") & "', RequestedRouting = '" & Request("iAirportFromCode") & "/" & Request("iAirportToCode") & "' WHERE AwbID = '" & aTableValues(0, 0)  & "'"
	    'response.write "(awb_frame2)<br>" & QuerySelect & "<br>"        
        Conn.Execute(QuerySelect)
        response.write "Cambios fueron guardados!<br>"        
    end if  

    

    CountList6Values = -1
    '                       0       1           2           3           4           5           6               7               8       9           10      11              12              13          14              15                  16          17
    QuerySelect = "SELECT AwbID, CreatedDate, CreatedTime, AwbNumber, HAwbNumber, Countries, ConsignerData, TotNoOfPieces, TotWeight, Routing, CarrierID, AirportDepID, AirportDesID, SalespersonID, ShipperData, AgentContactSignature, RoutingID, AgentData FROM " & IIf(Request("AT") = "1","Awb","Awbi") 
    '(AwbNumber='" & AWBNumber & "' AND HAwbNumber <> '') AND (AwbNumber='" & AWBNumber & "' AND HAwbNumber<>'" & AWBNumber & "')"

    QuerySelect = QuerySelect & " WHERE AwbNumber='" & AWBNumber & "' ORDER by HAwbNumber" 

    'response.write QuerySelect & "<br><br>"

    Set rst = Conn.Execute(QuerySelect)
    If Not rst.EOF Then
        aList6Values = rst.GetRows
        CountList6Values = rst.RecordCount
    End If
    CloseOBJ rst

    'response.write "(" & CountList6Values & ")<br>"
    

        OpenConn2 Conn2			
        For i = 0 To CountList6Values-1

                if CheckNum(aList6Values(16,i)) > 0 and CheckNum(aList6Values(9,i)) = 0 then
		            QuerySelect = "select routing from routings where id_routing = '" & aList6Values(16,i) & "'"
                    'response.write QuerySelect & "<br>"
		            Set rst = Conn2.Execute(QuerySelect)
		            If Not rst.EOF Then
                        if rst("routing") <> "" then
                            aList6Values(9,i) = rst("routing")
                        end if
                    end if
                    CloseOBJ rst
                end if

                'response.write "(" & aList6Values(16,i) & ")(" & aList6Values(9,i) & ")<br>"
   		Next
        CloseOBJ Conn2


    '                       0           1               2           3           4           5           6               7           8           9       10          11              12          13              14      15      16      17          18          19              20          21                                      22        23                                                                                            22        23
    'QuerySelect = "SELECT AwbID, a.CreatedDate, a.CreatedTime, AwbNumber, HAwbNumber, a.Countries, ConsignerData, TotNoOfPieces, TotWeight, Routing, a.CarrierID, AirportDepID, AirportDesID, SalespersonID, b.Name, c.Name, d.Name, Routing, c.AirportCode, d.AirportCode, flg_master,	flg_totals" & IIf(Request("AT") = "1",", replica, RoutingID FROM Awb",", case when AwbNumber = HAwbNumber then 'DIRECTA' else 'CONSOLIDADO' end as replica, RoutingID FROM Awbi") & " a, Carriers b, Airports c, Airports d WHERE a.CarrierID = b.CarrierID AND a.AirportDepID = c.AirportID AND a.AirportDesID = d.AirportID AND AwbID = " & aList6Values(0,0) 

    '                       0           1               2           3           4           5           6               7           8           9       10          11              12          13              14   15  16      17   18  19      20          21                                22        23            24        25                                                                                                22        23        24              25
    QuerySelect = "SELECT AwbID, a.CreatedDate, a.CreatedTime, AwbNumber, HAwbNumber, a.Countries, ConsignerData, TotNoOfPieces, TotWeight, Routing, a.CarrierID, AirportDepID, AirportDesID, SalespersonID, b.Name, '', '', Routing, '', '', flg_master, flg_totals" & IIf(AwbType = "1", ", replica, RoutingID, ShipperData, AgentData FROM Awb", ", case when AwbNumber = HAwbNumber then 'DIRECTA' else 'CONSOLIDADO' end as replica, RoutingID, ShipperData, AgentData FROM Awbi") & " a, Carriers b WHERE a.CarrierID = b.CarrierID AND AwbID = " & aList6Values(0,0)

    'response.write QuerySelect & "<br>"
    Set rst = Conn.Execute(QuerySelect)
    If rst.EOF Then 
        response.write "Existio un error al intentar traer Aereo Data<br>"
        response.write QuerySelect & "<br>"
        Response.End
    end if 

    if rst("Routing") <> "" then
        Routing2 = rst("Routing")
    else
        Routing2 = aList6Values(9,0)
    end if



    Country2 = rst("Countries")

    if Request("Transportista2") = "" then
        Transportista2 = rst("CarrierID")
    else        
        if Request("BtnCancela") = "Cancela Linea Aerea" then            
            Transportista2 = rst("CarrierID")
        else            
            Transportista2 = Request("Transportista2")            
        end if
    end if
    
    if Request("AirportDepID2") = "" then        
        AirportDepID2 = rst("AirportDepID")
    else        
        if Request("BtnCancela") = "Cancela Aeropuerto Salida" then            
            AirportDepID2 = rst("AirportDepID")
        else            
            AirportDepID2 = Request("AirportDepID2")
        end if            
    end if
    
    if Request("AirportDesID2") = "" then        
        AirportDesID2 = rst("AirportDesID")
    else        
        if Request("BtnCancela") = "Cancela Aeropuerto Destino" then            
            AirportDesID2 = rst("AirportDesID")
        else            
            AirportDesID2 = Request("AirportDesID2")
        end if            
    end if

    if Request("HAWBNumber2") <> "" then

        if Request("HAWB_new") <> "" then '2018-04-18

        else        
            HAWBNumber2 = Request("HAWBNumber2")
            'response.write "1(" & Request("HAWBNumber2") & ")<br>"
        end if

    end if

    if Request("BtnCancela") = "Cancela HawbNumber" then            
        BtnHouse2 = ""
        Piezas2 = ""
        Peso2 = ""
    else            
        BtnHouse2 = Request("BtnHouse2")
    end if            

    if Request("BtnAcepta") <> "" then

        if Request("BtnAcepta") = "Acepta Linea Aerea" then
            if Request("Transportista2") <> "" then
                if CheckNum(Request("Transportista2")) <> CheckNum(rst("CarrierID")) then
                    AirportDepID2 = ""
                    AirportDesID2 = ""
                end if
            end if
        end if


        if AwbType = 1 then        

            if Request("BtnAcepta") = "Acepta Aeropuerto Salida" then
                if Request("AirportDepID2") <> "" then
                    if CheckNum(Request("Transportista2")) <> CheckNum(rst("CarrierID")) or CheckNum(Request("AirportDepID2")) <> CheckNum(rst("AirportDepID")) then
                        AirportDesID2 = ""
                    end if
                end if
            end if

            'if Request("BtnAcepta") = "Acepta Aeropuerto Destino" then
            '    if Request("AirportDesID2") <> "" then
            '    end if
            'end if

        else 

            if Request("BtnAcepta") = "Acepta Aeropuerto Destino" then
                if Request("AirportDesID2") <> "" then
                    if CheckNum(Request("Transportista2")) <> CheckNum(rst("CarrierID")) or CheckNum(Request("AirportDesID2")) <> CheckNum(rst("AirportDesID")) then                    
                        AirportDepID2 = ""
                    end if
                end if
            end if

            'if Request("BtnAcepta") = "Acepta Aeropuerto Salida" then          
            '    if Request("AirportDepID2") <> "" then        
            '    end if
            'end if

        end if

    end if

    
    'response.write "(" & Transportista2 & ")(" & AirportDepID2 & ")(" & AirportDesID2 & ")<br>"
   
    Select Case Request("BtnEdita")
	Case "Edita Linea Aerea"
        Transportista2 = ""                
        
	Case "Edita Aeropuerto Salida"
        AirportDepID2 = ""

    Case "Edita Aeropuerto Destino"
        AirportDesID2 = ""

    End Select



    

    AwbNumber = rst("AwbNumber")
    HawbNumber = rst("HAwbNumber")

    documento = ""

    '2022-03-02 ya no debe leer nada del baw
    'QuerySelect = "SELECT distinct b.AWBID, b.DocTyp, b.InvoiceID, b.DocType, b.Expired FROM Awb a, ChargeItems b WHERE a.AWBID = '" & rst("AwbID") & "' AND a.AwbID = b.AWBID AND b.Expired = '0' AND b.DocType IN (1,4) AND DocTyp = '0' LIMIT 1"
    'response.write(QuerySelect & "<br>")
    'Set rs_ch = Conn.Execute(QuerySelect)  
	'If Not rs_ch.EOF Then
    '    openConnBAW ConnBaw1
    '    if rs_ch(3) = 1 then 'factura
    '        QuerySelect = "SELECT tfa_serie, tfa_correlativo, tfa_id, tfa_hbl, tfa_mbl FROM tbl_facturacion WHERE tfa_id = '" & rs_ch(2) & "' AND tfa_ted_id != '3' LIMIT 1"
    '    end if                    
    '    if rs_ch(3) = 4 then 'nota credito
    '        QuerySelect = "SELECT tnc_serie, tnc_correlativo FROM tbl_nota_credito WHERE tnc_id = '" & rs_ch(2) & "' AND tnc_ted_id != '3' LIMIT 1"
    '    end if
    '    'response.write(QuerySelect & "<br>")
    '    if rs_ch(3) = 1 or rs_ch(3) = 4 then 'factura
    '        Set rs_bw = ConnBaw1.Execute(QuerySelect)  
    '        documento = rs_bw(0) & " - " & rs_bw(1)
    '    end if
    '    CloseOBJ ConnBaw1
    'end if





    sigue = true

    if documento <> "" then 'si hay facturacion no edita keys
        sigue = false
    end if

    if rst(22) = "" then 'Import: blank; Export: replica    
        if AwbNumber = HawbNumber then        
            replica = "Directo"
        else
            replica = "Consolidado"
        end if
    else
        replica = rst("flg_master")
    end if

    TipoCarga2 = ""
    QuerySelect = "SELECT aiee_AwbID_fk, aiee_ImpExp, aiee_TipoAwb, aiee_replica, aiee_master_hija, aiee_TipoCarga FROM Awb_IE_Expansion WHERE aiee_AwbID_fk = " & rst("AwbID") & " AND aiee_ImpExp = " & AwbType
    'response.write QuerySelect & "<br>"
    Set rs = Conn.Execute(QuerySelect)
    If Not rs.EOF Then        
        If InStr(1, Session("Pricing"), "'" & Country2 & "'") = 0 Then
            replica = rs("aiee_replica") 'Consolidado / Directo
        Else
            replica = rs("aiee_TipoAwb") 'Master-Hija / Hija-Directa / Master-Master-Hija 
        End If
        TipoCarga2 = rs("aiee_TipoCarga")
    end if 


    if replica = "" or replica = "0" then


        if AWBNumber <> HAWBNumber and HAWBNumber <> "" then 'hija                
            replica = "Consolidado"
        end if                                      

        if AWBNumber <> HAWBNumber and HAWBNumber = "" then 'master consolidada
            replica = "Consolidado"
        end if
                        
        if AWBNumber = HAWBNumber and HAWBNumber <> "" then 'master directa
            replica = "Directo"
        end if

    end if

    'response.write "(replica=" & replica & ")(" & rst(22) & ")<br>"
    

    QuerySelect = "select CarrierID, UPPER(Name), Countries from Carriers where Expired = 0 and Countries = '" & Country2 & "' order by Name, Countries"
    'response.write QuerySelect
	Set rs = Conn.Execute(QuerySelect)
	If Not rs.EOF Then
		aList1Values = rs.GetRows
		CountList1Values = rs.RecordCount
	End If
	CloseOBJ rs      


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

<table width=95% align=center border=0 id="awb_frame">
<tr>
	<td align=center colspan=4>    
    <h1>GUIA MASTER :: <%=Country2%> - <%=IIf(AwbType = 1,"EXPORT","IMPORT")%> - <%=rst("AwbID")%></h1>    
    </td>
</tr>
<tr>
	<th align=left width=20%>AWB Master</th>
    <td>      
        <form name="frm_awb_frame00" action="InsertData.asp"  onsubmit="move();">
        <input type="hidden" name="replica" value="<%=replica%>" />		
        <input type="hidden" name="awb_frame2" value="1" />
        <input type="hidden" name="AWBNumber2" value="<%=AWBNumber%>" />        		
        <input type="hidden" name="Country2" value="<%=Country2%>" />
        <input type="hidden" name="Transportista2" value="<%=rst("CarrierID")%>" />
        <input type="hidden" name="AirportDepID2" value="<%=rst("AirportDepID")%>" />
        <input type="hidden" name="AirportDesID2" value="<%=rst("AirportDesID")%>" />		
        <input type="hidden" name="OID" value="<%=rst("AwbID")%>" />
        <input type="hidden" name="GID" value="1" />
        <input type="hidden" name="CD" value="<%=rst("CreatedDate")%>" />
        <input type="hidden" name="CT" value="<%=rst("CreatedTime")%>" />
        <input type="hidden" name="AT" value="<%=AwbType%>" />        
        <!-- <input type="hidden" name="Routing2" value="<%=rst("RoutingID")%>"/>          -->
        <input type="hidden" name="Routing2" value="<%=Routing2%>"/>
        <input type="hidden" name="TipoCarga2" value="<%=TipoCarga2%>" />

        <!-- <input type=submit value="" title="Editar Master <%=rst("AwbID")%>" class="Boton cBlue"> -->
        
        <input type="text" value="<%=AWBNumber%>" size=30 readonly>

        <button type=submit class="Boton2 cOrag" title="Edita Master"><img src="img/glyphicons_150_edit.png" /></button>

        </form>
    </td>
	<th align=left width=20%>Fecha</th><td><input type="text" name="" value="<%=rst("CreatedDate") & " " & rst("CreatedTime")%>" disabled size=25/></td>
</tr>


<form name="frm_awb_frame01" action="InsertData.asp"  onsubmit="move();">
        
<input type="hidden" name="OID" value="<%=rst("AwbID")%>" />
<input type="hidden" name="GID" value="1" />
<input type="hidden" name="CD" value="<%=rst("CreatedDate")%>" />
<input type="hidden" name="CT" value="<%=rst("CreatedTime")%>" />
<input type="hidden" name="AT" value="<%=AwbType%>" />
<input type="hidden" name="awb_frame2" value="2" />


<!------------ LINEA AEREA ------------->

<tr>	
	<th align=left>Linea Aerea</th><td>

        <% if Transportista2 = "" then %>
            
			<select id="Transportista2" name="Transportista2" style="width:200px"> <!-- onChange="document.frm_awb_frame01.submit();"> -->
			<option value="-1">Seleccionar</option>
			<%  
				For i = 0 To CountList1Values-1
                    CarrierName = ""
                    response.write CheckNum(aList1Values(0,i)) & " " & CheckNum(rst("CarrierID")) 
                    if CheckNum(aList1Values(0,i)) = CheckNum(rst("CarrierID")) then
                        CarrierName = " selected "
                    end if
			%>
			<option value="<%=aList1Values(0,i)%>" <%=CarrierName%> ><%=aList1Values(1,i) & " - " & aList1Values(2,i) & " - " & aList1Values(0,i)%></option>
			<%
				Next
			%>
			</select>
            <!--
            <input type=submit name="BtnAcepta" value="Acepta Linea Aerea" class="Boton cBlue" onclick="if (document.getElementById('Transportista2').value == -1) { alert('Debe Seleccionar Aerolinea'); return false; }">
            <input type=submit name="BtnCancela" value="Cancela Linea Aerea" class="Boton cRed">
            -->

            <button type=submit name="BtnAcepta" value="Acepta Linea Aerea" class="Boton2 cBlue2" title="Acepta Linea Aerea"><img src="img/glyphicons_193_circle_ok.png" /></button>
            <button type=submit name="BtnCancela" value="Cancela Linea Aerea" class="Boton2 cRed2" title="Cancela Linea Aerea"><img src="img/glyphicons_192_circle_remove.png" /></button>


        <% else %>

            <%
			CarrierName = "*"
			For i = 0 To CountList1Values-1                
				if CheckNum(aList1Values(0,i)) = CheckNum(Transportista2) then                    
					CarrierName = aList1Values(0,i) & " - " & TranslateCompany(aList1Values(2,i)) & " - " & aList1Values(1,i)
				end if				
			Next
            if CarrierName = "*" then
                CarrierName = rst("CarrierID") & " - " & TranslateCompany(Country2) & " - " & rst(14)
            end if
            %>
            
            <input type="hidden" name="Transportista2" value="<%=Transportista2%>"/>
            <input type="text" value="<%=CarrierName%>" readonly size=30 />
            <%'CarrierName%>


            <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                <!-- <input type=submit name="BtnEdita" value="Edita Linea Aerea" class="Boton cOrag"> -->                
                <button type=submit name="BtnEdita" value="Edita Linea Aerea" class="Boton2 cOrag" title="Edita Linea Aerea"><img src="img/glyphicons_150_edit.png" /></button>

            <% else %>                                
                <!-- <input type=button name="BtnEdita" value="Edita Linea Aerea" class="cGray" disabled> -->

                <!-- <button name="BtnEdita" value="Edita Linea Aerea" class="Boton2" title="Edita Linea Aerea" disabled><img src="img/glyphicons_150_edit.png" /></button> -->

            <% end if %>

        <% end if %>
    </td>
	<th align=left>Vendedor</th><td><input type="text" name="" value="<%=rst("SalespersonID")%>" disabled size=25/></td>
</tr>



<%

    'response.write "(" & Request("replica") & ")(" & replica & ")<br>"

    if replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija"  then 
        OpenConn3 Conn
		QuerySelect = "SELECT ""tpp_pk"", UPPER(""tpp_codigo""), UPPER(""tpp_nombre""), ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
    else
        QuerySelect = "SELECT b.AirportID, UPPER(b.AirportCode), UPPER(b.Name) FROM Airports b WHERE b.Expired=0 order by b.Name"
    end if    

    'response.write QuerySelect         
	Set rs = Conn.Execute(QuerySelect)
	If Not rs.EOF Then
		aList2Values = rs.GetRows
		CountList2Values = rs.RecordCount	
	End If
	CloseOBJ rs      
		
    if BtnReplica2 = "Master-Hija" or BtnReplica2 = "Hija-Directa" or BtnReplica2 = "Master-Master-Hija"  then 
    	CloseOBJ Conn     
    end if  


%>
<!------------ AEROPUERTO SALIDA ------------->

<tr>

	<th align=left>Aeropuerto Salida</th><td>
        <%
        Select Case AwbType
	    Case 1  
        '////////////////////////////////////// AEROPUERTO SALIDA EXPORT /////////////////////////////////////////////////              

        ''QuerySelect = "SELECT distinct a.AirportID, a.AirportCode, a.Name FROM Airports a, CarrierRates b WHERE a.AirportID = b.AirportDepID AND b.CarrierID = " & CheckNum(Transportista2) & " ORDER BY a.AirportCode LIMIT 50"        
        'QuerySelect = "SELECT b.AirportID, b.AirportCode, b.Name FROM CarrierDepartures a, Airports b WHERE a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID =" & CheckNum(Request("Transportista2")) & " order by b.Name"
		
        'QuerySelect = "SELECT ""tpp_pk"", ""tpp_codigo"", ""tpp_nombre"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
        'Set rs = Conn3.Execute(QuerySelect)					
		'If Not rs.EOF Then
		'    aList2Values = rs.GetRows
		'	CountList2Values = rs.RecordCount
		'End If
		'CloseOBJ rs
        %>

        <% if AirportDepID2 = "" then %>

            <% response.write "<script>.clear(); console.log('AEROPUERTO SALIDA EXPORT'); console.log('" & QuerySelect & "')</script>" %>

			<select id="AirportDepID2" name="AirportDepID2" style="width:200px"> <!--  onChange="document.frm_awb_frame01.submit();"> -->
			<option value="-1">Seleccionar</option>
			<%  
				For i = 0 To CountList2Values-1
                    CarrierName = ""
                    if CheckNum(aList2Values(0,i)) = CheckNum(rst("AirportDepID")) then
                        CarrierName = " selected "
                    end if
			%>
			<option value="<%=aList2Values(0,i)%>" <%=CarrierName%> ><%=aList2Values(2,i) & " - " & aList2Values(1,i) & " - " & aList2Values(0,i)%></option>
			<%
				Next
			%>
			</select>
            
            <!-- <input type=submit name="BtnAcepta" value="Acepta Aeropuerto Salida" class="Boton cBlue" onclick="if (document.getElementById('AirportDepID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }"> -->
            <button type=submit name="BtnAcepta" value="Acepta Aeropuerto Salida" class="Boton2 cBlue2" title="Acepta Aeropuerto Salida" onclick="if (document.getElementById('AirportDepID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }"><img src="img/glyphicons_193_circle_ok.png" /></button>

            <% if CheckNum(Transportista2) = CheckNum(rst("CarrierID")) then %>
            <!-- <input type=submit name="BtnCancela" value="Cancela Aeropuerto Salida" class="Boton cRed"> -->
            <button type=submit name="BtnCancela" value="Cancela Aeropuerto Salida" class="Boton2 cRed2" title="Cancela Aeropuerto Salida"><img src="img/glyphicons_192_circle_remove.png" /></button>
            <% end if %>

        <% else %>

			<%
			CarrierName = "*"
			For i = 0 To CountList2Values-1
				if CheckNum(aList2Values(0,i)) = CheckNum(AirportDepID2) then
                    CarrierName = aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)
                    iAirportFromCode = aList2Values(1,i)
				end if				
			Next
            if CarrierName = "*" then
                CarrierName = rst("AirportDepID") & " - " & rst(18) & " - " & rst(15)
            end if
            %>	


            <input type="hidden" name="AirportDepID2" value="<%=AirportDepID2%>" />				
            <input type="text" value="<%=CarrierName%>" readonly size=30 />
            <%'CarrierName%>

            <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                <!-- <input type=submit name="BtnEdita" value="Edita Aeropuerto Salida" class="Boton cOrag">  -->
                <button type=submit name="BtnEdita" value="Edita Aeropuerto Salida" class="Boton2 cOrag" title="Edita Aeropuerto Salida"><img src="img/glyphicons_150_edit.png" /></button>
            <% else %>
                <!-- <input type=button name="BtnEdita" value="Edita Aeropuerto Salida" class="cGray" disabled> -->
            <% end if %>

        <% end if 

	    
        
        
        
        Case 2
        '////////////////////////////////////// AEROPUERTO SALIDA IMPORT ///////////////////////////////////////   
		
        'Obteniendo listado de Aeropuertos Destino
        'QuerySelect = "select b.AirportID, b.AirportCode, b.Name, a.TerminalFeePD, a.TerminalFeeCS, a.CustomFee, a.SecurityFee, a.FuelSurcharge from CarrierDepartures a, Airports b where a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID = '" & CheckNum(Transportista2) & "' and b.AirportID <> '" & AirportDesID2 & "' order by b.Name"                
        'QuerySelect = "select b.AirportID, b.AirportCode, b.Name from CarrierDepartures a, Airports b where a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID =" & CheckNum(Request("Transportista2")) & " order by b.Name"
        
        'QuerySelect = "SELECT ""tpp_pk"", ""tpp_codigo"", ""tpp_nombre"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
        ''response.write QuerySelect         
		'Set rs = Conn3.Execute(QuerySelect)
		'If Not rs.EOF Then
		'	aList2Values = rs.GetRows
		'	CountList2Values = rs.RecordCount
		'End If
        %>

        <% if AirportDepID2 = "" then %>

            <% response.write "<script>console.clear();console.log('AEROPUERTO SALIDA IMPORT*');console.log('" & Replace(QuerySelect,"'","") & "')</script>" %>

			<select id="AirportDepID2" name="AirportDepID2" style="width:200px"> <!-- onChange="document.frm_awb_frame01.submit();" id="Select2"> -->
			<option value="-1">Seleccionar</option>
			<%  
				For i = 0 To CountList2Values-1
                    CarrierName = ""
                    if CheckNum(aList2Values(0,i)) = CheckNum(rst("AirportDepID")) then
                        CarrierName = " selected "
                    end if
			%>
			<option value="<%=aList2Values(0,i)%>" <%=CarrierName%> ><%=aList2Values(2,i) & " - " & aList2Values(1,i) & " - " & aList2Values(0,i)%></option>
			<%
				Next
			%>
			</select>	


            <!--<input type=submit name="BtnAcepta" value="Acepta Aeropuerto Salida" class="Boton cBlue" onclick="if (document.getElementById('AirportDepID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }">-->
            <button type=submit name="BtnAcepta" value="Acepta Aeropuerto Salida" class="Boton2 cBlue2" title="Acepta Aeropuerto Salida" onclick="if (document.getElementById('AirportDepID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }"><img src="img/glyphicons_193_circle_ok.png" /></button>
			    
            <% if CheckNum(Request("Transportista2")) = CheckNum(rst("CarrierID")) and CheckNum(Request("AirportDesID2")) = CheckNum(rst("AirportDesID")) then %>
            <!--<input type=submit name="BtnCancela" value="Cancela Aeropuerto Salida" class="Boton cRed">-->
            <button type=submit name="BtnCancela" value="Cancela Aeropuerto Salida" class="Boton2 cRed2" title="Cancela Aeropuerto Salida"><img src="img/glyphicons_192_circle_remove.png" /></button>
            <% end if %>

		<% else %>

			<%
			CarrierName = "*"
			For i = 0 To CountList2Values-1
				if CheckNum(aList2Values(0,i)) = CheckNum(AirportDepID2) then
					CarrierName = aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)
                    iAirportFromCode = aList2Values(1,i)
				end if				
			Next
            if CarrierName = "*" then
                CarrierName = rst("AirportDepID") & " - " & rst(18) & " - " & rst(15)
            end if
			%>					
			<input type="hidden" name="AirportDepID2" value="<%=AirportDepID2%>" />
            <input type="text" value="<%=CarrierName%>" readonly size=30 />
            <!-- <input type="hidden" name="iAirportFromCode" value="<%=iAirportFromCode%>"/> -->
				
            <%'CarrierName%>

            <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                <!-- <input type=submit name="BtnEdita" value="Edita Aeropuerto Salida" class="Boton cOrag">  -->
                <button type=submit name="BtnEdita" value="Edita Aeropuerto Salida" class="Boton2 cOrag" title="Edita Aeropuerto Salida"><img src="img/glyphicons_150_edit.png" /></button>

            <% else %>
                <!-- <input type=button name="BtnEdita" value="Edita Aeropuerto Salida" class="cGray" disabled>  -->
            <% end if %>

        <% end if

        End Select
 %>
    </td>

<!------------ AEROPUERTO DESTINO ------------->

	<th align=left>Aeropuerto Destino</th><td>
        <%
        Select Case AwbType
	    Case 1            
        '////////////////////////////////////// AEROPUERTO DESTINO EXPORT //////////////////////////////////////   

	    'Obteniendo listado de Aeropuertos Destino
        ''QuerySelect = "SELECT distinct a.AirportID, a.AirportCode, a.Name FROM Airports a, CarrierRates b WHERE a.AirportID = b.AirportDesID AND b.CarrierID = " & CheckNum(Transportista2) & " AND b.AirportDepID = " & CheckNum(AirportDepID2) & " and a.AirportID <> " & CheckNum(AirportDepID2) & " ORDER BY a.AirportCode LIMIT 50"        
        'QuerySelect = "select AirportID, AirportCode, Name from Airports where Expired=0 order by Name"

        'QuerySelect = "SELECT ""tpp_pk"", ""tpp_codigo"", ""tpp_nombre"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
	    ''response.write QuerySelect
	    'Set rs = Conn3.Execute(QuerySelect)
	    'If Not rs.EOF Then
		'    aList2Values = rs.GetRows
		'    CountList2Values = rs.RecordCount
	    'End If
        %>

        <% if AirportDepID2 <> "" then %>        

            <% if AirportDesID2 = "" then %>

                <% response.write "<script>console.clear(); console.log('AEROPUERTO DESTINO EXPORT'); console.log('" & QuerySelect & "')</script>" %>

			    <select id="AirportDesID2" name="AirportDesID2" style="width:200px"> <!-- onChange="document.frm_awb_frame01.submit();" id="Select2"> -->
			    <option value="-1">Seleccionar</option>
			    <%  
				    For i = 0 To CountList2Values-1
                        CarrierName = ""
                        if CheckNum(aList2Values(0,i)) = CheckNum(rst("AirportDesID")) then
                            CarrierName = " selected "
                        end if
			    %>
			    <option value="<%=aList2Values(0,i)%>" <%=CarrierName%> ><%=aList2Values(2,i) & " - " & aList2Values(1,i) & " - " & aList2Values(0,i)%></option>
			    <%
				    Next
			    %>
			    </select>	

                <!-- <input type=submit name="BtnAcepta" value="Acepta Aeropuerto Destino" class="Boton cBlue" onclick="if (document.getElementById('AirportDesID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }"> -->
                <button type=submit name="BtnAcepta" value="Acepta Aeropuerto Destino" class="Boton2 cBlue2" title="Acepta Aeropuerto Destino" onclick="if (document.getElementById('AirportDesID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }"><img src="img/glyphicons_193_circle_ok.png" /></button>

                <% if CheckNum(Request("Transportista2")) = CheckNum(rst("CarrierID")) and CheckNum(Request("AirportDepID2")) = CheckNum(rst("AirportDepID")) then %>
			    <!-- <input type=submit name="BtnCancela" value="Cancela Aeropuerto Destino" class="Boton cRed"> -->
                <button type=submit name="BtnCancela" value="Cancela Aeropuerto Destino" class="Boton2 cRed2" title="Cancela Aeropuerto Destino"><img src="img/glyphicons_192_circle_remove.png" /></button>
                <% end if %>


		    <% else %>

			    <%
			    CarrierName = "*"
			    For i = 0 To CountList2Values-1
				    if CheckNum(aList2Values(0,i)) = CheckNum(AirportDesID2) then
					    CarrierName = aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)
                        iAirportToCode = aList2Values(1,i) 
				    end if				
			    Next
                if CarrierName = "*" then
                    CarrierName = rst("AirportDesID") & " - " & rst(19) & " - " & rst(16)
                end if
			    %>					
			    <input type="hidden" name="AirportDesID2" value="<%=AirportDesID2%>" />
                <input type="text" value="<%=CarrierName%>" readonly size=30 />
                <%'CarrierName%>

                <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                    <!-- <input type=submit name="BtnEdita" value="Edita Aeropuerto Destino" class="Boton cOrag"> -->
                    <button type=submit name="BtnEdita" value="Edita Aeropuerto Destino" class="Boton2 cOrag" title="Edita Aeropuerto Destino"><img src="img/glyphicons_150_edit.png" /></button>
                <% else %>
                    <!--<input type=button name="BtnEdita" value="Edita Aeropuerto Destino" class="cGray" disabled>-->
                <% end if %>

            <% end if %>

        <% end if 
                
	    Case 2            
        '////////////////////////////////////// AEROPUERTO DESTINO IMPORT //////////////////////////////////////    -->

        ''QuerySelect = "select AirportID, AirportCode, Name from Airports where Country = '" & Country2 & "' and Expired=0 order by Name"
        'QuerySelect = "select AirportID, AirportCode, Name from Airports where Expired=0 order by Name"

        'QuerySelect = "SELECT ""tpp_pk"", ""tpp_codigo"", ""tpp_nombre"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
        ''response.write QuerySelect 
        'Set rs = Conn3.Execute(QuerySelect)					
	    'If Not rs.EOF Then
		'    aList2Values = rs.GetRows
		'    CountList2Values = rs.RecordCount
	    'End If
	    'CloseOBJ rs
        %>

        <% if AirportDepID2 <> "" then %>        

            <% if AirportDesID2 = "" then %>

                <% response.write "<script>console.clear();console.log('AEROPUERTO DESTINO IMPORT');console.log('" & Replace(QuerySelect,"'","") & "')</script>" %>

			    <select id="AirportDesID2" name="AirportDesID2" style="width:200px"> <!-- style="width:200px" onChange="document.frm_awb_frame01.submit();"> -->
			    <option value="-1">Seleccionar</option>
			    <%  
				    For i = 0 To CountList2Values-1
                        CarrierName = ""
                        if CheckNum(aList2Values(0,i)) = CheckNum(rst("AirportDesID")) then
                            CarrierName = " selected "
                        end if
			    %>
			    <option value="<%=aList2Values(0,i)%>" <%=CarrierName%> ><%=aList2Values(2,i) & " - " & aList2Values(1,i) & " - " & aList2Values(0,i)%></option>
			    <%
				    Next
			    %>
			    </select>

                <!-- <input type=submit name="BtnAcepta" value="Acepta Aeropuerto Destino" class="Boton cBlue" onclick="if (document.getElementById('AirportDesID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }">           -->
                <button type=submit name="BtnAcepta" value="Acepta Aeropuerto Destino" class="Boton2 cBlue2" title="Acepta Aeropuerto Destino" onclick="if (document.getElementById('AirportDesID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }"><img src="img/glyphicons_193_circle_ok.png" /></button>

                <% if CheckNum(Transportista2) = CheckNum(rst("CarrierID")) then %>
                <!-- <input type=submit name="BtnCancela" value="Cancela Aeropuerto Destino" class="Boton cRed"> -->
                <button type=submit name="BtnCancela" value="Cancela Aeropuerto Destino" class="Boton2 cRed2" title="Cancela Aeropuerto Destino"><img src="img/glyphicons_192_circle_remove.png" /></button>
                <% end if %>

            <% else %>

			    <%
			    CarrierName = "*"
			    For i = 0 To CountList2Values-1
				    if CheckNum(aList2Values(0,i)) = CheckNum(AirportDesID2) then
                        CarrierName = aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)
                        iAirportToCode = aList2Values(1,i) 
				    end if				
			    Next
                if CarrierName = "*" then
                    CarrierName = rst("AirportDesID") & " - " & rst(19) & " - " & rst(16)
                end if
                %>	
                <input type="hidden" name="AirportDesID2" value="<%=AirportDesID2%>" />				
                <input type="text" value="<%=CarrierName%>" readonly size=30 />
                <!--<input type="hidden" name="iAirportToCode" value="<%=iAirportToCode%>" />-->
                <%'CarrierName%>

                <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                    <!-- <input type=submit name="BtnEdita" value="Edita Aeropuerto Destino" class="Boton cOrag""> -->
                    <button type=submit name="BtnEdita" value="Edita Aeropuerto Destino" class="Boton2 cOrag" title="Edita Aeropuerto Destino"><img src="img/glyphicons_150_edit.png" /></button>
                <% else %>
                    <!-- <input type=button name="BtnEdita" value="Edita Aeropuerto Destino" class="cGray" disabled> -->
                <% end if %>

            <% end if %>
        
        <% end if 
        
        End Select
        %>


        <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>

        <input type="hidden" name="iAirportFromCode" value="<%=iAirportFromCode%>"/>
        <input type="hidden" name="iAirportToCode" value="<%=iAirportToCode%>" />
        <!-- <input type=submit name="AWB_edit" value="Guardar Cambios" class="Boton cBlue"> -->

        <button type=submit name="AWB_edit" value="Guardar Cambios" class="Boton2 cBlue2" title="Guardar Cambios"><img src="img/glyphicons_416_disk_saved.png" /></button>

        <% end if %>


    </td>	
</tr>

        <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
        <!--
<tr>
    <td align=center colspan=2>	    
        <input type="hidden" name="iAirportFromCode" value="<%=iAirportFromCode%>"/>
        <input type="hidden" name="iAirportToCode" value="<%=iAirportToCode%>" />
        <!-- <input type=submit name="AWB_edit" value="Guardar Cambios" class="Boton cBlue"> -->
<!--
        <button name="AWB_edit" value="Guardar Cambios" class="Boton2 cBlue" title="Guardar Cambios"><img src="img/glyphicons_444_floppy_saved.png" /></button>

    </td>	
</tr>-->
        <% end if %>

</form>



<!------------ FLAGS ------------->

<tr>
	<th align=left>Total No.Of Pieces</th><td><input type="text" name="" value="<%=rst("TotNoOfPieces")%>" disabled size=10/>
    <% if CheckNum(rst("flg_master")) = 1 then 'flg_master %>
    <font color=red>Master incompleta</font>
    <% end if %>
    </td>

	<th align=left>Total Weight</th><td><input type="text" name="" value="<%=rst("TotWeight")%>" disabled size=25/>    
    <% if CheckNum(rst("flg_totals")) = 1 then 'flg_totals %>
    <font color=red>Totales no cuadran</font>
    <% end if %>    
    </td>	
</tr>


<!------------ TIPO AWB / FACTURACION ------------->
<tr>
	<th align=left>Tipo Awb</th><td><input type="text" name="" value="<%                
                select case Iif(InStr(1, Session("Pricing"), "'" & Country2 & "'") = 0, replica, "")
                case "Master-Hija"
                    response.write "Consolidado"
                case "Hija-Directa"
                    response.write "Directo"
                case "Master-Master-Hija"
                    response.write "Consolidado2"
                case else
                    response.write replica
                end select %> " disabled size=25/></td>
	<th align=left>Doc. Facturacion</th><td><input type="text" name="" value="<%=documento%>" disabled size=25/></td>	
</tr>


<!------------ ROUTING ------------->
<tr>
    <% 'if rst("Routing") <> "" then %>
    <th align=left>Routing</th><td><input type="text" name="" value="<%=Routing2%>" readonly size=25/></td>
    <% 'end if %>
       
    <% 'if rst("AwbNumber") = rst("HAwbNumber") then %>
	<th align=left>HAWBNumber</th><td><input type="text" name="" value="<%=HawbNumber%>" disabled size=25/></td>
    <% 'end if %>
</tr>



    <% if HawbNumber = "" then %>    


<tr>
	<td align=center colspan=4 style="height:10px;"></td>
</tr>
<tr>
	<td align=center colspan=4 style="background-color:#98b0f9;border:1px solid navy;margin:5px;">
	    
	
		<form name="frm_awb_frame02" action="InsertData.asp"  onsubmit="move();">
		<input type="hidden" name="awb_frame2" value="2" /> <!-- 2 -->
		<input type="hidden" name="AWBNumber2" value="<%=AWBNumber%>" />        		
		<input type="hidden" name="Country2" value="<%=Country2%>" />                		
		<input type="hidden" name="OID" value="<%=rst("AwbID")%>" />
		<input type="hidden" name="GID" value="1" />
		<input type="hidden" name="CD" value="<%=rst("CreatedDate")%>" />
		<input type="hidden" name="CT" value="<%=rst("CreatedTime")%>" />
		<input type="hidden" name="AT" value="<%=AwbType%>" />		
		<input type="hidden" name="replica" value="<%=replica%>" />		
			
        <h1>NUEVO HOUSE AWB</h1>
                             

        <table width=100%>
        <tr>	
	        <td align=left></td>
	        <th align=center colspan=2>Aeropuertos</th>

<% if replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija"  then  %>
            <th align=left>Tipo Carga</th>
            <td align=left><input type="text" name="TipoCarga2" value="<%=TipoCarga2%>" readonly size=15 style="background-color:yellow"/></td>

<% end if %>
        </tr>
        <tr>	
	        <th align=left>Linea Aerea</th>
	        <th align=center><%=IIF(AwbType = 1,"Salida","Destino")%></th>
            <th align=center><%=IIF(AwbType = 1,"Destino","Salida")%></th>
	        <th align=left>Hawb</th>
    
            <% 'If inStr(1, Session("Countries"), "GT", 1) > 0 AND AwbType = 1 Then %> <th align=left>Piezas</th> <% 'end if %>
            <th align=left>Peso</th>
            <td align=left></td>
        </tr>
        <tr>
            <td valign=top>        
            <%        
            'QuerySelect = "select CarrierID, Name, Countries from Carriers where Expired = 0 and Countries = '" & Country2 & "' order by Countries, Name"
            ''response.write QuerySelect
		    'Set rs = Conn.Execute(QuerySelect)
		    'If Not rs.EOF Then
			'    aList1Values = rs.GetRows
			'    CountList1Values = rs.RecordCount
		    'End If
		    'CloseOBJ rs      
            %>

            <% if Transportista2 = "" then %>
            
			    <select style="display:none" id="Transportista2" name="Transportista2"> <!-- onChange="document.frm_awb_frame01.submit();"> -->
			    <option value="-1">Seleccionar</option>
			    <%  
				    For i = 0 To CountList1Values-1
                        CarrierName = ""
                        response.write CheckNum(aList1Values(0,i)) & " " & CheckNum(rst("CarrierID")) 
                        if CheckNum(aList1Values(0,i)) = CheckNum(rst("CarrierID")) then
                            CarrierName = " selected "
                        end if
			    %>
			    <option value="<%=aList1Values(0,i)%>" <%=CarrierName%> ><%response.write aList1Values(0,i) & " - " & TranslateCompany(aList1Values(2,i)) & " - " & aList1Values(1,i) %></option>
			    <%
				    Next
			    %>
			    </select>

                <input type="text" value="" readonly size=20 />

                <input style="display:none" type=submit name="BtnAcepta" value="Acepta Linea Aerea" class="Boton cBlue" onclick="if (document.getElementById('Transportista2').value == -1) { alert('Debe Seleccionar Aerolinea'); return false; }">
                <input style="display:none" type=submit name="BtnCancela" value="Cancela Linea Aerea" class="Boton cRed">

            <% else %>

                <%
			    CarrierName = "*"
			    For i = 0 To CountList1Values-1                
				    if CheckNum(aList1Values(0,i)) = CheckNum(Transportista2) then                    
					    CarrierName = aList1Values(0,i) & " - " & TranslateCompany(aList1Values(2,i)) & " - " & aList1Values(1,i)
				    end if				
			    Next
                if CarrierName = "*" then
                    CarrierName = rst("CarrierID") & " - " & TranslateCompany(Country2) & " - " & rst(14)
                end if
                %>
            
                <input type="hidden" name="Transportista2" value="<%=Transportista2%>"/>
                <input type="text" value="<%=CarrierName%>" readonly size=20 />
                <%'CarrierName%>

                <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                <!--<input type=submit name="BtnEdita" value="Edita Linea Aerea" class="Edita"> -->
                <% else %>
                <!--<input type=button name="BtnEdita" value="Edita Linea Aerea" class="Read" disabled> -->
                <% end if %>

            <% end if %>                
            </td>


            <% if AwbType = 1 then 'EXPORT %>
            <!--  ////////////////////////////////////// AEROPUERTO SALIDA EXPORT ///////////////////////////////////////    -->
            <td valign=top>
                <%
                ''QuerySelect = "SELECT distinct a.AirportID, a.AirportCode, a.Name FROM Airports a, CarrierRates b WHERE a.AirportID = b.AirportDepID AND b.CarrierID = " & CheckNum(Transportista2) & " ORDER BY a.AirportCode LIMIT 50"      
                'QuerySelect = "SELECT b.AirportID, b.AirportCode, b.Name FROM CarrierDepartures a, Airports b WHERE a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID =" & CheckNum(Request("Transportista2")) & " order by b.Name"

                'QuerySelect = "SELECT ""tpp_pk"", ""tpp_codigo"", ""tpp_nombre"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
		        'Set rs = Conn.Execute(QuerySelect)					
		        'If Not rs.EOF Then
		        '    aList2Values = rs.GetRows
			    '    CountList2Values = rs.RecordCount
		        'End If
		        'CloseOBJ rs
                %>

                <% if AirportDepID2 = "" then %>

                    <% response.write "<script>console.clear(); console.log('AEROPUERTO SALIDA EXPORT'); console.log('" & QuerySelect & "')</script>" %>

			        <select style="display:none" id="AirportDepID2" name="AirportDepID2" style="width:200px"> <!-- style="width:200px" onChange="document.frm_awb_frame01.submit();"> -->
			        <option value="-1">Seleccionar</option>
			        <%  
				        For i = 0 To CountList2Values-1
                            CarrierName = ""
                            if CheckNum(aList2Values(0,i)) = CheckNum(rst("AirportDepID")) then
                                CarrierName = " selected "
                            end if
			        %>
			        <option value="<%=aList2Values(0,i)%>" <%=CarrierName%> ><%response.Write aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)%></option>
			        <%
				        Next
			        %>
			        </select>
                    <input style="display:none" type=submit name="BtnAcepta" value="Acepta Aeropuerto Salida" class="Boton cBlue" onclick="if (document.getElementById('AirportDepID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }">


                    <input type="text" value="" readonly size=20 />

                    <% if CheckNum(Transportista2) = CheckNum(rst("CarrierID")) then %>
                    <input style="display:none" type=submit name="BtnCancela" value="Cancela Aeropuerto Salida" class="Boton cRed">
                    <% end if %>

                <% else %>

			        <%
			        CarrierName = "*"
			        For i = 0 To CountList2Values-1
				        if CheckNum(aList2Values(0,i)) = CheckNum(AirportDepID2) then
                            CarrierName = aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)
                            iAirportFromCode = aList2Values(1,i)
				        end if				
			        Next
                    if CarrierName = "*" then
                        CarrierName = rst("AirportDepID") & " - " & rst(15) & " - " & rst("AirportCode")
                    end if
                    %>	
                    <input type="hidden" name="AirportDepID2" value="<%=AirportDepID2%>" />				
                    <input type="text" value="<%=CarrierName%>" readonly size=20 />
                    <%'CarrierName%>

                    <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                    <!-- <input type=submit name="BtnEdita" value="Edita Aeropuerto Salida" class="Edita"> -->
                    <% else %>
                    <!-- <input type=button name="BtnEdita" value="Edita Aeropuerto Salida" class="Read" disabled> -->
                    <% end if %>

                <% end if %>        
            </td>



            <!--  ////////////////////////////////////// AEROPUERTO DESTINO EXPORT //////////////////////////////////////    -->
            <td valign=top>
		        <% 'Obteniendo listado de Aeropuertos Destino
                ''QuerySelect = "SELECT distinct a.AirportID, a.AirportCode, a.Name FROM Airports a, CarrierRates b WHERE a.AirportID = b.AirportDesID AND b.CarrierID = " & CheckNum(Transportista2) & " AND b.AirportDepID = " & CheckNum(AirportDepID2) & " and a.AirportID <> " & CheckNum(AirportDepID2) & " ORDER BY a.AirportCode LIMIT 50"        
                'QuerySelect = "select AirportID, AirportCode, Name from Airports where Expired=0 order by Name"
        
                'QuerySelect = "SELECT ""tpp_pk"", ""tpp_codigo"", ""tpp_nombre"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
		        ''response.write QuerySelect
		        'Set rs = Conn3.Execute(QuerySelect)
		        'If Not rs.EOF Then
			    '    aList2Values = rs.GetRows
			    '    CountList2Values = rs.RecordCount
		        'End If
                %>

                <% if AirportDepID2 <> "" then %>        

                    <% if AirportDesID2 = "" then %>

                        <% response.write "<script>console.clear(); console.log('AEROPUERTO DESTINO EXPORT'); console.log('" & QuerySelect & "')</script>" %>

			            <select style="display:none" id="AirportDesID2" name="AirportDesID2" style="width:200px"> <!-- onChange="document.frm_awb_frame01.submit();" id="Select2"> -->
			            <option value="-1">Seleccionar</option>
			            <%  
				            For i = 0 To CountList2Values-1
                                CarrierName = ""
                                if CheckNum(aList2Values(0,i)) = CheckNum(rst("AirportDesID")) then
                                    CarrierName = " selected "
                                end if
			            %>
			            <option value="<%=aList2Values(0,i)%>" <%=CarrierName%> ><%response.write aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)%></option>
			            <%
				            Next
			            %>
			            </select>	
                        
                        <input type="text" value="" readonly size=20 />

                        <input style="display:none" type=submit name="BtnAcepta" value="Acepta Aeropuerto Destino" class="Boton cBlue" onclick="if (document.getElementById('AirportDesID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }">

                        <% if CheckNum(Request("Transportista2")) = CheckNum(rst("CarrierID")) and CheckNum(Request("AirportDepID2")) = CheckNum(rst("AirportDepID")) then %>
			            <input style="display:none" type=submit name="BtnCancela" value="Cancela Aeropuerto Destino" class="Boton cRed">
                        <% end if %>

		            <% else %>

			            <%
			            CarrierName = "*"
			            For i = 0 To CountList2Values-1
				            if CheckNum(aList2Values(0,i)) = CheckNum(AirportDesID2) then
					            CarrierName = aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)
                                iAirportToCode = aList2Values(1,i) 
				            end if				
			            Next
                        if CarrierName = "*" then
                            CarrierName = rst("AirportDesID") & " - " & rst("AirportCode") & " - " & rst(16)
                        end if
			            %>					
			            <input type="hidden" name="AirportDesID2" value="<%=AirportDesID2%>" />
                        <input type="text" value="<%=CarrierName%>" readonly size=20 />
                        <%'CarrierName%>

                        <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                        <!-- <input type=submit name="BtnEdita" value="Edita Aeropuerto Destino" class="Edita"> -->
                        <% else %>
                        <!-- <input type=button name="BtnEdita" value="Edita Aeropuerto Destino" class="Read" disabled> -->
                        <% end if %>
                    <% end if %>
                <% end if %>   
            </td>	
            <% end if %>









            <% if AwbType = 2 then 'IMPORT %>
            <!--  ////////////////////////////////////// AEROPUERTO DESTINO IMPORT //////////////////////////////////////    -->
            <td valign=top>
                <%
                'QuerySelect = "select AirportID, AirportCode, Name from Airports where Country = '" & Country2 & "' and Expired=0 order by Name"        
                'QuerySelect = "select AirportID, AirportCode, Name from Airports where Expired=0 order by Name"                    
        
                'QuerySelect = "SELECT ""tpp_pk"", ""tpp_codigo"", ""tpp_nombre"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
                ''response.write QuerySelect 
                'Set rs = Conn3.Execute(QuerySelect)					
		        'If Not rs.EOF Then
		        '    aList2Values = rs.GetRows
			    '    CountList2Values = rs.RecordCount
		        'End If
		        'CloseOBJ rs
                %>

                <% if AirportDesID2 = "" then %>

                    <% response.write "<script>console.clear();console.log('AEROPUERTO DESTINO IMPORT');console.log('" & Replace(QuerySelect,"'","") & "')</script>" %>

			        <select style="display:none" id="AirportDesID2" name="AirportDesID2" style="width:200px"> <!-- style="width:200px" onChange="document.frm_awb_frame01.submit();"> -->
			        <option value="-1">Seleccionar</option>
			        <%  
				        For i = 0 To CountList2Values-1
                            CarrierName = ""
                            if CheckNum(aList2Values(0,i)) = CheckNum(rst("AirportDesID")) then
                                CarrierName = " selected "
                            end if
			        %>
			        <option value="<%=aList2Values(0,i)%>" <%=CarrierName%> ><%response.Write aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)%></option>
			        <%
				        Next
			        %>
			        </select>

                    <input type="text" value="" readonly size=20 />

                    <input style="display:none" type=submit name="BtnAcepta" value="Acepta Aeropuerto Destino" class="Boton cBlue" onclick="if (document.getElementById('AirportDesID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }">          
                    <% if CheckNum(Transportista2) = CheckNum(rst("CarrierID")) then %>
                    <input style="display:none" type=submit name="BtnCancela" value="Cancela Aeropuerto Destino" class="Boton cRed">
                    <% end if %>


                <% else %>

			        <%
			        CarrierName = "*"
			        For i = 0 To CountList2Values-1
				        if CheckNum(aList2Values(0,i)) = CheckNum(AirportDesID2) then
                            CarrierName = aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)
                            iAirportToCode = aList2Values(1,i) 
				        end if				
			        Next
                    if CarrierName = "*" then
                        CarrierName = rst("AirportDesID") & " - " & rst("AirportCode") & " - " & rst(16)
                    end if
                    %>	
                    <input type="hidden" name="AirportDesID2" value="<%=AirportDesID2%>" />				
                    <input type="text" value="<%=CarrierName%>" readonly size=20 />
                    <!--<input type="hidden" name="iAirportToCode" value="<%=iAirportToCode%>" />-->
                    <%'CarrierName%>

                    <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                    <!-- <input type=submit name="BtnEdita" value="Edita Aeropuerto Destino" class="Edita"> -->
                    <% else %>
                    <!-- <input type=button name="BtnEdita" value="Edita Aeropuerto Destino" class="Read" disabled> -->
                    <% end if %>

                <% end if %>        
            </td>

            <!--  ////////////////////////////////////// AEROPUERTO SALIDA IMPORT ///////////////////////////////////////    -->
            <td valign=top>
		        <% 'Obteniendo listado de Aeropuertos Destino
                ''QuerySelect = "select b.AirportID, b.AirportCode, b.Name, a.TerminalFeePD, a.TerminalFeeCS, a.CustomFee, a.SecurityFee, a.FuelSurcharge from CarrierDepartures a, Airports b where a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID = '" & CheckNum(Transportista2) & "' and b.AirportID <> '" & AirportDesID2 & "' order by b.Name"                
                'QuerySelect = "select b.AirportID, b.AirportCode, b.Name from CarrierDepartures a, Airports b where a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID =" & CheckNum(Request("Transportista2")) & " order by b.Name"
        
                'QuerySelect = "SELECT ""tpp_pk"", ""tpp_codigo"", ""tpp_nombre"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
                ''response.write QuerySelect         
		        'Set rs = Conn3.Execute(QuerySelect)
		        'If Not rs.EOF Then
			    '    aList2Values = rs.GetRows
			    '    CountList2Values = rs.RecordCount
		        'End If
                %>

                <% if AirportDesID2 <> "" then %>

                    <% if AirportDepID2 = "" then %>

                        <% response.write "<script>console.clear();console.log('AEROPUERTO SALIDA IMPORT*');console.log('" & Replace(QuerySelect,"'","") & "')</script>" %>

			            <select style="display:none" id="AirportDepID2" name="AirportDepID2" style="width:200px"> <!-- onChange="document.frm_awb_frame01.submit();" id="Select2"> -->
			            <option value="-1">Seleccionar</option>
			            <%  
				            For i = 0 To CountList2Values-1
                                CarrierName = ""
                                if CheckNum(aList2Values(0,i)) = CheckNum(rst("AirportDepID")) then
                                    CarrierName = " selected "
                                end if
			            %>
			            <option value="<%=aList2Values(0,i)%>" <%=CarrierName%> ><%response.write aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)%></option>
			            <%
				            Next
			            %>
			            </select>	

                        <input type="text" value="" readonly size=20 />

                        <input style="display:none" type=submit name="BtnAcepta" value="Acepta Aeropuerto Salida" class="Boton cBlue" onclick="if (document.getElementById('AirportDepID2').value == -1) { alert('Debe Seleccionar Un Aeropuerto'); return false; }">
			    
                        <% if CheckNum(Request("Transportista2")) = CheckNum(rst("CarrierID")) and CheckNum(Request("AirportDesID2")) = CheckNum(rst("AirportDesID")) then %>
                        <input style="display:none" type=submit name="BtnCancela" value="Cancela Aeropuerto Salida" class="Boton cRed">
                        <% end if %>

		            <% else %>

			            <%
			            CarrierName = "*"
			            For i = 0 To CountList2Values-1
				            if CheckNum(aList2Values(0,i)) = CheckNum(AirportDepID2) then
					            CarrierName = aList2Values(0,i) & " - " & aList2Values(1,i) & " - " & aList2Values(2,i)
                                iAirportFromCode = aList2Values(1,i)
				            end if				
			            Next
                        if CarrierName = "*" then
                            CarrierName = rst("AirportDepID") & " - " & rst("AirportCode") & " - " & rst(15)
                        end if
			            %>					
			            <input type="hidden" name="AirportDepID2" value="<%=AirportDepID2%>" />
                        <input type="text" value="<%=CarrierName%>" readonly size=20 />
                        <!-- <input type="hidden" name="iAirportFromCode" value="<%=iAirportFromCode%>"/> -->
				
                        <%'CarrierName%>

                        <% if Request("BtnEdita") = "" and CountList6Values = -1 then %>
                        <!-- <input type=submit name="BtnEdita" value="Edita Aeropuerto Salida" class="Edita"> -->
                        <% else %>
                        <!-- <input type=button name="BtnEdita" value="Edita Aeropuerto Salida" class="Read" disabled> -->
                        <% end if %>

                    <% end if %>
                <% end if %>
            </td>	
            <% end if %>





            <td valign=top>
            <% if BtnHouse2 = "" then %>
                <% if AirportDepID2 = "" or AirportDesID2 = "" then %>        
                    <input type=button value="Asignar" title="Nuevo HAWB" class="Read" disabled style="margin-bottom:3px" >
                    <input type=button value="Manual" title="Nuevo HAWB" class="Read" disabled >
                <% else %>
                    <% if AwbType = 1 then 'EXPORT %>
                        <input type=submit name="BtnHouse2" value="Asignar" title="Asignar Correlativo HAWB" <% if Request("BtnEdita") = "" then %> class="Boton cBlue" <% else %> class="cGray" disabled <% end if %> style="margin-bottom:3px" > 
                    <% end if %>   
                    <input type=submit name="BtnHouse2" value="Manual" title="Digitar HAWB" <% if Request("BtnEdita") = "" then %> class="Boton cBlue" <% else %> class="cGray" disabled <% end if %> >
                <% end if %>             
            <% else %>
			    <% if BtnHouse2 = "Asignar" then %>
				    <% HAWBNumber = NextHAWBNumber(HAWBNumber, Conn, AwbType, Country2, "", "Asignar", AWBNumber) %>
				    <input type=text name=HAWBNumber2 value="<%=HAWBNumber%>" readonly>
			    <% else %>			
                    <% if Request("HAWBNumber2") <> "" then 

                        if Request("HAWB_new") <> "" then '2018-04-18

                        else        
                            HAWBNumber = Request("HAWBNumber2")
                        end if
                        
                    %>
				        
                        <input type=text name=HAWBNumber2 value="<%=HAWBNumber%>" readonly />
                    <% else %>                    
				        <input type=text id=HAWBNumber2 name=HAWBNumber2 autocomplete=off minlength=3>                        
                        <!--<input type=submit name="BtnAcepta" value="Acepta HawbNumber" class="Boton cBlue" >-->
                        <button type=submit name="BtnAcepta" value="Acepta HawbNumber" class="Boton2 cBlue2"  onclick="if (document.getElementById('HAWBNumber2').value == '') { alert('Debe Digitar Numero House'); return false; }" title="Acepta HawbNumber"><img src="img/glyphicons_193_circle_ok.png" /></button>

                    <% end if %>
			    <% end if %>    
                <input type="hidden" name="iAirportFromCode" value="<%=iAirportFromCode%>"/>
                <input type="hidden" name="iAirportToCode" value="<%=iAirportToCode%>" />
                <input type=hidden name=BtnHouse2 value="<%=BtnHouse2%>"> 			            
                <button type=submit name="BtnCancela" value="Cancela HawbNumber" class="Boton2 cRed2" title="Cancela HawbNumber" ><img src="img/glyphicons_192_circle_remove.png" /></button>
            <% end if %>
            </td>



            <% 'If inStr(1, Session("Countries"), "GT", 1) > 0  AND AwbType = 1 Then %>

            <% If 1 = 1 Then %>
            <!-- PIEZAS -->
            <td valign=top>

                <% if HAWBNumber <> "" then %>                
			        <% if Piezas2 = "" then %>                
                        <input type=number  id=Piezas2 name=Piezas2  autocomplete=off style="width:100px" />
                        <button type=submit name="BtnAcepta" value="Acepta Piezas" class="Boton2 cBlue2"  onclick="if (document.getElementById('Piezas2').value == '') { alert('Debe Digitar Piezas'); return false; }"  title="Acepta Piezas"><img src="img/glyphicons_193_circle_ok.png" /></button>

                    <% else %>							
                        <input type=text id=Piezas2 name=Piezas2 value="<%=Piezas2%>" readonly style="width:100px" />				                       
                        <button type=submit name="BtnCancela" value="Cancela Piezas" class="Boton2 cRed2" title="Cancela Piezas" ><img src="img/glyphicons_192_circle_remove.png" /></button>

			        <% end if %>        
                <% end if %>
            </td>
			<!-- </form> 2018-04-18 -->



            <!-- PESO -->
            <td valign=top>

                <% if HAWBNumber <> "" then %>                
			        <% if Peso2 = "" then %>                
                        <input type=text  id=Peso2 name=Peso2  autocomplete=off style="width:100px" />
                        <button type=submit name="BtnAcepta" value="Acepta Peso" class="Boton2 cBlue2"  onclick="if (document.getElementById('Peso2').value == '') { alert('Debe Digitar Peso'); return false; }"  title="Acepta Peso"><img src="img/glyphicons_193_circle_ok.png" /></button>

                    <% else %>							
                        <input type=text id=Peso2 name=Peso2 value="<%=Peso2%>" readonly style="width:100px" />				                       
                        <button type=submit name="BtnCancela" value="Cancela Peso" class="Boton2 cRed2" title="Cancela Peso" ><img src="img/glyphicons_192_circle_remove.png" /></button>

			        <% end if %>        
                <% end if %>
            </td>
            
            <td valign=top nowrap>		
                
                
			<!-- valida la master hija -->
            <% if IsNull(aList6Values(6,0)) and IsNull(aList6Values(14,0)) and IsNull(aList6Values(17,0)) and replica = "Master-Master-Hija" then %>
                
                    <input type=button value="Express" title="Nuevo HAWB Express" class="Boton cBlue" onclick="alert('Primero complete la Master-Hija')" style="margin-bottom:3px">
                    <input type=button value="Nuevo" title="Nuevo HAWB" class="Boton cBlue" onclick="alert('Primero complete la Master-Hija')" >

            <% else %>							
                
                
                <% if Peso2 <> "" then %>							


<table><tr><td valign=top>

                    <input type=submit name="HAWB_new" value="<%=IIf(replica = "Master-Master-Hija","Grabar","Express") %>" title="Nuevo HAWB Express" class="Boton cBlue"  style="margin-bottom:3px">

			        </form> <!-- 2018-04-18 -->

</td><td>
<%              if replica <> "Master-Master-Hija" then

%>
                    <form name="frm_awb_frame03" action="InsertData.asp"  onsubmit="move();"> <!-- 2018-04-18  aca retornaba a Awb.asp / Awbi.asp para complementar datos -->
			        <input type="hidden" name="awb_frame2" value="1" />				
			        <input type="hidden" name="AWBNumber2" value="<%=AWBNumber%>" />
                    <input type="hidden" name="HAWBNumber2" value="<%=HAWBNumber%>" />
			        <input type="hidden" name="Country2" value="<%=Country2%>" />
			        <input type="hidden" name="Transportista2" value="<%=Transportista2%>" />
			        <input type="hidden" name="AirportDepID2" value="<%=AirportDepID2%>" />
			        <input type="hidden" name="AirportDesID2" value="<%=AirportDesID2%>" />		
			        <input type="hidden" name="replica" value="Consolidado" />						
			        <input type="hidden" name="GID" value="1" />                        
			        <input type="hidden" name="AT" value="<%=AwbType%>" />   
                    <input type="hidden" name="ObjectID2" value="<%=rst("AwbID")%>" />
                    <input type="hidden" name="vars" value="<%="OID=" & rst("AwbID") & "&CD=" & rst("CreatedDate") & "&CT=" & rst("CreatedTime")%>" />            
                    <input type="hidden" name="iAirportFromCode" value="<%=iAirportFromCode%>"/>
                    <input type="hidden" name="iAirportToCode" value="<%=iAirportToCode%>" />
            
                    <input type=submit name="HAWB_new" value="Nuevo" title="Nuevo HAWB" class="Boton cBlue" >

<%              end if

%>

</td></tr></table>

                <% else %>
                    <input type=button value="Express" title="Nuevo HAWB Express" class="cGray" disabled style="margin-bottom:3px">
                    <input type=button value="Nuevo" title="Nuevo HAWB" class="cGray" disabled>
                <% end if %>  


            <% end if 'fin de master-hija-master %>  



            </td>

            <% else %>

    
            <!-- NO REQUIERE PIEZAS NI REALIZA HOUSE EXPRESS -->  
            <td valign=top nowrap>		

                    <% if HAWBNumber <> "" then %>							
           
			        </form> <!-- 2018-04-18 -->

                    <form name="frm_awb_frame03" action="InsertData.asp"  onsubmit="move();"> <!-- 2018-04-18  aca retornaba a Awb.asp / Awbi.asp para complementar datos -->
			        <input type="hidden" name="awb_frame2" value="1" />				
			        <input type="hidden" name="AWBNumber2" value="<%=AWBNumber%>" />
                    <input type="hidden" name="HAWBNumber2" value="<%=HAWBNumber%>" />
			        <input type="hidden" name="Country2" value="<%=Country2%>" />
			        <input type="hidden" name="Transportista2" value="<%=Transportista2%>" />
			        <input type="hidden" name="AirportDepID2" value="<%=AirportDepID2%>" />
			        <input type="hidden" name="AirportDesID2" value="<%=AirportDesID2%>" />		
			        <input type="hidden" name="replica" value="Consolidado" />						
			        <input type="hidden" name="GID" value="1" />                        
			        <input type="hidden" name="AT" value="<%=AwbType%>" />   
                    <input type="hidden" name="ObjectID2" value="<%=rst("AwbID")%>" />
                    <input type="hidden" name="vars" value="<%="OID=" & rst("AwbID") & "&CD=" & rst("CreatedDate") & "&CT=" & rst("CreatedTime")%>" />            
                    <input type="hidden" name="iAirportFromCode" value="<%=iAirportFromCode%>"/>
                    <input type="hidden" name="iAirportToCode" value="<%=iAirportToCode%>" />
            
                    <input type=submit name="HAWB_new" value="Nuevo" title="Nuevo HAWB" class="Boton cBlue" >

                <% else %>
            
                    <input type=button value="Nuevo" title="Nuevo HAWB" class="cGray" disabled>
                <% end if %>  

            </td>
            <% end if %>

			</form>  

        </tr>
        </table>


    <% end if %>    

    </td>
</tr>
            



<!----------------------- HOUSES DETAIL ----------------------------------------------------------->

<%
    CarrierRates = 0
    CarrierSubTot = 0	
    if CountList6Values > 1 then  ' -1 se quito el menos uno, porque para entrar aca por lo menos tiene que tener una master y una hija
%>

<tr>
	<td align=center colspan=4 style="height:10px;"></td>
</tr>
<tr>
	<td align=center colspan=4 style="border:1px solid navy;margin:5px;">
	 
    <h1>HOUSE(s) AWB</h1>
		<table width=100% align=center border=1>
		<tr>
        	<th>Id</th>
			<th>Editar HAWB</th>
            <th>Embarcador</th>
			<th>Consignatario</th>
            <th>Agente</th>
			<th>Piezas</th>
			<th>Peso</th>
			<th>Routing</th>
            <!--<th>Contact</th>-->
            <% if AwbType = 1 then 'EXPORT %>
    		<th>Etiqueta</th>                
            <% end if %>   
    		<th>Manifiesto</th>
		</tr>
        
        <% 
        'OpenConn2 Conn2


        iConta = 0
        For i = 0 To CountList6Values-1

		    if aList6Values(4,i) <> "" AND aList6Values(3,i) <> aList6Values(4,i) then   

                iConta = iConta + 1	
		        
                if iConta = 1 and replica = "Master-Master-Hija" then
                    response.write "<tr style='background-color:yellow'>" 
                else
                    response.write "<tr>" 
                end if

                response.write "<td>" & aList6Values(0,i) & "</td>"
                 
	
				'valida la master

				if IsNull(rst("ConsignerData")) and IsNull(rst("ShipperData")) and IsNull(rst("AgentData")) and replica = "Master-Master-Hija" then
                
                response.write "<td nowrap>" & _            
                        "<input type=text readonly value='" & aList6Values(4,i) & "' >&nbsp;&nbsp;" & _
                        "<button type=button class='Boton2 ' title='Edita House'  onclick=""alert('Completar datos de Master')""><img src='img/glyphicons_150_edit.png' /></button>" & _                        
                    "</td>" 		

                else
                

		        response.write "<td nowrap>" & _            
                        "<form name='frm_awb_frame04' action='InsertData.asp'  onsubmit='move();'>" & _
                        "<input type='hidden' name='replica' value='" & replica & "' />" & _
                        "<input type='hidden' name='awb_frame2' value='1' />" & _
					    "<input type='hidden' name='AWBNumber2' value='" & aList6Values(3,i) & "' />" & _
                        "<input type='hidden' name='HAWBNumber2' value='" & aList6Values(4,i) & "' />" & _
                        "<input type='hidden' name='Routing2' value='" & aList6Values(9,i) & "' />" & _		                        
					    "<input type='hidden' name='Country2' value='" & Country2 & "' />" & _
					    "<input type='hidden' name='Transportista2' value='" & aList6Values(10,i) & "' />" & _
					    "<input type='hidden' name='AirportDepID2' value='" & aList6Values(11,i) & "' />" & _
					    "<input type='hidden' name='AirportDesID2' value='" & aList6Values(12,i) & "' />" & _	
                        "<input type='hidden' name='GID' value='1' />" & _
                        "<input type='hidden' name='AT' value='" & AwbType & "' />" & _
                        "<input type='hidden' name='OID' value='" & aList6Values(0,i) & "' />" & _
                        "<input type='hidden' name='CD' value='" & aList6Values(1,i) & "' />" & _
                        "<input type='hidden' name='CT' value='" & aList6Values(2,i) & "' />" & _                                                
                        "<input type='hidden' name='No' value='" & iConta & "' />" & _                                                
                        "<input type='hidden' name='TipoCarga2' value='" & TipoCarga2 & "' />" & _                                                
                        "<input type=text readonly value='" & aList6Values(4,i) & "' >&nbsp;&nbsp;" & _
                        "<button type=submit class='Boton2 cOrag' title='Edita House'><img src='img/glyphicons_150_edit.png' /></button>" & _                        
                        "</form>" & _
                    "</td>" 		
                    
                end if


                    '"<input type=submit value='" & aList6Values(4,i) & "' title='Editar House " & aList6Values(0,i) & "' class='Boton cBLue'>" & _
		
                response.write "<td>" & trim(aList6Values(14,i)) & "</td>"
                response.write "<td>" & trim(aList6Values(6,i)) & "</td>"
                response.write "<td>" & trim(aList6Values(17,i)) & "</td>"
                response.write "<td align=right>" & trim(aList6Values(7,i)) & "</td>"
                response.write "<td align=right>" & trim(aList6Values(8,i)) & "</td>"                
                response.write "<td>" & aList6Values(9,i) & "</td>"
		        
                if AwbType = 1 then 'EXPORT 
                    response.write "<td align=center><a href='http://" & IIf(InStr(1,iIps,Request.ServerVariables("LOCAL_ADDR")) > 0,"localhost:2020","10.10.1.21:8181/admin") & "/Reports.asp?Action=4&OID=" & aList6Values(0,i) & "&AT=" & AwbType & "' target=_blank class='Boton2 cOrag' title='Etiqueta'><img src='img/glyphicons_389_new_window_alt.png' /></a></td>"
                end if 

                response.write "<td align=center><a href='http://" & IIf(InStr(1,iIps,Request.ServerVariables("LOCAL_ADDR")) > 0,"localhost:2020","10.10.1.21:8181/admin") & "/Reports.asp?Action=2&OID=" & aList6Values(0,i) & "&AT=" & AwbType & "' target=_blank class='Boton2 cOrag' title='Manifiesto Cargo'><img src='img/glyphicons_389_new_window_alt.png' /></a></td>"

                response.write "</tr>"
		
                CarrierRates = CarrierRates + aList6Values(7,i)
                CarrierSubTot = CarrierSubTot + aList6Values(8,i)

		    end if
            		
   		Next
        'CloseOBJ Conn2
	    %>
	
	    <tr>
		    <td></td>
		    <td></td>
            <th>Totales</td>
		    <th align=right><%=CarrierRates%></th>
		    <th align=right><%=CarrierSubTot%></th>
		    <td></td>
	    </tr>
	    </table>
	    </td>
    </tr>

    <% end if %>

</table>



<%
CloseOBJ Conn  
Response.End
%>

