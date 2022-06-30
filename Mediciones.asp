<%    
    Dim medicion, diaUno, diaDos, diaTres, diaCuatro, diaCinco, horauno, horados, horatres, horacuatro, horacinco, MedicionUno, MedicionDos

    Dim Uno, Dos, Tres, Cuatro, Cinco, Segmentos, LabelUnoMain, LabelUnoDate, LabelUnoTime, LabelDosMain, LabelDosDate, LabelDosTime, LabelTresMain, LabelTresDate, LabelTresTime, LabelCuatroMain, LabelCuatroDate, LabelCuatroTime, LabelCincoMain, LabelCincoDate, LabelCincoTime, LabelUnoMedicion, LabelDosMedicion, hh, mm, ss

    Uno = "1"
    Dos = "2" 
    Tres = "3"
    Cuatro = "4" 
    Cinco = "5"

    hh = "hh"
    mm = "mm"
    ss = "ss"


    if CountTableValues >= 0 then

        medicion = aTableValues(0, 0)
    	AWBID = aTableValues(3, 0)
        diaUno = aTableValues(7, 0)
        horauno = aTableValues(8, 0) 
        diaDos = aTableValues(9, 0)
        horados = aTableValues(10, 0)
        diaTres = aTableValues(11, 0)
        horatres = aTableValues(12, 0)
        diaCuatro = aTableValues(13, 0)
        horacuatro = aTableValues(14, 0)

        diaCinco = aTableValues(25, 0)
        horacinco = aTableValues(26, 0)

        MedicionUno = aTableValues(15, 0)
        MedicionDos = aTableValues(16, 0)
        Action = 2

        HAWBNumber = aTableValues(5, 0)
		TotNoOfPieces = aTableValues(17, 0)
		TotWeight = aTableValues(18, 0)            
        Destinity = aTableValues(19, 0)    

    else 

        medicion = 0

		AWBID = ObjectID 
        Action = 1    
    end if

    if AWBID > 0 then
        openConn Conn 'aereo

        QuerySelect = "select AWBID, CreatedDate, CreatedTime, Expired, " & _
						"AWBNumber, AccountShipperNo, ShipperData, AccountConsignerNo, " & _
						"ConsignerData, AgentData, AccountInformation, IATANo, AccountAgentNo, AirportDepID, RequestedRouting, " & _
						"AirportToCode1, CarrierID, AirportToCode2, AirportToCode3, CarrierCode2, CarrierCode3, CurrencyID, " & _
						"ChargeType, ValChargeType, OtherChargeType, DeclaredValue, AduanaValue, AirportDesID, FlightDate1, " & _
						"FlightDate2, SecuredValue, HandlingInformation, Observations, NoOfPieces, Weights, WeightsSymbol, " & _
						"Commodities, ChargeableWeights, CarrierRates, CarrierSubTot, NatureQtyGoods, TotNoOfPieces, TotWeight, " & _
						"TotCarrierRate, TotChargeWeightPrepaid, TotChargeWeightCollect, TotChargeValuePrepaid, TotChargeValueCollect, " & _
						"TotChargeTaxPrepaid, TotChargeTaxCollect, AnotherChargesAgentPrepaid, AnotherChargesAgentCollect, " & _
						"AnotherChargesCarrierPrepaid, AnotherChargesCarrierCollect, TotPrepaid, TotCollect, TerminalFee, CustomFee, " & _
						"FuelSurcharge, SecurityFee, PBA, TAX, AdditionalChargeName1, AdditionalChargeVal1, AdditionalChargeName2, " & _
						"AdditionalChargeVal2, Invoice, ExportLic, AgentContactSignature, CommoditiesTypes, TotWeightChargeable, " & _
						"Instructions, AgentSignature, AWBDate, " & _
						"AdditionalChargeName3, AdditionalChargeVal3, AdditionalChargeName4, AdditionalChargeVal4, Countries, HAWBNumber, " & _
						"AdditionalChargeName5, AdditionalChargeVal5, AdditionalChargeName6, AdditionalChargeVal6, " & _
						"DisplayNumber, AdditionalChargeName7, AdditionalChargeVal7, AdditionalChargeName8, AdditionalChargeVal8, WType, " & _
  						"AdditionalChargeName9, AdditionalChargeVal9, AdditionalChargeName10, AdditionalChargeVal10, ShipperID, ConsignerID, AgentID, SalespersonID, " & _
						"ShipperAddrID, ConsignerAddrID, AgentAddrID, " & _
						"AdditionalChargeName11, AdditionalChargeVal11, AdditionalChargeName12, AdditionalChargeVal12, AdditionalChargeName13, AdditionalChargeVal13, " & _
						"AdditionalChargeName14, AdditionalChargeVal14, AdditionalChargeName15, AdditionalChargeVal15, Voyage, PickUp, Intermodal, SedFilingFee, " & _
                        "CalcAdminFee, Routing, RoutingID, CTX, TCTX, TPTX, Closed, ConsignerColoader, ShipperColoader, AgentNeutral, ManifestNo, Destinity "
        if AwbType = 1 then
			QuerySelect = QuerySelect & " from Awb "
		else
			QuerySelect = QuerySelect & " from Awbi"
		end if	
            
        QuerySelect = QuerySelect & " WHERE AWBID = " & AWBID

        'response.write (QuerySelect)
            
		Set rs = Conn.Execute(QuerySelect)
		If Not rs.EOF Then
			aTableValues = rs.GetRows
			CountTableValues = rs.RecordCount
		End If
		closeOBJ rs

	    if CountTableValues >= 0 then

            CreatedDate = aTableValues(1, 0)
            CreatedTime = aTableValues(2, 0)

            Countries = aTableValues(78, 0)
            CarrierID = aTableValues(16, 0)        
            AWBNumber = aTableValues(4, 0)
            HAWBNumber = aTableValues(79, 0)
            Routing = aTableValues(116, 0)
            AccountShipperNo = aTableValues(5, 0)
	
		    ShipperData = aTableValues(6, 0)
		    AccountConsignerNo = aTableValues(7, 0)
		    ConsignerData = aTableValues(8, 0)
		    AgentData = aTableValues(9, 0)

		    TotNoOfPieces = aTableValues(41, 0)
		    TotWeight = aTableValues(42, 0)            
            Destinity = aTableValues(126, 0)            
            
            'Obteniendo listado de Carriers
	        Set rs = Conn.Execute("select CarrierID, Name, Countries from Carriers where Expired = 0 and Countries in " & Session("Countries") & " order by Name, Countries")
	        If Not rs.EOF Then
   		        aList1Values = rs.GetRows
       	        CountList1Values = rs.RecordCount
            End If
	        CloseOBJ rs

        end if




        'si pais no tiene registro de etiquetas, toma las default que son guatemala
        SQLQuery = "SELECT * FROM mediciones_labels WHERE countries IN  ('" & Countries & "','GT') AND AwbType = " & AwbType & " ORDER BY id DESC LIMIT 1"
        'response.write SQLQuery
        set rs = Conn.Execute(SQLQuery)        
        if Not rs.EOF then
            Segmentos = rs(3)
	        LabelUnoMain = rs(5)
            LabelUnoDate = rs(6)
            LabelUnoTime = rs(7)

	        LabelDosMain = rs(9)
            LabelDosDate = rs(10)
            LabelDosTime = rs(11)

	        LabelTresMain = rs(13)
            LabelTresDate = rs(14)
            LabelTresTime = rs(15)

	        LabelCuatroMain = rs(17)
            LabelCuatroDate = rs(18)
            LabelCuatroTime = rs(19)

            LabelCincoMain = rs(21)
            LabelCincoDate = rs(22)
            LabelCincoTime = rs(23)

            LabelUnoMedicion = rs(4)
            LabelDosMedicion = rs(8)
	    end if
	    CloseOBJ rs


        CloseOBJ Conn

    end if
    
    For i = 0 To CountList1Values-1
    	if aList1Values(0,i) = CarrierID then
	    	CarrierName = aList1Values(1,i)
    		'Countries = aList1Values(2,i)
	    end if
    next







%>
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<!--<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>-->
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>

<LINK REL="stylesheet" type="text/css" HREF="img/2016.css">

<script language="JavaScript">
    function abrir(Label) {
        var DateSend, Subject;
        if (parseInt(navigator.appVersion) < 5) {
            DateSend = document.forma(Label).value;
        } else {
            DateSend = document.getElementById(Label).value;
        }
        Subject = '';
        window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject, 'Seleccionar', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
    }


    function validar(Action) {

        var continua = false;

        if (document.forma.DateTo1) { 
            if (document.forma.DateTo1.value == '') { alert('NO PUEDE IR VACIO "<%=LabelUnoDate%>"'); return false; };        
            if (document.forma.DateTo1.value != '<%=diaUno%>') { continua = true; };
        }

        //if (document.forma.DateTo2.value == '') { alert('<%=LabelDosDate%>'); return false; };        
        if (document.forma.DateTo2)
        if (document.forma.DateTo2.value != '<%=diaDos%>') { continua = true; };
        
        //if (document.forma.DateTo3.value == '') { alert('<%=LabelTresDate%>'); return false; };        
        if (document.forma.DateTo3) 
        if (document.forma.DateTo3.value != '<%=diaTres%>') { continua = true; };
        
        //if (document.forma.DateTo4.value == '') { alert('<%=LabelCuatroDate%>'); return false; };        
        if (document.forma.DateTo4)
        if (document.forma.DateTo4.value != '<%=diaCuatro%>') { continua = true; };

        if (document.forma.DateTo5)
        if (document.forma.DateTo5.value != '<%=diaCinco%>') { continua = true; };
        
        if (document.forma.Comodin.value == '1') continua = true;

        if (continua == false) {
            alert('No hay cambios');
            return false;
        };

        document.forma.Action.value = Action;
        document.forma.submit();
    }

    function ChangeSelIniciada() {
        document.forma.submit();
    }

    function CommodynFire() {
        document.forma.Comodin.value = 1;
    }

</SCRIPT>

<!DOCTYPE html>

    

    <table class="GridView" width="96%" align="center">
        <thead><tr><th colspan=8>DATOS GENERALES DEL LA GUIA AEREA <%=medicion%> :: <%=AWBID%></th></tr></thead>
        <tbody>            
            <tr>
            <td rowspan=2 align=center><h2><% If AwbType = 1 Then %> EXPORT <% Else %> IMPORT <% End If %></h2></td>
            <th>Fecha </th><th>Pais</th><th>AWBNumber</th><th>HAWBNumber</th><th>Transportista</th> <th>ROUTING</th><th>No. de Cuenta del Embarcador</th></tr>
            <tr><td><%=CreatedDate & " " & CreatedTime %></td><td><%=Countries%></td><td><%=AWBNumber%></td>   <td><%=HAWBNumber%></td><td><%=CarrierName%></td><td><%=Routing%></td><td><%=AccountShipperNo%></td></tr>            
            <!--
            <tr><th>Fecha </th><td><%=CreatedDate & " " & CreatedTime %></td></tr>
            <tr><th>Transportista</th><td><%=CarrierName%></td></tr>
            <tr><th>No. de Cuenta del Embarcador</th><td><%=AccountShipperNo%></td></tr>
            <tr><th>ROUTING</th><td><%=Routing%></td></tr>
            <tr><th colspan=2>Nombre y Direccion del Embarcador</th>    <td colspan=4><%=ShipperData%></td></tr>
            <tr><th colspan=2>Nombre y Direccion del Destinatario </th> <td colspan=4><%=ConsignerData%></td></tr>
            <tr><th colspan=2>Nombre y Direccion del CoLoader</th>      <td colspan=4><%=ColoaderData%></td></tr>
            <tr><th colspan=2>Agente del Transportista Emisor, Nombre y Ciudad </th><td colspan=4><%=AgentData%></td></tr>
            -->
        </tbody>
    </table>


    <table class="GridView" width="96%" align="center">
        <tbody>
            <tr><th colspan=2>Nombre y Direccion del Embarcador</th>    <td colspan=4><%=ShipperData%></td></tr>
            <tr><th colspan=2>Nombre y Direccion del Destinatario </th> <td colspan=4><%=ConsignerData%></td></tr>
            <tr><th colspan=2>Nombre y Direccion del CoLoader</th>      <td colspan=4><%=ColoaderData%></td></tr>
            <tr><th colspan=2>Agente del Transportista Emisor, Nombre y Ciudad </th><td colspan=4><%=AgentData%></td></tr>
        </tbody>
    </table>

    <BR />
    <%
    
    '                       0           1           2           3       4           5           6       7       8           9       10      11          12      13          14          15          16              17              18      19          20          21              22          23          24                          
    'QuerySelect = "SELECT MedicionID, CreatedDate, CreatedTime, AwbID, AwbNumber, HAwbNumber, AwbType, DateUno, TimeUno, DateDos, TimeDos, DateTres, TimeTres, DateCuatro, TimeCuatro, MedicionUno, MedicionDos, TotNoOfPieces, TotWeight, Destinity, ShipperData, UserInsert, UserUpdate, DateUpdate, Status FROM "

        %>



<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT type=hidden name="Action" value=0>
	<INPUT type=hidden name="SO" value="<%=SearchOption%>">
	<INPUT type=hidden name="GID" value="<%=GroupID%>">
	<INPUT type=hidden name="OID" value="<%=ObjectID%>">
    <INPUT type=hidden name="AT" value="<%=AwbType%>">
	<!--<INPUT name="CD" value="<%=CreatedDate%>">
	<INPUT name="CT" value="<%=CreatedTime%>">-->
    <INPUT type=hidden name="AwbNumber" value="<%=AwbNumber%>">
    <INPUT type=hidden name="HAwbNumber" value="<%=HAwbNumber%>">   
    <INPUT type=hidden name="TotNoOfPieces" value="<%=TotNoOfPieces%>">
    <INPUT type=hidden name="TotWeight" value="<%=TotWeight%>">
    <INPUT type=hidden name="Destinity" value="<%=Destinity%>">
    <INPUT type=hidden name="ShipperData" value="<%=ShipperData%>">
    <INPUT type=hidden name="Comodin" value="">

    <style>    
    .segmento { color:rgb(255,0,0); font-weight:bolder; font-size:22; position:absolute; display:block; }        
    .LineTop { border:1px solid blue;border-bottom:0px; }
    .LineLeft { border-left:1px solid blue; }
    .LineRight { border-right:1px solid blue; }
    .LineBot { border:1px solid blue;border-top:0px; }
    .tbl_seg { display:inline-block; width:250px; vertical-align:top; margin:5px;}
    .spc { background-color:white}
    </style>

    <%



Function Segmento(Uno1,LabelMain,LabelDate,LabelTime,uno,dia) 

    Segmento = "<table border=0 align=center width='250px'  class='GridView tbl_seg' cellpadding=3 cellspacing=0>" & _
    "<thead>" & _
        "<tr>" & _            
		    "<th colspan='2' class='LineTop'><span class='segmento'>" & Uno1 & "</span>" & LabelMain & "</th>" & _            
        "</tr>" & _
        "<tr>" & _            
		    "<td class='small LineLeft' height=80>" & LabelDate & "</th>" & _
		    "<td class='small LineRight' height=80>" & LabelTime & "</th>" & _            
        "</tr>" & _
        "<tr>" & _            
		    "<td class='LineLeft'><INPUT type=button value='Seleccionar' onClick=JavaScript:" & "abrir('DateTo" & uno & "'); class=label></td>" & _
		    "<td class='LineRight'><table width=100% ><thead><tr><th>" & hh & "</th><th>" & mm & "</th><th>" & ss & "</th></tr></thead></table></td>" & _            
        "</tr>" & _
        "<tr>" & _            
		    "<td class='LineLeft'><INPUT readonly='readonly' name='DateTo" & uno & "' id='DateTo" & uno & "' type=text value='" & dia & "' size=14 maxLength=14 class=label id='" & LabelDate & "'></td>" & _
		    "<td class='LineRight' nowrap>" & _
                "<select onchange=CommodynFire() name='Hrs" & uno & "'>" 
                For i = 0 To 23    	   
                    Segmento = Segmento & "<option value='" & TwoDigits(i) & "'>" & TwoDigits(i) & "</option>"
                next  
                Segmento = Segmento & "</select>" & _
                "<select onchange=CommodynFire() name='Min" & uno & "'>" 
                For i = 0 To 59    	   
                    Segmento = Segmento & "<option value='" & TwoDigits(i) & "'>" & TwoDigits(i) & "</option>"
                next  
                Segmento = Segmento & "</select>" & _
                "<select onchange=CommodynFire() name='Sec" & uno & "'>" 
                For i = 0 To 59    	   
                    Segmento = Segmento & "<option value='" & TwoDigits(i) & "'>" & TwoDigits(i) & "</option>"
                next  
                Segmento = Segmento & "</select>" & _
            "</td>" & _            
        "</tr>" & _
        "<tr>" & _            
            "<td class='LineBot' colspan=2>&nbsp;</td>" & _              
        "</tr>" & _
    "</thead>" & _
	"</table>" 

End Function



    If InStr(1,Segmentos, "1") > 0 Then
        response.write Segmento(Uno,LabelUnoMain,LabelUnoDate,LabelUnoTime,"1",diaUno) 
    End If

    If InStr(1,Segmentos, "2") > 0 Then
        response.write Segmento(Dos,LabelDosMain,LabelDosDate,LabelDosTime,"2",diaDos) 
    End If

    If InStr(1,Segmentos, "3") > 0 Then
        response.write Segmento(Tres,LabelTresMain,LabelTresDate,LabelTresTime,"3",diaTres) 
    End If

    If InStr(1,Segmentos, "4") > 0 Then
        response.write Segmento(Cuatro,LabelCuatroMain,LabelCuatroDate,LabelCuatroTime,"4",diaCuatro) 
    End If

    If InStr(1,Segmentos, "5") > 0 Then
        response.write Segmento(Cinco,LabelCincoMain,LabelCincoDate,LabelCincoTime,"5",diaCinco) 
    End If
    

    %>


	
    <table border="0" align="center" width="250px" class="GridView tbl_seg" cellpadding="3" cellspacing="0">
    <thead>
        <tr>
            <td colspan=2 class="LineTop" style="height:1px"></td>
        </tr>
        <tr>
            <th class="LineLeft"><%=LabelUnoMedicion%><span class="segmento"><%=Dos%> - <%=Uno%></span></th><td class="LineRight"><%=MedicionUno%></td>
        </tr>
        <tr>
            <td colspan=2 class="LineBot"></td>
        </tr>
    </thead>
	</table>


    <table border="0" align="center" width="250px" class="GridView tbl_seg" cellpadding="3" cellspacing="0">
    <thead>
        <tr>
            <td colspan=2 class="LineTop" style="height:1px"></td>
        </tr>
        <tr>
            <th class="LineLeft"><%=LabelDosMedicion%><span class="segmento"><%=Cuatro%> - <%=Tres%></span></th><td class="LineRight"><%=MedicionDos%></td>
        </tr>
        <tr>
            <td colspan=2 class="LineBot"></td>
        </tr>
    </thead>
	</table>



  
    <% if Action = 1 then %>
    <INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class="label">
    <% end if %>

    <% if Action = 2 then %>
    <INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class="label">
    <% end if %>
       
    <br />

</form>



<FORM name="forma2" action="Search_ResultsAdmin.asp" method="post">
    <input type=submit value="<< Regresar"/>
    <INPUT name="GID" type=hidden value="<%=GroupID%>">
  	<INPUT name="Action" type=hidden value=1>
	<INPUT name="P" type=hidden value=1>
    <INPUT name="AwbType" type=hidden value="<%=AwbType%>">
	<INPUT name="DateTo" type=hidden value="<%=Session("DateTo")%>">    
    <INPUT name="DateFrom" type=hidden value="<%=Session("DateFrom")%>">	
</FORM>

<script>

    var TimeUno = '<%=horauno%>';
    var TimeDos = '<%=horados%>';
    var TimeTres = '<%=horatres%>';
    var TimeCuatro = '<%=horacuatro%>';
    var TimeCinco = '<%=horacinco%>';
    //HH:MM:SS
    //console.clear();
    //console.log(TimeDos.substr(0, 2));
    //console.log(TimeDos.substr(3, 2));
    //console.log(TimeDos.substr(6, 2));

    if (document.forma.Hrs1) document.forma.Hrs1.value = TimeUno.substr(0, 2);
    if (document.forma.Min1) document.forma.Min1.value = TimeUno.substr(3, 2);
    if (document.forma.Sec1) document.forma.Sec1.value = TimeUno.substr(6, 2);

    if (document.forma.Hrs2) document.forma.Hrs2.value = TimeDos.substr(0, 2);
    if (document.forma.Min2) document.forma.Min2.value = TimeDos.substr(3, 2);
    if (document.forma.Sec2) document.forma.Sec2.value = TimeDos.substr(6, 2);

    if (document.forma.Hrs3) document.forma.Hrs3.value = TimeTres.substr(0, 2);
    if (document.forma.Min3) document.forma.Min3.value = TimeTres.substr(3, 2);
    if (document.forma.Sec3) document.forma.Sec3.value = TimeTres.substr(6, 2);

    if (document.forma.Hrs4) document.forma.Hrs4.value = TimeCuatro.substr(0, 2);
    if (document.forma.Min4) document.forma.Min4.value = TimeCuatro.substr(3, 2);
    if (document.forma.Sec4) document.forma.Sec4.value = TimeCuatro.substr(6, 2);

    if (document.forma.Hrs5) document.forma.Hrs5.value = TimeCinco.substr(0, 2);
    if (document.forma.Min5) document.forma.Min5.value = TimeCinco.substr(3, 2);
    if (document.forma.Sec5) document.forma.Sec5.value = TimeCinco.substr(6, 2);

</script>

