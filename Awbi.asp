<%
Checking "0|1|2"

Session.LCID = 4106

Response.CharSet = "utf-8"

On Error Resume Next






GuidesAWBNumber = false
CantItems = -1
    
    'response.write "(" & CheckNum(Request("awb_frame2")) & ")"
    'response.write "(" & Request("No") &  ")<br>"

TipoCarga2 = Request("TipoCarga2")
replica = Request("replica")
No = Request("No")
    
    if Session("PBAValue") = "" then
        Session("PBAValue") = "0"
    end if


    if Request("BtnBorra") <> "" then
       %> <!--#include file=awb_new.asp--> <%           'viene de awb_new.asp 
    end if

    if Request("BtnEdita") <> "" or Request("BtnCancela") <> "" then
       %> <!--#include file=awb_frame.asp--> <%         'viene de awb_frame.asp 
    end if

	Select Case CheckNum(Request("awb_frame2"))
	Case 1                                              'viene de awb_frame
        'response.write Request("Country2") & " Carrier=" & _
        'Request("Transportista2") & " AiportDEP=" & _        
        'Request("AirportDepID2") & " AirportDES=" & _
        'Request("AirportDesID2") & " " & _        
        'Request("BtnMaster2") & " " & _        
        'Request("AWBNumber2") & " " & _
        'Request("HAWBNumber2") & " " & _
        'Request("BtnReplica2") & " "              

	Case 2                                              
        %> <!--#include file=awb_frame.asp--> <%        'viene de Utils.asp DisplaySearchAdminResults

	Case 3                                              
        %> <!--#include file=awb_new.asp--> <%          'viene de Menu.asp 

    case else
        'response.write "**" & Request("Country2") & " Carrier=" & _
        'Request("Transportista2") & " AiportDEP=" & _        
        'Request("AirportDepID2") & " AirportDES=" & _
        'Request("AirportDesID2") & " " & _        
        'Request("BtnMaster2") & " " & _        
        'Request("AWBNumber2") & " " & _
        'Request("HAWBNumber2") & " " & _
        'Request("BtnReplica2") & " "            
	End Select

    
if Action <> 3 and Action <> 4 then

	if CountTableValues >= 0 and Request.Form("CallRouting") = "" then 'CallRouting lo que hace es que cargue todos los datos proporcionados por el ro
	
		CreatedDate = aTableValues(1, 0)
		CreatedTime = aTableValues(2, 0)
		Expired = aTableValues(3, 0)
		'AWBNumber = aTableValues(4, 0)
		AccountShipperNo = aTableValues(5, 0)
		ShipperData = aTableValues(6, 0)
		AccountConsignerNo = aTableValues(7, 0)
		ConsignerData = aTableValues(8, 0)
		AgentData = aTableValues(9, 0)
		AccountInformation = aTableValues(10, 0)
		IATANo = aTableValues(11, 0)
		AccountAgentNo = aTableValues(12, 0)
		AirportDepID = aTableValues(13, 0)
		RequestedRouting = aTableValues(14, 0)
		AirportToCode1 = aTableValues(15, 0)
		
		AirportToCode2 = aTableValues(17, 0)
		AirportToCode3 = aTableValues(18, 0)
		CarrierCode2 = aTableValues(19, 0)
		CarrierCode3 = aTableValues(20, 0)
		CurrencyID = aTableValues(21, 0)
		ChargeType = aTableValues(22, 0)
		ValChargeType = aTableValues(23, 0)
		OtherChargeType = aTableValues(24, 0)
		DeclaredValue = aTableValues(25, 0)
		AduanaValue = aTableValues(26, 0)
		AirportDesID = aTableValues(27, 0)
		FlightDate1 = aTableValues(28, 0)
		FlightDate2 = aTableValues(29, 0)
		SecuredValue = aTableValues(30, 0)
		HandlingInformation = aTableValues(31, 0)
		Observations = aTableValues(32, 0)
		NoOfPieces = aTableValues(33, 0)
		Weights = aTableValues(34, 0)
		WeightsSymbol = aTableValues(35, 0)
		Commodities = aTableValues(36, 0)
		ChargeableWeights = aTableValues(37, 0)
		CarrierRates = aTableValues(38, 0)
		CarrierSubTot = aTableValues(39, 0)
		NatureQtyGoods = aTableValues(40, 0)
		TotNoOfPieces = aTableValues(41, 0)
		TotWeight = aTableValues(42, 0)
		TotCarrierRate = aTableValues(43, 0)
		TotChargeWeightPrepaid = aTableValues(44, 0)
		TotChargeWeightCollect = aTableValues(45, 0)
		TotChargeValuePrepaid = aTableValues(46, 0)
		TotChargeValueCollect = aTableValues(47, 0)
		TotChargeTaxPrepaid = aTableValues(48, 0)
		TotChargeTaxCollect = aTableValues(49, 0)
		AnotherChargesAgentPrepaid = aTableValues(50, 0)
		AnotherChargesAgentCollect = aTableValues(51, 0)
		AnotherChargesCarrierPrepaid = aTableValues(52, 0)
		AnotherChargesCarrierCollect = aTableValues(53, 0)
		TotPrepaid = aTableValues(54, 0)
		TotCollect = aTableValues(55, 0)
		TerminalFee = aTableValues(56, 0)
		CustomFee = aTableValues(57, 0)
		FuelSurcharge = aTableValues(58, 0)
		SecurityFee = aTableValues(59, 0)
		PBA = aTableValues(60, 0)
		TAX = CheckNum(aTableValues(61, 0))
		AdditionalChargeName1 = aTableValues(62, 0)
		AdditionalChargeVal1 = aTableValues(63, 0)
		AdditionalChargeName2 = aTableValues(64, 0)
		AdditionalChargeVal2 = aTableValues(65, 0)
		Invoice = aTableValues(66, 0)
		ExportLic = aTableValues(67, 0)
		AgentContactSignature = aTableValues(68, 0)
		CommoditiesTypes = aTableValues(69, 0)
		TotWeightChargeable = aTableValues(70, 0)
		Instructions = aTableValues(71, 0)
		AgentSignature = aTableValues(72, 0)
		AWBDate = aTableValues(73, 0)
		AdditionalChargeName3 = aTableValues(74, 0)
		AdditionalChargeVal3 = aTableValues(75, 0)
		AdditionalChargeName4 = aTableValues(76, 0)
		AdditionalChargeVal4 = aTableValues(77, 0)
		Countries = aTableValues(78, 0)

        if Countries = "GT" then

            'if Request.Form("AWBNumber") <> "#**#" then
            '    AWBNumber = aTableValues(4, 0)            
            'end if    
            'if CInt(CarrierID) = CInt(Request.Form("CarrierID")) and Request.Form("AWBNumber") = "#**#" then        
            '    AWBNumber = aTableValues(4, 0)            
            'end if        
            
            AWBNumberAnt = aTableValues(4, 0)            
            HAWBNumberAnt = aTableValues(79, 0)
            CarrierIDAnt = aTableValues(16, 0)
        end if 

        CarrierID = aTableValues(16, 0)        
        AWBNumber = aTableValues(4, 0)
        HAWBNumber = aTableValues(79, 0)

		AdditionalChargeName5 = aTableValues(80, 0)
		AdditionalChargeVal5 = aTableValues(81, 0)
		AdditionalChargeName6 = aTableValues(82, 0)
		AdditionalChargeVal6 = aTableValues(83, 0)
		DisplayNumber = aTableValues(84, 0)
		AdditionalChargeName7 = aTableValues(85, 0)
		AdditionalChargeVal7 = aTableValues(86, 0)
		AdditionalChargeName8 = aTableValues(87, 0)
		AdditionalChargeVal8 = aTableValues(88, 0)
		WType = aTableValues(89, 0)
		AdditionalChargeName9 = aTableValues(90, 0)
		AdditionalChargeVal9 = aTableValues(91, 0)
		AdditionalChargeName10 = aTableValues(92, 0)
		AdditionalChargeVal10 = aTableValues(93, 0)
		ShipperID = aTableValues(94, 0)
		ConsignerID = aTableValues(95, 0)
		AgentID = aTableValues(96, 0)
		SalespersonID = aTableValues(97, 0)
		ShipperAddrID = aTableValues(98, 0)
		ConsignerAddrID = aTableValues(99, 0)
		AgentAddrID = aTableValues(100, 0)
		AdditionalChargeName11 = aTableValues(101, 0)
		AdditionalChargeVal11 = aTableValues(102, 0)
		AdditionalChargeName12 = aTableValues(103, 0)
		AdditionalChargeVal12 = aTableValues(104, 0)
		AdditionalChargeName13 = aTableValues(105, 0)
		AdditionalChargeVal13 = aTableValues(106, 0)
		AdditionalChargeName14 = aTableValues(107, 0)
		AdditionalChargeVal14 = aTableValues(108, 0)
		AdditionalChargeName15 = aTableValues(109, 0)
		AdditionalChargeVal15 = aTableValues(110, 0)
		Voyage = aTableValues(111, 0)
		PickUp = aTableValues(112, 0)
		Intermodal = aTableValues(113, 0)
		SedFilingFee = aTableValues(114, 0)
		CalcAdminFee = aTableValues(115, 0)
		Routing = aTableValues(116, 0)
		RoutingID = aTableValues(117, 0)
		CTX = aTableValues(118, 0)
		TCTX = aTableValues(119, 0)
		TPTX = aTableValues(120, 0)
		Closed = aTableValues(121, 0)
        ConsignerColoader = aTableValues(122, 0)
        ShipperColoader = aTableValues(123, 0)
        AgentNeutral = aTableValues(124, 0)
        ManifestNumber = aTableValues(125, 0)

        'id_coloader = aTableValues(126, 0)
		'TotCarrierRate_Routing = aTableValues(127, 0)
		'FuelSurcharge_Routing = aTableValues(128, 0)
		'SecurityFee_Routing = aTableValues(129, 0)
		'CustomFee_Routing = aTableValues(130, 0)
		'TerminalFee_Routing = aTableValues(131, 0)
		'PickUp_Routing = aTableValues(132, 0)
		'SedFilingFee_Routing = aTableValues(133, 0)
		'Intermodal_Routing = aTableValues(134, 0)
		'PBA_Routing = aTableValues(135, 0)
		'TAX_Routing = aTableValues(136, 0)
		'AdditionalChargeName1_Routing = aTableValues(137, 0)
		'AdditionalChargeName2_Routing = aTableValues(138, 0)
		'AdditionalChargeName3_Routing = aTableValues(139, 0)
		'AdditionalChargeName4_Routing = aTableValues(140, 0)
		'AdditionalChargeName5_Routing = aTableValues(141, 0)
		'AdditionalChargeName6_Routing = aTableValues(142, 0)
		'AdditionalChargeName7_Routing = aTableValues(143, 0)
		'AdditionalChargeName8_Routing = aTableValues(144, 0)
		'AdditionalChargeName9_Routing = aTableValues(145, 0)
		'AdditionalChargeName10_Routing = aTableValues(146, 0)
		'AdditionalChargeName11_Routing = aTableValues(147, 0)
		'AdditionalChargeName12_Routing = aTableValues(148, 0)
		'AdditionalChargeName13_Routing = aTableValues(149, 0)
		'AdditionalChargeName14_Routing = aTableValues(150, 0)
		'AdditionalChargeName15_Routing = aTableValues(151, 0)		

        OtherChargeName1 = aTableValues(126, 0)
		OtherChargeName2 = aTableValues(127, 0)
		OtherChargeName3 = aTableValues(128, 0)	
		OtherChargeName4 = aTableValues(129, 0)	
		OtherChargeName5 = aTableValues(130, 0)	
		OtherChargeName6 = aTableValues(131, 0)	
		OtherChargeVal1 = aTableValues(132, 0)	
		OtherChargeVal2 = aTableValues(133, 0)	
		OtherChargeVal3 = aTableValues(134, 0)	
		OtherChargeVal4 = aTableValues(135, 0)	
		OtherChargeVal5 = aTableValues(136, 0)	
		OtherChargeVal6 = aTableValues(137, 0)	
        id_coloader = aTableValues(138, 0)	
		TotCarrierRate_Routing = aTableValues(139, 0)
		FuelSurcharge_Routing = aTableValues(140, 0)
		SecurityFee_Routing = aTableValues(141, 0)
		PickUp_Routing = aTableValues(142, 0)
		SedFilingFee_Routing = aTableValues(143, 0)
		Intermodal_Routing = aTableValues(144, 0)
		PBA_Routing = aTableValues(145, 0)
		AdditionalChargeName1_Routing = aTableValues(146, 0)
		AdditionalChargeName2_Routing = aTableValues(147, 0)
		AdditionalChargeName3_Routing = aTableValues(148, 0)
		AdditionalChargeName4_Routing = aTableValues(149, 0)
		AdditionalChargeName5_Routing = aTableValues(150, 0)
		AdditionalChargeName6_Routing = aTableValues(151, 0)
		AdditionalChargeName7_Routing = aTableValues(152, 0)
		AdditionalChargeName8_Routing = aTableValues(153, 0)
		AdditionalChargeName9_Routing = aTableValues(154, 0)
		AdditionalChargeName10_Routing = aTableValues(155, 0)
		AdditionalChargeName11_Routing = aTableValues(156, 0)
		AdditionalChargeName12_Routing = aTableValues(157, 0)
		AdditionalChargeName13_Routing = aTableValues(158, 0)
		AdditionalChargeName14_Routing = aTableValues(159, 0)
		AdditionalChargeName15_Routing = aTableValues(160, 0)
		OtherChargeName1_Routing = aTableValues(161, 0)
		OtherChargeName2_Routing = aTableValues(162, 0)
		OtherChargeName3_Routing = aTableValues(163, 0)
		OtherChargeName4_Routing = aTableValues(164, 0)
		OtherChargeName5_Routing = aTableValues(165, 0)
		OtherChargeName6_Routing = aTableValues(166, 0)		

        id_cliente_order = aTableValues(167, 0)
        id_cliente_orderData = aTableValues(168, 0)
        ReplicaAwbID = CheckNum(aTableValues(169, 0))
        flg_master = aTableValues(170, 0)
        flg_totals = aTableValues(171, 0)
        file = aTableValues(172, 0)        
        
        OpenConn2 Conn2
        SQLQuery = "select seguro, poliza_seguro, routing_seg, routing_adu, routing_ter from routings where id_routing=" & RoutingID

        SQLQuery = "SELECT seguro, poliza_seguro, routing_seg, routing_adu, routing_ter, COALESCE(a.id_facturar,0), COALESCE(b.nombre_cliente,'') FROM routings a LEFT JOIN clientes b ON b.id_cliente = a.id_facturar WHERE a.id_routing = " & RoutingID

        'response.write SQLQuery
        Set rs = Conn2.Execute(SQLQuery )
        if Not rs.EOF then
            Seguro = rs(0)
			routing_seg = rs(2)
            routing_adu = rs(3)
            routing_ter = rs(4)
            'facturar_a = rs(5)
            'facturar_a_nombre = rs(6)        
        end if
        CloseOBJs rs, Conn2
		
        'response.write "(1)(" & aTableValues(157, 0) & ")"

        'iMinimo = aTableValues(157, 0)

        'OpenConn2 Conn
        SQLQuery = "select tarifa_minimo from Awb_Columns where AwbId=" & aTableValues(0, 0) & " and DocTyp = '" & IIf(AWBType = 1,"0","1") & "'"
        'response.write SQLQuery
        Set rs = Conn.Execute(SQLQuery )
        if Not rs.EOF then
            iMinimo = rs(0)
        end if
        'CloseOBJs rs, Conn
        CloseOBJ rs
		   
	else

        'response.write "(2)(" & Request.Form("iMinimo") & ")"
        
        iMinimo = Request.Form("iMinimo")

		'response.write( "(" & RoutingID & ")" )

		RoutingID = CheckNum(Request.Form("RoutingID"))
	
        Seguro = CheckNum(Request.Form("Seguro"))
		routing_seg = CheckNum(Request.Form("routing_seg"))
        routing_adu = CheckNum(Request.Form("routing_adu"))
        routing_ter = CheckNum(Request.Form("routing_ter"))

		'AWBNumber = Request.Form("AWBNumber")
		'if AWBNumber = "" then
		'	AWBNumber = Session("AWBNumber")
		'end if
		AccountShipperNo = Request.Form("AccountShipperNo")
		ShipperData = Request.Form("ShipperData")
		AccountConsignerNo = Request.Form("AccountConsignerNo")
		ConsignerData = Request.Form("ConsignerData")
		AgentData = Request.Form("AgentData")
		AccountInformation = Request.Form("AccountInformation")
		if AccountInformation = "" then
			AccountInformation = Session("Xchange")
		end if
		IATANo = Request.Form("IATANo")
		AccountAgentNo = Request.Form("AccountAgentNo")
		
        if Request("AirportDepID2") <> "" then 
            AirportDepID = Request("AirportDepID2")
        else
            AirportDepID = CheckNum(Request.Form("AirportDepID"))
        end if

		RequestedRouting = Request.Form("RequestedRouting")
		AirportToCode1 = Request.Form("AirportToCode1")        
		AirportToCode2 = Request.Form("AirportToCode2")
		AirportToCode3 = Request.Form("AirportToCode3")
		CarrierCode2 = Request.Form("CarrierCode2")
		CarrierCode3 = Request.Form("CarrierCode3")
		CurrencyID = CheckNum(Request.Form("CurrencyID"))
		ChargeType = Request.Form("ChargeType")
		ValChargeType = Request.Form("ValChargeType")
		OtherChargeType = Request.Form("OtherChargeType")
		DeclaredValue = Request.Form("DeclaredValue")
		AduanaValue = Request.Form("AduanaValue")

        if Request("AirportDesID2") <> "" then 
            AirportDesID = Request("AirportDesID2")
        else
    		AirportDesID = CheckNum(Request.Form("AirportDesID"))            
        end if

		FlightDate1 = Request.Form("FlightDate1")
		FlightDate2 = Request.Form("FlightDate2")
		SecuredValue = Request.Form("SecuredValue")
		HandlingInformation = Request.Form("HandlingInformation")
		Observations = Request.Form("Observations")
		NoOfPieces = Request.Form("NoOfPieces")
		Weights = Request.Form("Weights")
		WeightsSymbol = Request.Form("WeightsSymbol")
		Commodities = Request.Form("Commodities")
		ChargeableWeights = Request.Form("ChargeableWeights")
		CarrierRates = Request.Form("CarrierRates")
		CarrierSubTot = Request.Form("CarrierSubTot")
		NatureQtyGoods = Request.Form("NatureQtyGoods")
		TotNoOfPieces = Request.Form("TotNoOfPieces")
		TotWeight = Request.Form("TotWeight")
		
		TotChargeWeightPrepaid = Request.Form("TotChargeWeightPrepaid")
		TotChargeWeightCollect = Request.Form("TotChargeWeightCollect")
		TotChargeValuePrepaid = Request.Form("TotChargeValuePrepaid")
		TotChargeValueCollect = Request.Form("TotChargeValueCollect")
		TotChargeTaxPrepaid = Request.Form("TotChargeTaxPrepaid")
		TotChargeTaxCollect = Request.Form("TotChargeTaxCollect")
		AnotherChargesAgentPrepaid = Request.Form("AnotherChargesAgentPrepaid")
		AnotherChargesAgentCollect = Request.Form("AnotherChargesAgentCollect")
		AnotherChargesCarrierPrepaid = Request.Form("AnotherChargesCarrierPrepaid")
		AnotherChargesCarrierCollect = Request.Form("AnotherChargesCarrierCollect")
		TotPrepaid = Request.Form("TotPrepaid")
		TotCollect = Request.Form("TotCollect")
		
		
		TAX = CheckNum(Request.Form("TAX"))

		
		
		
		
		
		
		AdditionalChargeVal1 = Request.Form("AdditionalChargeVal1")
		AdditionalChargeVal2 = Request.Form("AdditionalChargeVal2")
		AdditionalChargeVal3 = Request.Form("AdditionalChargeVal3")		
		AdditionalChargeVal4 = Request.Form("AdditionalChargeVal4")		
		AdditionalChargeVal5 = Request.Form("AdditionalChargeVal5")			
		AdditionalChargeVal6 = Request.Form("AdditionalChargeVal6")
		AdditionalChargeVal7 = Request.Form("AdditionalChargeVal7")		
		AdditionalChargeVal8 = Request.Form("AdditionalChargeVal8")
		AdditionalChargeVal9 = Request.Form("AdditionalChargeVal9")		
		AdditionalChargeVal10 = Request.Form("AdditionalChargeVal10")
		AdditionalChargeVal11 = Request.Form("AdditionalChargeVal11")		
		AdditionalChargeVal12 = Request.Form("AdditionalChargeVal12")		
		AdditionalChargeVal13 = Request.Form("AdditionalChargeVal13")		
		AdditionalChargeVal14 = Request.Form("AdditionalChargeVal14")		
		AdditionalChargeVal15 = Request.Form("AdditionalChargeVal15")
		
		
        AdditionalChargeName1 = Request.Form("AdditionalChargeName1")
		AdditionalChargeName1 = Replace(AdditionalChargeName1 , ",", ".")

        AdditionalChargeName2 = Request.Form("AdditionalChargeName2")
		AdditionalChargeName2 = Replace(AdditionalChargeName2 , ",", ".")

		AdditionalChargeName3 = Request.Form("AdditionalChargeName3")
		AdditionalChargeName3 = Replace(AdditionalChargeName3 , ",", ".")

		AdditionalChargeName4 = Request.Form("AdditionalChargeName4")
		AdditionalChargeName4 = Replace(AdditionalChargeName4 , ",", ".")

		AdditionalChargeName5 = Request.Form("AdditionalChargeName5")
		AdditionalChargeName5 = Replace(AdditionalChargeName5 , ",", ".")

		AdditionalChargeName6 = Request.Form("AdditionalChargeName6")
		AdditionalChargeName6 = Replace(AdditionalChargeName6 , ",", ".")
		
		AdditionalChargeName7 = Request.Form("AdditionalChargeName7")
		AdditionalChargeName7 = Replace(AdditionalChargeName7 , ",", ".")

		AdditionalChargeName8 = Request.Form("AdditionalChargeName8")
		AdditionalChargeName8 = Replace(AdditionalChargeName8 , ",", ".")

		AdditionalChargeName9 = Request.Form("AdditionalChargeName9")
		AdditionalChargeName9 = Replace(AdditionalChargeName9 , ",", ".")

		AdditionalChargeName10 = Request.Form("AdditionalChargeName10")
		AdditionalChargeName10 = Replace(AdditionalChargeName10 , ",", ".")

		AdditionalChargeName11 = Request.Form("AdditionalChargeName11")
		AdditionalChargeName11 = Replace(AdditionalChargeName11 , ",", ".")

		AdditionalChargeName12 = Request.Form("AdditionalChargeName12")
		AdditionalChargeName12 = Replace(AdditionalChargeName12 , ",", ".")

		AdditionalChargeName13 = Request.Form("AdditionalChargeName13")
		AdditionalChargeName13 = Replace(AdditionalChargeName13 , ",", ".")

		AdditionalChargeName14 = Request.Form("AdditionalChargeName14")
		AdditionalChargeName14 = Replace(AdditionalChargeName14 , ",", ".")

		AdditionalChargeName15 = Request.Form("AdditionalChargeName15")
		AdditionalChargeName15 = Replace(AdditionalChargeName15 , ",", ".")
		
        OtherChargeVal1 = Request.Form("OtherChargeVal1")
		OtherChargeVal2 = Request.Form("OtherChargeVal2")
		OtherChargeVal3 = Request.Form("OtherChargeVal3")
		OtherChargeVal4 = Request.Form("OtherChargeVal4")
		OtherChargeVal5 = Request.Form("OtherChargeVal5")
		OtherChargeVal6 = Request.Form("OtherChargeVal6")				
		OtherChargeName1 = Request.Form("OtherChargeName1")
		OtherChargeName2 = Request.Form("OtherChargeName2")
		OtherChargeName3 = Request.Form("OtherChargeName3")
		OtherChargeName4 = Request.Form("OtherChargeName4")
		OtherChargeName5 = Request.Form("OtherChargeName5")
		OtherChargeName6 = Request.Form("OtherChargeName6")

		TotCarrierRate = Request.Form("TotCarrierRate")
        'TotCarrierRate = Replace(TotCarrierRate , ",", ".")

		FuelSurcharge = Request.Form("FuelSurcharge")
        'FuelSurcharge = Replace(FuelSurcharge , ",", ".")

		SecurityFee = Request.Form("SecurityFee")
        'SecurityFee = Replace(SecurityFee , ",", ".")

		PickUp = Request.Form("PickUp")
        'PickUp = Replace(PickUp , ",", ".")

		SedFilingFee = Request.Form("SedFilingFee")
        'SedFilingFee = Replace(SedFilingFee , ",", ".")

		Intermodal = Request.Form("Intermodal")
        'Intermodal = Replace(Intermodal , ",", ".")

		PBA = Request.Form("PBA")
        'PBA = Replace(PBA , ",", ".") 
        
		TerminalFee = Request.Form("TerminalFee")
        'TerminalFee = Replace(TerminalFee , ",", ".")

		CustomFee = Request.Form("CustomFee")
        'CustomFee = Replace(CustomFee , ",", ".")

		if PBA = "" then
			PBA = Session("PBAValue")
		end if
		
		Invoice = Request.Form("Invoice")
		ExportLic = Request.Form("ExportLic")
		AgentContactSignature = Request.Form("AgentContactSignature")
		CommoditiesTypes = Request.Form("CommoditiesTypes")
		TotWeightChargeable = Request.Form("TotWeightChargeable")
		Instructions = Request.Form("Instructions")
		AgentSignature = Request.Form("AgentSignature")
		AWBDate = Request.Form("AWBDate")			
		Countries = Request.Form("Countries")        

    
        'if Request("Countries2") <> "" then
        '    Countries = Request("Countries2")
        'end if

        
        if Request("Country2") <> "" then
            Countries = Request("Country2")
        end if        

        if InStr(1,Countries,"GT",1) > 0 then
            'if Request.Form("AWBNumber") = "#**#" then
            '    AWBNumber = ""
            'else
            '    AWBNumber = Request.Form("AWBNumber")            
            'end if 		
            AWBNumberAnt = Request.Form("AWBNumberAnt")            
            HAWBNumberAnt = Request.Form("HAWBNumberAnt")
            CarrierIDAnt = Request.Form("CarrierIDAnt")

            if HandlingInformation = "" then
                HandlingInformation = "*** ENVELOPE WITH DOCUMENTS ORIGINALS ATTACHED *** *** PLS, NOTIFY TO CONSIGNEE UPON ARRIVAL ***"
            end if
        end if



            if AgentSignature = "" then

                AgentSignature = "AIMAR"

                if InStr(1,Countries,"TLA",1) > 0 then
                    AgentSignature = "GRUPO TLA"
                end if

                 if InStr(1,Countries,"LTF",1) > 0 then
                    AgentSignature = "LATIN FREIGHT"
                end if

                if InStr(1,Countries,"N1",1) > 0 then
                    AgentSignature = "GRH"
                end if

            end if

            if AgentContactSignature = "" then
                AgentContactSignature = Session("OperatorName")
            end if

            if AWBDate = "" then
                AWBDate = Date
            end if

            'response.write "(" & Request("iAirportFromCode") & ")(" & Request("iAirportToCode") & ")"


            if AirportToCode1 = "" then
                AirportToCode1 = Request("iAirportToCode")
            end if

            if RequestedRouting = "" then
                RequestedRouting = Request("iAirportFromCode") & "/" & Request("iAirportToCode")
            end if           



        

        if Request("AWBNumber2") <> "" then
            AWBNumber = Request("AWBNumber2") 
            
            if Request("BtnMaster2") <> "" then 'es master

                if Request("BtnReplica2") = "Directo" then 'es master
                    HAWBNumber = AWBNumber
                else
                    HAWBNumber = ""
                end if
                 
            else
                HAWBNumber = ""
            end if

        else
            AWBNumber = Request.Form("AWBNumber")
            HAWBNumber = Request.Form("HAWBNumber")
        end if

        if Request("HAWBNumber2") <> "" then
            HAWBNumber = Request("HAWBNumber2") 
        end if
		

        if Request.Form("CarrierID") = "" then
            CarrierID = -1
        else
		    CarrierID = CheckNum(Request.Form("CarrierID"))
        end if


        if Request("Transportista2") <> "" then
            CarrierID = Request("Transportista2")
        end if

		
		DisplayNumber = Request.Form("DisplayNumber")


		
		WType = Request.Form("WType")
		
		ShipperID = CheckNum(Request.Form("ShipperID"))
		ConsignerID = CheckNum(Request.Form("ConsignerID"))
		AgentID = CheckNum(Request.Form("AgentID"))
		SalespersonID = CheckNum(Request.Form("SalespersonID"))
		ShipperAddrID = CheckNum(Request.Form("ShipperAddrID"))
		ConsignerAddrID = CheckNum(Request.Form("ConsignerAddrID"))
		AgentAddrID = CheckNum(Request.Form("AgentAddrID"))
		
		Voyage = Request.Form("Voyage")
		
		
		
		CalcAdminFee = Request.Form("CalcAdminFee")
		Routing = Request.Form("Routing")
		
		CTX = Request.Form("CTX")
		TCTX = Request.Form("TCTX")
		TPTX = Request.Form("TPTX")
		Closed = Request.Form("Closed")
        ConsignerColoader = Request.Form("ConsignerColoader")
        ShipperColoader = Request.Form("ShipperColoader")
        AgentNeutral = Request.Form("AgentNeutral")
        ManifestNumber = Request.Form("ManifestNumber")
        ClientCollectID = Request.Form("ClientCollectID")
        ClientsCollect = Request.Form("ClientsCollect")
        
		'ItemCurrs = Request.Form("ItemCurrs")
        'ItemIDs = Request.Form("ItemIDs")
        'ItemVals = Request.Form("ItemVals")
        'ItemLocs = Request.Form("ItemLocs")
        'ItemNames = Request.Form("ItemNames")
		'ItemNames_Routing = Request.Form("ItemNames_Routing")
        'ItemOVals = Request.Form("ItemOVals")
        'ItemPPCCs = Request.Form("ItemPPCCs")
        'ItemServIDs = Request.Form("ItemServIDs")
        'ItemServNames = Request.Form("ItemServNames")
        'ItemInvoices = Request.Form("ItemInvoices")
        'ItemCalcInBls = Request.Form("ItemCalcInBls")
        'ItemIntercompanyIDs = Request.Form("ItemIntercompanyIDs")
        'if Request.Form("CantItems")="" then
        '    CantItems = -1
		'else
		'	CantItems = CheckNum(Request.Form("CantItems"))
        'end if
		
        id_coloader = Request.Form("id_coloader")
		
		
		TotCarrierRate_Routing  =  Request.Form("TotCarrierRate_Routing")
		FuelSurcharge_Routing  =  Request.Form("FuelSurcharge_Routing")
		SecurityFee_Routing  =  Request.Form("SecurityFee_Routing")
		CustomFee_Routing  =  Request.Form("CustomFee_Routing")
		TerminalFee_Routing  =  Request.Form("TerminalFee_Routing")
		PickUp_Routing  =  Request.Form("PickUp_Routing")
		SedFilingFee_Routing  =  Request.Form("SedFilingFee_Routing")
		Intermodal_Routing  =  Request.Form("Intermodal_Routing")
		PBA_Routing  =  Request.Form("PBA_Routing")
		TAX_Routing  =  Request.Form("TAX_Routing")
		AdditionalChargeName1_Routing  =  Request.Form("AdditionalChargeName1_Routing")
		AdditionalChargeName2_Routing  =  Request.Form("AdditionalChargeName2_Routing")
		AdditionalChargeName3_Routing  =  Request.Form("AdditionalChargeName3_Routing")
		AdditionalChargeName4_Routing  =  Request.Form("AdditionalChargeName4_Routing")
		AdditionalChargeName5_Routing  =  Request.Form("AdditionalChargeName5_Routing")
		AdditionalChargeName6_Routing  =  Request.Form("AdditionalChargeName6_Routing")
		AdditionalChargeName7_Routing  =  Request.Form("AdditionalChargeName7_Routing")
		AdditionalChargeName8_Routing  =  Request.Form("AdditionalChargeName8_Routing")
		AdditionalChargeName9_Routing  =  Request.Form("AdditionalChargeName9_Routing")
		AdditionalChargeName10_Routing  =  Request.Form("AdditionalChargeName10_Routing")
		AdditionalChargeName11_Routing  =  Request.Form("AdditionalChargeName11_Routing")
		AdditionalChargeName12_Routing  =  Request.Form("AdditionalChargeName12_Routing")
		AdditionalChargeName13_Routing  =  Request.Form("AdditionalChargeName13_Routing")
		AdditionalChargeName14_Routing  =  Request.Form("AdditionalChargeName14_Routing")
		AdditionalChargeName15_Routing  =  Request.Form("AdditionalChargeName15_Routing")

        id_cliente_order =  Request.Form("id_cliente_order")
        id_cliente_orderData =  Request.Form("id_cliente_orderData")


        'if Request("BtnReplica2") <> "" then
        '    replica = Request("BtnReplica2")
        'else
        '    replica = Request.Form("replica")
        'end if
		
        file = Request.Form("file")

	end if





	if Action = 0 And Request.Form("CantItems") <> "" then
			
        ItemCurrs = Request.Form("ItemCurrs")
        ItemIDs = Request.Form("ItemIDs")
        ItemVals = Request.Form("ItemVals")
        ItemLocs = Request.Form("ItemLocs")
        ItemNames = Request.Form("ItemNames")
		ItemNames_Routing = Request.Form("ItemNames_Routing")
        ItemOVals = Request.Form("ItemOVals")
        ItemPPCCs = Request.Form("ItemPPCCs")
        ItemServIDs = Request.Form("ItemServIDs")
        ItemServNames = Request.Form("ItemServNames")
        ItemInvoices = Request.Form("ItemInvoices")
        ItemCalcInBls = Request.Form("ItemCalcInBls")
        ItemIntercompanyIDs = Request.Form("ItemIntercompanyIDs")        

        if Request.Form("CantItems")="" then
            CantItems = -1
		else
			CantItems = CheckNum(Request.Form("CantItems"))
        end if

	End If
	
end if




    if iMinimo = "" then
        iMinimo = "1"
    end if

    'response.write "(4)(" & iMinimo & ")"



TaxRate = 0
OpenConn Conn

    
    'response.write "(" & Request("replica") &  ")<br>"
    'QuerySelect = "SELECT aiee_AwbID_fk, aiee_ImpExp, aiee_TipoAwb, aiee_replica, aiee_master_hija FROM Awb_IE_Expansion WHERE aiee_AwbID_fk = " & ObjectID & " AND aiee_ImpExp = " & AwbType    
    ''response.write QuerySelect & "<br>"
    ''response.write "(" & Session("Pricing") & ")(" & Countries & ")" & "<br>"
    'Set rs = Conn.Execute(QuerySelect)
    'If Not rs.EOF Then 
    '    If InStr(1, Session("Pricing"), Countries) = 0 Then
    '        replica = rs("aiee_replica") 'Consolidado / Directo
    '    Else
    '        replica = rs("aiee_TipoAwb") 'Master-Hija / Hija-Directa / Master-Master-Hija 
    '    End If
    'end if 

    if Action=1 or Action=2 then        
		'SaveChargeItems Conn, ObjectID, IIf(AWBType = 1,"0","1") 2022-06-10

        '/////////////////////////////////BLOQUE PARA RECALCUL O DE PESO EN RUBROS 2022-05-03
        'response.write "ChargeableWeights(" & ChargeableWeights & ") Peso2Anterior(" & Peso2Anterior & ")<br>"
        'if CDbl(ChargeableWeights) <> CDbl(Peso2Anterior) and CDbl(Peso2Anterior) > 0 then 'si modifican el peso     'Weights
        '    SQLQuery = "UPDATE ChargeItems SET Value = cast(TarifaPricing as decimal(11,2))*" & Cdbl(ChargeableWeights) & " WHERE AWBID = " & ObjectID & " AND Expired = 0 AND DocTyp = " & IIf(AwbType = 1,0,1) & " AND cast(TarifaPricing as decimal(11,2)) > 0"
        '    response.write SQLQuery & "<br>"
        '    Conn.Execute(SQLQuery)
        'end if
        '////////////////////////////////////////////////////////////////////7

        'Agregando Rubros Intercompany, la primera vez que se carga el RO
        if CheckNum(Request("CantItems"))>=0 then
		
            'Agregando Cliente a Colectar en Destino									
			'esto se tuvo que hacer porque no habia forma que pudiera evaluar el if ? 2016-03-31
			ClientCollectID_tmp = CheckNum(Request("ClientCollectID")) 
			if ClientCollectID_tmp <> "0" then			
				QuerySelect = "Update " & IIf(AwbType = 1,"Awb","Awbi") & " set ClientCollectID=" & ClientCollectID_tmp & ", ClientsCollect='" & Request("ClientsCollect") & "' where AWBID=" & ObjectID				
				Conn.Execute(QuerySelect)			
			end if
			
            'if CheckNum(Request("ClientCollectID"))<>0 then
            '    Conn.Execute("Update Awb set ClientCollectID=" & CheckNum(Request("ClientCollectID")) & ", ClientsCollect='" & Request("ClientsCollect") & "' where AWBID=" & ObjectID)
            'end if
            'el tercer parametro =0=Export            			
            'BAWResult = SaveInterChargeItems (Conn, ObjectID, 0, Countries)	2022-04-27 lo que es baw se comenta
        end if


        'response.write "(" & AWBNumber & ")(" & HAWBNumber & ")(" & esMaster & ")(" & ObjectID & ")(" & replica & ")<br>"

        if replica = "Master-Master-Hija" or replica = "Master-Hija" then '2022-04-27 se crea la nueva funcion

            'replica el cambio de peso solo cuando es la master y recalcula tambien para la M-H
            'response.write "(Peso=" & IIf(CDbl(ChargeableWeights) <> CDbl(Peso2Anterior) and CDbl(Peso2Anterior) > 0 and esMaster = True, CDbl(ChargeableWeights), 0) & ")<br>"
            
            '                                                                                   nada            nada             esHija Peso                      
            ReplicarHeaderRubros Conn, rs, esMaster, AwbType, AWBNumber, HAwbNumber, ObjectID, ObjectIDtmp, ClientCollectID_tmp, False, 0 'IIf(CDbl(ChargeableWeights) <> CDbl(Peso2Anterior) and CDbl(Peso2Anterior) > 0 and esMaster = True, CDbl(ChargeableWeights), 0)

        end if
    
	end if
    
    Peso2Anterior = CDbl(ChargeableWeights) '2022-05-03

    'response.write "ChargeableWeights(" & ChargeableWeights & ") Peso2Anterior(" & Peso2Anterior & ")<br>"



    'Obteniendo listado de Carriers
	Set rs = Conn.Execute("select CarrierID, Name, Countries from Carriers where Expired = 0 and Countries in " & Session("Countries") & " order by Countries, Name")
	If Not rs.EOF Then
   		aList1Values = rs.GetRows
       	CountList1Values = rs.RecordCount
    End If
	CloseOBJ rs

    'if CarrierID = -1 then Countries = "" end if

    '/////////////////TOMA EL PAIS DE LA LINEA AEREA///////////////////////////////////////////////////////////////////////////
 	'For i = 0 To CountList1Values-1 2016-12-07 se comento debe ser del usuario combinado con el ro
	'	if aList1Values(0,i) = CarrierID then			    
	'		Countries = aList1Values(2,i)                
	'	end if
   	'Next
    
    if trim(Countries) = "" then
        a=Split(Session("Countries"),",")
        b = Replace(a(0),"'","")
        b = Replace(b,"(","")
        Countries = Trim(Replace(b,")",""))
        'response.write b 'Request.Form("CarrierID") & ")(" & CarrierID & ")"
    end if



    if Request("Transportista2") = "" then

    'response.write(" CountTableValues=" & CountTableValues & " CarrierID=" & CarrierID &  "-" & CarrierIDAnt  )

    if Request.Form("BtnHouse") = "Nuevo" or Request.Form("BtnHouse") = "Ultimo" or Request.Form("BtnHouse") = "Espacio" then 
        TipoMaster = Request.Form("BtnHouse")
        'TipoHouse = "Asignar"
    else 
        if Request.Form("TipoMaster") = "" then
            TipoMaster = "Ultimo"
        else
            TipoMaster = Request.Form("TipoMaster")
        end if
    end if

    if Request.Form("BtnHouse") = "Asignar" or Request.Form("BtnHouse") = "Directo" or Request.Form("BtnHouse") = "Manual" then 
        TipoHouse = Request.Form("BtnHouse")
        'TipoMaster = "Ultimo"
    else             
        if Request.Form("TipoHouse") = "" then
            'TipoHouse = "Asignar"
        else
            'TipoHouse = Request.Form("TipoHouse")
        end if
    end if
    
    'response.write("TipoMaster=" & TipoMaster & " TipoHouse=" & TipoHouse)
    'response.write( "CarrierID=" & CarrierID &  " CarrierIDAnt=" & CarrierIDAnt & " BtnHouse=" & Request.Form("BtnHouse") & " Countries=" & Countries )
    
    if ((CarrierID <> CarrierIDAnt) or (Request.Form("BtnHouse") <> "")) and Countries = "GT" then        
        'OpenConn Conn
        'response.write "ENTRO"
        AWBNumber = NextAWBNumber(Conn, AwbType, CarrierID, TipoMaster)

        if AWBNumber <> "" then             
            GuidesAWBNumber = true
        end if                    

        if TipoHouse <> "" then
            HAWBNumber = NextHAWBNumber(HAWBNumber, Conn, AwbType, Countries, TipoMaster, TipoHouse, AWBNumber)
        end if
        
    end if      
    
    end if    

    '//////////////////////////////////////////////////////////////////////////////////////////////////////

    'response.write "(" & Request.Form("CarrierID") & ")(" & CarrierID & ")<br>"
	'if CheckNum(Request.Form("CarrierID")) > -1 then
	'	CarrierID = CheckNum(Request.Form("CarrierID"))
	'end if
    'response.write "(" & Request.Form("CarrierID") & ")(" & CarrierID & ")<br>"

    
	
	if CarrierID > -1 then

        'if RoutingID = 0 then '2018-04-19
		    'Obteniendo listado de Aeropuertos Salida asignados al Carrier
		    'QuerySelect = "select b.AirportID, b.Name, b.AirportCode, a.TerminalFeePD, a.TerminalFeeCS, a.CustomFee, a.SecurityFee, a.FuelSurcharge from CarrierDepartures a, Airports b where a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID =" & CarrierID & " order by b.Name"
		    'QuerySelect = "select b.AirportID, b.Name, b.AirportCode, a.TerminalFeePD, a.TerminalFeeCS, a.CustomFee, a.SecurityFee, a.FuelSurcharge from CarrierDepartures a, Airports b where a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID =" & CarrierID & " and b.AirportID = " & AirportDepID & " order by b.Name"
            'QuerySelect = "select b.AirportID, b.Name, b.AirportCode, " & IIf(TerminalFee = "","''","a.TerminalFeePD") & " as TerminalFeePD, " & IIf(TerminalFee = "","''","a.TerminalFeeCS") & " as TerminalFeeCS, " & IIf(CustomFee = "","''","a.CustomFee") & " as CustomFee, " & IIf(SecurityFee = "","''","a.SecurityFee") & " as SecurityFee, " & IIf(FuelSurcharge = "","''","a.FuelSurcharge") & " as FuelSurcharge from CarrierDepartures a, Airports b where a.AirportID = b.AirportID and b.Expired=0 order by b.Name"

            '2018-04-19
            'QuerySelect = "select b.AirportID, b.Name, b.AirportCode, " & IIf(TerminalFee = "","''","a.TerminalFeePD") & " as TerminalFeePD, " & IIf(TerminalFee = "","''","a.TerminalFeeCS") & " as TerminalFeeCS, " & IIf(CustomFee = "","''","a.CustomFee") & " as CustomFee, " & IIf(SecurityFee = "","''","a.SecurityFee") & " as SecurityFee, " & IIf(FuelSurcharge = "","''","a.FuelSurcharge") & " as FuelSurcharge from CarrierDepartures a, Airports b where a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID =" & CarrierID & " and b.AirportID = " & AirportDepID & " order by b.Name"

		
        OpenConn3 Conn

        QuerySelect = "SELECT ""tpp_pk"", ""tpp_nombre"", ""tpp_codigo"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
        'response.write QuerySelect & "<br>"
        Set rs = Conn.Execute(QuerySelect)
		If Not rs.EOF Then
    		aList2Values = rs.GetRows
        	CountList2Values = rs.RecordCount
	    End If
		CloseOBJ rs

        'end if '2018-04-19

		'Obteniendo listado de Aeropuertos Destino
        'QuerySelect = "select AirportID, Name, AirportCode from Airports where Expired=0 order by Name"
        QuerySelect = "SELECT ""tpp_pk"", ""tpp_nombre"", ""tpp_codigo"", ""tpp_pais_iso_fk"" FROM ""ti_pricing_puerto"" WHERE ""tpp_transporte_fk"" = '1' AND ""tpp_tps_fk"" = '1' AND ""tpp_codigo"" != ' ' AND ""tpp_nombre"" != ' ' ORDER BY ""tpp_nombre"", ""tpp_codigo"""
        'response.write QuerySelect & "<br>"
		Set rs = Conn.Execute(QuerySelect)
		If Not rs.EOF Then
    		aList3Values = rs.GetRows
        	CountList3Values = rs.RecordCount
	    End If
		CloseOBJ rs
        CloseOBJ Conn


        OpenConn Conn
		'Obteniendo listado de Rangos
        QuerySelect = "select a.RangeID, abs(b.Val) from CarrierRanges a, Ranges b where a.RangeID = b.RangeID and b.Expired = 0 and a.CarrierID =" & CarrierID & " order by b.Val"
        'response.write "CarrierRanges<br>"
        'response.write QuerySelect & "<br>"
		Set rs = Conn.Execute(QuerySelect)
        If Not rs.EOF Then
    		aList4Values = rs.GetRows
        	CountList4Values = rs.RecordCount
	    End If
		CloseOBJ rs

		'Obteniendo tarifas
        'QuerySelect = "select a.CarrierID, a.AirportDepID, a.AirportDesID, a.RangeID, a.Val from CarrierRates a where a.CarrierID = " & CarrierID & " order by CarrierID, AirportDepID, AirportDesID, RangeID"
        QuerySelect = "select a.CarrierID, a.AirportDepID, a.AirportDesID, a.RangeID, a.Val from CarrierRates a where a.CarrierID = " & CarrierID & " and a.AirportDepID = " & AirportDepID & " and a.AirportDesID = " & AirportDesID & " order by CarrierID, AirportDepID, AirportDesID, RangeID"        
        'response.write QuerySelect & "<br>"	
		Set rs = Conn.Execute(QuerySelect)
        If Not rs.EOF Then
    		aList5Values = rs.GetRows
 	   	   	CountList5Values = rs.RecordCount
		End If
		CloseOBJ rs


		'Obteniendo Monedas
		'Set rs = Conn.Execute("select a.CurrencyID, a.CurrencyCode, a.Xchange, a.Symbol from Currencies a where a.Expired=0 and a.Countries=(select b.Countries from Carriers b where b.CarrierID=" & CarrierID & ") order by a.CurrencyCode")
		Set rs = Conn.Execute("select a.CurrencyID, a.CurrencyCode, a.Xchange, a.Symbol from Currencies a, Carriers b where a.Expired=0 and a.Countries=b.Countries and b.CarrierID=" & CarrierID & " order by a.CurrencyCode")
		If Not rs.EOF Then
    		aList6Values = rs.GetRows
 	   	   	CountList6Values = rs.RecordCount
		End If

		CloseOBJ rs
		'Obteniendo Impuestos
		Set rs = Conn.Execute("select a.Tax from Taxes a, Carriers b where a.Countries=b.Countries and b.CarrierID=" & CarrierID)
		If Not rs.EOF Then
    		TaxRate = CheckNum(rs(0))
		End If
		CloseOBJ rs
        if Countries <> "GT" then
			if AWBNumber = "" then
                QuerySelect = "SELECT AWBNumber FROM " & IIf(AwbType = 1,"Awb","Awbi") & " WHERE AWBID = (SELECT MAX(AWBID) FROM " & IIf(AwbType = 1,"Awb","Awbi") & " WHERE AwbNumber <> '' AND CarrierID=" & CarrierID  & ")"
                'response.write(QuerySelect)
			    Set rs = Conn.Execute(QuerySelect)
				If Not rs.EOF Then
					AWBNumber = rs(0)
					'Reference = rs(1)
				else
					AWBNumber = ""
					'Reference = ""
				end if
				CloseOBJ rs
				'Reference = UpdateReference(Reference, AwbType)
			end if		
        end if
	End If

	'Obteniendo listado de Rubros
	CountList9Values = -1

    if Request.Form("CallRouting") = "" then '2018-04-23 no debe leer cuando se esta asignando el routing
                        '     0           1       2           3     4       5       6       7           8               9           10          11      12
        'QuerySelect = "Select ItemID, AgentTyp, CurrencyID, Local, Value, ItemName, Pos, ServiceID, ServiceName, PrepaidCollect, InvoiceID, CalcInBL, DocType from ChargeItems where Expired=0 and AwbID=" & ObjectID & " and DocTyp=0 order by AgentTyp"

                        '     0           1       2           3     4       5       6       7           8               9           10                                              11       12         13          14              15      16
        QuerySelect = "Select ItemID, AgentTyp, CurrencyID, Local, Value, ItemName, Pos, ServiceID, ServiceName, PrepaidCollect, CASE WHEN DocType = 9 THEN 0 ELSE InvoiceID END, CalcInBL, DocType, InvoiceID, TarifaPricing, Regimen, TarifaTipo from ChargeItems where Expired=0 and AwbID=" & ObjectID & " and DocTyp=" & IIf(AWBType = 1,"0","1") & " order by AgentTyp"

        'response.write QuerySelect 
        Set rs = Conn.Execute(QuerySelect)
        If Not rs.EOF Then
	        aList9Values = rs.GetRows
	        CountList9Values = rs.RecordCount - 1
        End If
    end if
	CloseOBJ rs
CloseOBJ Conn





    'response.write "(CarrierID=" & CarrierID & ")(CarrierIDAnt=" & CarrierIDAnt & ") &nbsp; "
    'response.write "(AWB=" & AWBNumber & ")(AWBAnt=" & AWBNumberAnt & ") &nbsp; "
    'response.write "(HAWB=" & HAWBNumber & ")(HAWBAnt=" & HAWBNumberAnt & ")<br>"


OpenConn2 Conn


    id_coloader = CheckNum(id_coloader) 

    if ColoaderData = "" and id_coloader > 0 then  
        QuerySelect = "select nombre_cliente, direccion_completa from clientes a, direcciones b where a.id_cliente = b.id_cliente and a.id_cliente = " & id_coloader & " and id_estatus in (1,2)"
        Set rs = Conn.Execute(QuerySelect)
	    If Not rs.EOF Then   		
       	    ColoaderData = rs(0) & Chr(13) & rs(1)
        End If    
	    CloseOBJ rs
    end if    

	'Obteniendo Productos
	'Set rs = Conn.Execute("select a.CommodityCode, cast(a.NameES as text), a.TypeVal, CommodityId from Commodities a where Expired=0 order by CommodityCode")
	'If Not rs.EOF Then
	'	aList7Values = rs.GetRows
	'	CountList7Values = rs.RecordCount
	'End If
	'CloseOBJ rs
	'Obteniendo Usuarios Vendedores
	CountList8Values = -1
	'Set rs = Conn.Execute("select u.id_usuario, u.nombre, u.id_pais from usuarios u, perfiles_usuarios p where u.id_usuario = p.id_usuario and p.id_perfil=4 and u.id_pais in " & Session("Countries") & " order by u.id_pais, u.nombre")
	Set rs = Conn.Execute("select id_usuario, pw_gecos, pais from usuarios_empresas where tipo_usuario=1 and pais in " & Session("Countries") & " order by pw_gecos, pais")
	If Not rs.EOF Then
		aList8Values = rs.GetRows
		CountList8Values = rs.RecordCount-1
	End If
	CloseOBJ rs
	'Obteniendo Monedas
	'Set rs = Conn.Execute("select moneda_id, pais, simbolo from monedas where pais in " & Session("Countries") & " order by pais, simbolo")
	'Do While Not rs.EOF
	'	Currencies = Currencies & "<option value=" & rs(0) & ">" & rs(1) & "-"  & rs(2) & "</option>"
	'	rs.MoveNext
	'Loop
	Set rs = Conn.Execute("select distinct simbolo from monedas where pais in " & Session("Countries") & " order by simbolo")
	Do While Not rs.EOF
		Currencies = Currencies & "<option value=" & rs(0) & ">" & rs(0) & "</option>"
		rs.MoveNext
	Loop

CloseOBJs rs, Conn

    Function IntLoc(num) 
		Select Case num 
		Case 0
			IntLoc = "INT"
		Case 1
			IntLoc = "LOC"
		Case Else
			IntLoc = "---"
		End Select
	End Function


    Function PrepColl(num) 
		Select Case num 
		Case 0
			PrepColl = "PREP"
		Case 1
			PrepColl = "COLL"
		Case Else
			PrepColl = "---"
		End Select
	End Function
    		

'response.write("(" & Request("master") & ")")
'response.write("(" & Request("house") & ")")
'response.write("(" & CheckNum(Request("OID")) & ")")
'response.write("(" & CheckNum(Request("CarrierID2")) & ")")
'response.write("(" & CheckNum(Request("AirportDepID2")) & ")")
'response.write("(" & CheckNum(Request("AirportDesID2")) & ")")

'if Request("master") <> "" and Request("house") = "" then
'    if CheckNum(Request("OID")) = 0 then
        'AWBNumber = Request("master")
'        CarrierID = CheckNum(Request("CarrierID2"))
	    'AirportDepID = CheckNum(Request("AirportDepID2"))
	    'AirportDesID = CheckNum(Request("AirportDesID2"))
'    end if
'end if


if CheckNum(Request("CarrierID2")) <> 0 then
    CarrierID = CheckNum(Request("CarrierID2"))
end if

'if Request("master") <> "" then
'    AWBNumber = Request("master")
'end if

if CheckNum(Request("AirportDepID2")) <> 0 then
    AirportDepID = CheckNum(Request("AirportDepID2"))
end if

if CheckNum(Request("AirportDesID2")) <> 0 then
    AirportDesID = CheckNum(Request("AirportDesID2"))
end if
%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<LINK REL="stylesheet" type="text/css" HREF="img/estilos.css">
<style type="text/css">
<!--
body {
	margin: 0px;
}
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: bold;
	color: #999999;
}
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}

.readonly { border:0px;
            background:silver;
            color:navy;
            font-size: 10px; 
            font-family: Verdana, Arial, Helvetica, sans-serif;  
            width:auto; }

.ids    {   border:0px;
            color:navy;
            font-weight:normal;
            background:silver;
            font-size: 8px;
            width:auto; 
            }


-->


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

.erpLab {
    color:white;background-color:gray;height:20px;display:block;padding:2px;
}

.erpFil {
    background-color:rgb(255,232,159);
}


</style>
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<script type="text/javascript">

function GetTarifaResponse(texto){

    var GetTarifaResponseStr = '';

    try {

        //console.log(texto);
        var myArray = texto.split('|');
        console.log(myArray);

        var theform = document.forms[0];
        var TotWeight = parseFloat(myArray[1])+0;
        var Tarifa = parseFloat(myArray[5])+0;
        var ItemTarifa = theform.elements[myArray[7]];
        var ItemTarifaHidden = theform.elements[myArray[8]];
        var ItemMonto = theform.elements[myArray[9]];
    
        //console.log('Resultados');
        //console.log(TotWeight + ' ' + Tarifa + ' ' + ItemTarifa.value + ' ' + ItemMonto.value);
        
        ItemTarifa.value = Tarifa;

        if (ItemTarifa.value > 0)
            ItemMonto.value = ItemTarifa.value * TotWeight;
        else
            ItemMonto.value = '';

        if (ItemTarifaHidden)
            ItemTarifaHidden.value = ItemMonto.value;

        //console.log(myArray[7] + ' ' + ItemTarifa.value + ' ' + ItemMonto.value);
        if (myArray[7] == 'AF_Tarifa') {
            theform.CarrierRates.value = ItemTarifa.value;
            theform.CarrierSubTot.value = ItemMonto.value;
        }

        if (myArray[3] == "0" && myArray[4] == "" && myArray[5] == "0" && myArray[6] == "") {
            var myArray = myArray[10].split('<br>');
            GetTarifaResponseStr = myArray[1];
        }

    } catch (error) {
      console.log('GetTarifaResponse Error------------');
      console.log(error);
      // expected output: ReferenceError: nonExistentFunction is not defined
      // Note - error messages will vary depending on browser
    }

    return GetTarifaResponseStr;
}

function TarifaRubro(tarifaOculta, tarifa, monto, peso, servicio, rubro) {


    var TarifaRubroStr = '';

    //if (parseFloat(document.forms[0].elements[tarifa].value+0) > 0) {
    if (document.forms[0].elements[tarifa+'Tipo'].value != '') {

        // Creating the XMLHttpRequest object
        var request = new XMLHttpRequest();

        var RoutingUrl = "OperAjax=GetTarifa&AwbType=<%=AwbType%>&Countries=<%=Countries%>&ObjectID=<%=ObjectID%>&ServiceID=" + servicio + "&ItemID=" + rubro + "&No=<%=No%>&ItemTarifa=" + tarifa + "&ItemTarifaHidden=" + tarifaOculta + "&ItemMonto=" + monto + "&peso=" + peso;

        console.log(RoutingUrl);

        // Instantiating the request object
        request.open("GET", "Utils.asp?" + RoutingUrl);

        // Defining event listener for readystatechange event
        request.onreadystatechange = function() {
            // Check if the request is compete and was successful
            if(this.readyState === 4 && this.status === 200) {
                
                TarifaRubroStr = GetTarifaResponse(this.responseText);

                if (TarifaRubroStr != '')
                    alert(TarifaRubroStr);

            } else {

                if(this.readyState != 2 && this.readyState != 3) {        
                    console.log('displayFullName ERROR:');
                    console.log(this);
                }        
            }
        };

        // Sending the request to the server
        request.send();
    }

    return TarifaRubroStr;

}


function RecalculoPeso(theForm) {

    var tio = 0;
    var msg1, msg2 = '';

    tio = '<%=IIf(CheckNum(No) < 2 and replica = "Master-Master-Hija",1,0)%>';
    
    if (tio == 0) return false; //solo realiza los calculos si es M ó MH y es modalidad MMH
     
    //var Peso2Anterior = <%=Peso2Anterior%>;

    var ChargeableWeights = theForm.ChargeableWeights.value;

    //if (parseFloat(Peso2Anterior) == parseFloat(ChargeableWeights)) return true;

    //Cargos
    msg1 = TarifaRubro('VAF', 'AF_Tarifa', 'TotCarrierRate', ChargeableWeights, 3, 11);    
    msg2 += msg1;

    console.log('MENSAJE');
    console.log(msg2);

    msg1 = TarifaRubro('VTF', 'TF_Tarifa', 'TerminalFee', ChargeableWeights, 3, 15);
    msg2 += msg1;
    msg1 = TarifaRubro('VCF', 'CF_Tarifa', 'CustomFee', ChargeableWeights, 3, 14);
    msg2 += msg1;
    msg1 = TarifaRubro('VFS', 'FS_Tarifa', 'FuelSurcharge', ChargeableWeights, 3, 12);
    msg2 += msg1;
    msg1 = TarifaRubro('VSF', 'SF_Tarifa', 'SecurityFee', ChargeableWeights, 3, 13);
    msg2 += msg1;
    msg1 = TarifaRubro('VPB', 'PB_Tarifa', 'PBA', ChargeableWeights, 3, 116);
    msg2 += msg1;
    msg1 = TarifaRubro('VPU', 'PU_Tarifa', 'PickUp', ChargeableWeights, 3, 31);
    msg2 += msg1;
    msg1 = TarifaRubro('VIM', 'IM_Tarifa', 'Intermodal', ChargeableWeights, 3, 115);
    msg2 += msg1;
    msg1 = TarifaRubro('VFF', 'FF_Tarifa', 'SedFilingFee', ChargeableWeights, 3, 38);
    msg2 += msg1;

	//Otros Cargos Agente
	for (i=1; i<=11;i++) {
        msg1 = TarifaRubro('VA'+i, 'A'+i+'_Tarifa', 'AdditionalChargeVal'+AgentsPos[i], ChargeableWeights, theForm.elements['SVIA'+i].value, theForm.elements['A'+i].value);
    msg2 += msg1;
	}

	//Otros Cargos Transportista
	for (i=1; i<=4;i++) {
        msg1 = TarifaRubro('VC'+i, 'C'+i+'_Tarifa', 'AdditionalChargeVal'+CarriersPos[i], ChargeableWeights, theForm.elements['SVIC'+i].value, theForm.elements['C'+i].value);
    msg2 += msg1;
	}

	//Otros Cargos
	for (i=1; i<=6;i++) {
        msg1 =TarifaRubro('VO'+i, 'O'+i+'_Tarifa', 'OtherChargeVal'+i, ChargeableWeights, theForm.elements['SVIO'+i].value, theForm.elements['O'+i].value);
    msg2 += msg1;
	}

    console.log(msg2);

    if (msg2 != '')
        alert(msg2);
}



    function move() {
        window.location = '#';
        document.forma.style.display = "none";
        document.getElementById('myBar').style.display = "block";
        var elem = document.getElementById("myBar");
        var width = 10;
        var id = setInterval(frame, 45);
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
	

    <%if BAWResult <> "" then %>
    alert("<%=BAWResult%>");
    <%end if %>

    <%if Action=1 or Action=2 then
        if HAWBNumber <> Request.Form("HAWBNumber") then%>
        alert('El numero de Guia House ya esta asignado a otro registro, no se puede repetir, favor coloque uno nuevo');
        <%end if
    end if %>

	function SetLabelID(Label) {
		var LabelID = "";
		if (Label == "AWBDate") {
			LabelID = "Fecha";
		}
		return LabelID;
	}

	function abrir(Label){
	var DateSend, Subject;
		if (parseInt(navigator.appVersion) < 5) {
			DateSend = document.forma(Label).value;
		} else {
			var LabelID = SetLabelID(Label);
			DateSend = document.getElementById(LabelID).value;
		}
		Subject = '';	
		window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
	}

	function LTrim(s){
		// Devuelve una cadena sin los espacios del principio
		var i=0;
		var j=0;
		
		// Busca el primer caracter <> de un espacio
		for(i=0; i<=s.length-1; i++)
			if(s.substring(i,i+1) != ' '){
				j=i;
				break;
			}
		return s.substring(j, s.length);
	}
	function RTrim(s){
		// Quita los espacios en blanco del final de la cadena
		var j=0;
		
		// Busca el último caracter <> de un espacio
		for(var i=s.length-1; i>-1; i--)
			if(s.substring(i,i+1) != ' '){
				j=i;
				break;
			}
		return s.substring(0, j+1);
	}
	function Trim(s){
		// Quita los espacios del principio y del final
		return LTrim(RTrim(s));
	}

	function SetFlag(Obj){
		if (Obj.value.length == 0) {
			return 0;
		} else {
			return 1;
		}		
	}






    ////////////////////////////////////////////AJAX 2018-02-01////////////////////////////////////////////////
    function GetXmlHttpObject(handler) {
        var objXmlHttp = null;

        if (navigator.userAgent.indexOf("MSIE") >= 0) {
            var strName = "Msxml2.XMLHTTP";
            if (navigator.appVersion.indexOf("MSIE 5.5") >= 0) {
                strName = "Microsoft.XMLHTTP";
            }
            try {
                objXmlHttp = new ActiveXObject(strName);
                objXmlHttp.onreadystatechange = handler;
                return objXmlHttp;
            }
            catch (e) {
                alert("Error. Scripting for ActiveX might be disabled");
                return;
            }
        }
        if (navigator.userAgent.indexOf("Mozilla") >= 0) {
            objXmlHttp = new XMLHttpRequest();
            objXmlHttp.onload = handler;
            objXmlHttp.onerror = handler;
            return objXmlHttp;
        }
    }

    function GetAjaxProcess(webpage, url) {
        xmlHttp = GetXmlHttpObject(stateChanged);
        xmlHttp.open("GET", webpage + '?' + url, true);
        xmlHttp.send(null);
    }
    /*
    function json_parser2(ResponseText){
            var a = ResponseText.split(',');
            var items;
            //var arraySelects = new Object();
            var arraySelects = new Array();
            for (var index = 0; index < a.length; ++index) {                
                items = a[index].split(':');                
                arraySelects[ items[0] ] = Number.isInteger(items[1]) ? parseInt(items[1]) : items[1];
            }            
            return arraySelects;    
    }

function json_parserObj(tar) {
    var obj = {};
    //var arr = new Object();
    tar = tar.replace(/^\{|\}$/g,'').split(',');
    for(var i=0,cur,pair;cur=tar[i];i++){
        pair = cur.split(':');
        obj[pair[0]] = /^\d*$/.test(pair[1]) ? +pair[1] : pair[1];
    }
    return obj;
}

function json_parserArr(tar) {
    //var arr = [];
    var arr = new Array();
    tar = tar.replace(/^\{|\}$/g,'').split(',');
    for(var i=0,cur,pair;cur=tar[i];i++){
        //arr[i] = {};
        pair = cur.split(':');
        arr[pair[0]] = /^\d*$/.test(pair[1]) ? +pair[1] : pair[1];
    }
    return arr;
}*/

    function stateChanged() {
        if (xmlHttp.readyState == 4 || xmlHttp.readyState == "complete") {

            //var myObject = json_parserArr(xmlHttp.responseText);
            var mystring = xmlHttp.responseText.split(',');
            var ierror,istatus,isql,imsg;
            ierror = parseInt(mystring[0]);
            istatus = parseInt(mystring[1]);
            isql = mystring[2];
            imsg = mystring[3];

            if (ierror > 0) {                
                alert('Error (1) : ' + ierror + '\n Query : ' + isql + '\n Message : ' + imsg);
                return false;
            } 

            switch (istatus) {
            case 0: //no realiza ninguna
                
                break;            

            case 1:
                move();
                document.forma.Action.value = 3;
                document.forma.submit();
                break;            
        
            default:
                alert('Error (2) : ' + ierror + '\n Query : ' + isql + '\n Message : ' + imsg);
            }
                                   
        } else {
            //alert('readyState:' + xmlHttp.readyState)
            console.log('readyState');
            console.log(xmlHttp);
        }
        return false;
    }
    ////////////////////////////////////////////AJAX 2018-02-01////////////////////////////////////////////////


    var child;
    var RoutingErrorSite = 'http' + '://10.10.1.20/reportar/routings_error.php'; 
    var RoutingErrorUrl = "OperAjax=GetRoutingErr&id_trafico=1&id_routing=" + <%=RoutingID%> + "&id_usuario=" + <%=Session("OperatorID")%>;

	function validar(Action) {

        if (Action != 3) {
            //esta parte se habilitara despues 2016-12-19
            if (Trim(document.forma.replica.value) == "") { alert('Seleccione Consolidado o Directo'); return false; }

            
          
        
            //if (!valTxt(document.forma.AWBNumber, 3, 5)){return (false)}
			//if (Trim(document.forma.HAWBNumber.value) != "") {
			//	if (!valTxt(document.forma.HAWBNumber, 3, 5)){return (false)}
			//}
         
			if (!valTxt(document.forma.AWBNumber, 3, 5)){ alert('Requiere numero de master'); return false }
            if (document.forma.TipoMaster == 'Ultimo') {
                if (document.forma.HAWBNumber.value.length == 0){ alert('Requiere numero de house'); return false }
            }
            //2016-04-18 se comenta porque aun no esta boletineado



            <% 
            couStr = "0"        
            if AWBNumber <> "" and HAWBNumber <> "" and replica = "Consolidado" then 'es house
                couStr = "1"
            end if
        
            if AWBNumber <> "" and HAWBNumber = AWBNumber and replica = "Directo" then 'es master          
                couStr = "1"
            end if 

            if ObjectID > 0 and  RoutingID > 0 then
                couStr = "0"
            end if
            %>

            <% if couStr = "1" then  %>
		    //if (TrimLen(document.forma.Routing.value) == 0) { alert('Requiere Routing'); document.forma.Routing.focus(); return false } //solicitado por Cesar solo para export 2016-02-09
            <% end if %>


            if (!valTxt(document.forma.ShipperData, 3, 9)){return (false)}
			if (!valTxt(document.forma.ConsignerData, 3, 9)){return (false)}
			if (!valTxt(document.forma.AgentData, 3, 9)){return (false)}
		
            //Validacion de Latin Freight y Aimar, el resto de empresas no tiene esta validacion, por ejemplo N1 (GRH)
            <%if FilterAimarLatin = 1 then%>
            if (document.forma.Countries.value!="N1") {
                if (document.forma.Countries.value.substr(2,3)=="LTF") {
                    if (document.forma.AgentNeutral.value == 0) {
                        alert("Para operaciones de Latin Freight solo puede utilizar agentes Neutrales");
				        document.forma.AgentData.focus();
                        return (false);
			        }
                } else {
                    var EconoCodes = /<%=PtrnEconoCodes%>/;
                    var Result = EconoCodes.exec(document.forma.AgentID.value)
                    if (Result == null) {
                        if (document.forma.ConsignerColoader.value == 1) {
                            alert("Para operaciones de Aimar, solo cuando el Agente es Econocaribe puede asignar Clientes o Shippers Coloaders, favor de consultar con su supervisor y revisar el administrador de catalogos");
				            document.forma.ConsignerData.focus();
                            return (false);
			            }
                        if (document.forma.ShipperColoader.value == 1) {
                            alert("Para operaciones de Aimar, solo cuando el Agente es Econocaribe puede asignar Clientes o Shippers Coloaders, favor de consultar con su supervisor y revisar el administrador de catalogos");
				            document.forma.ShipperData.focus();
                            return (false);
			            }
			        }
                }
            }
            <%end if %>
                        
            if (!valSelec(document.forma.AirportDepID)){return (false)};
			if (!valTxt(document.forma.RequestedRouting, 3, 5)){return (false)};
			if (!valTxt(document.forma.AirportToCode1, 3, 5)){return (false)};
			if (!valSelec(document.forma.CarrierID)){return (false)};
			if (!valSelec(document.forma.CurrencyID)){return (false)};
			if (!valSelec(document.forma.ChargeType)){return (false)};
			if (!valSelec(document.forma.ValChargeType)){return (false)};
			if (!valSelec(document.forma.OtherChargeType)){return (false)};
			if (!valTxt(document.forma.DeclaredValue, 2, 5)){return (false)};
			if (!valTxt(document.forma.AduanaValue, 2, 5)){return (false)};
            if (!valSelec(document.forma.AirportDesID)) { return (false) };
            if (!valSelec(document.forma.TipoCarga2)) { return (false) };

            if (document.forma.AirportDepID.value == document.forma.AirportDesID.value) {
                alert('Aeropuerto Origen debe ser distinto al Aeropuerto Destino');
                return false;
            }
			if (!valTxt(document.forma.FlightDate1, 3, 5)){return (false)};
			if (!valTxt(document.forma.FlightDate2, 3, 5)){return (false)};
			if (!valTxt(document.forma.SecuredValue, 3, 5)){return (false)};
			//if (document.forma.Countries.value!="HN"){
				if (!valTxt(document.forma.NoOfPieces, 1, 5)){return (false)};
				if (!valTxt(document.forma.Weights, 2, 5)){return (false)};
				if (!valTxt(document.forma.WeightsSymbol, 2, 5)){return (false)};
                if (!valTxt(document.forma.Commodities, 3, 5)){return (false)};
				if (!valTxt(document.forma.ChargeableWeights, 2, 5)){return (false)};
				if (!valTxt(document.forma.CarrierSubTot, 2, 5)){return (false)};
				if (!valTxt(document.forma.NatureQtyGoods, 2, 5)){return (false)};
				if (!valTxt(document.forma.TotWeight, 1, 5)){return (false)};
				if (!valTxt(document.forma.TotCarrierRate, 1, 5)){return (false)};
				if (!valSelec(document.forma.CAF)){return (false)};
				if (!valSelec(document.forma.TCAF)){return (false)};
			//};

            //Esta funcion revisa que la casilla de Simbolos solo puedan ingresar las palabras KG o LB
            if (document.forma.WeightsSymbol.value != ''){
                var Symbols = document.forma.WeightsSymbol.value.split("\r\n");
                document.forma.WeightsSymbol.value = "";
                var IncorrectSymbol = 1;
                var SymbolSep = "";
	                for (i=0;i<Symbols.length;i++) {
                        Symbols[i] = Symbols[i].toUpperCase();
                        if ((Symbols[i]== "KG") || (Symbols[i] == "KGS") || (Symbols[i] == "LB") || (Symbols[i] == "LBS")) {
			                IncorrectSymbol = 0; 
		                }
                        document.forma.WeightsSymbol.value = document.forma.WeightsSymbol.value + SymbolSep + Symbols[i];
                        SymbolSep = "\n"
	                }
                if (IncorrectSymbol==1) {
                    alert('En la casilla de simbolo de peso (kg/lb) solo puede ingresar las palabras "KG" o "LB"');
                    document.forma.WeightsSymbol.focus();
                    return (false);
                }
            }

            //XXXX Validando la asignacion de Monedas y Tipo de cobro (INT o LOC) de los Rubros
			if ((document.forma.TotCarrierRate.value!="") && (document.forma.TotCarrierRate.value>0)) {
				if (!valSelec(document.forma.CAF)){return (false)};
				if (!valSelec(document.forma.TCAF)){return (false)};
				if (!valSelec(document.forma.TPAF)){return (false)};
			};
			if ((document.forma.TerminalFee.value!="") && (document.forma.TerminalFee.value>0)) {
				if (!valSelec(document.forma.CTF)){return (false)};
				if (!valSelec(document.forma.TCTF)){return (false)};
				if (!valSelec(document.forma.TPTF)){return (false)};
			};
			if ((document.forma.CustomFee.value!="") && (document.forma.CustomFee.value>0)) {
				if (!valSelec(document.forma.CCF)){return (false)};
				if (!valSelec(document.forma.TCCF)){return (false)};
				if (!valSelec(document.forma.TPCF)){return (false)};
			};
			if ((document.forma.FuelSurcharge.value!="") && (document.forma.FuelSurcharge.value>0)) {
				if (!valSelec(document.forma.CFS)){return (false)};
				if (!valSelec(document.forma.TCFS)){return (false)};
				if (!valSelec(document.forma.TPFS)){return (false)};
			};
			if ((document.forma.SecurityFee.value!="") && (document.forma.SecurityFee.value>0)) {
				if (!valSelec(document.forma.CSF)){return (false)};
				if (!valSelec(document.forma.TCSF)){return (false)};
				if (!valSelec(document.forma.TPSF)){return (false)};
			};
			if ((document.forma.PBA.value!="") && (document.forma.PBA.value>0)) {
				if (!valSelec(document.forma.CPB)){return (false)};
				if (!valSelec(document.forma.TCPB)){return (false)};
				if (!valSelec(document.forma.TPPB)){return (false)};
			};
			if ((document.forma.TAX.value!="") && (document.forma.TAX.value>0)) {
				if (!valSelec(document.forma.CTX)){return (false)};
				if (!valSelec(document.forma.TCTX)){return (false)};
				if (!valSelec(document.forma.TPTX)){return (false)};
			};
			if ((document.forma.PickUp.value!="") && (document.forma.PickUp.value>0)) {
				if (!valSelec(document.forma.CPU)){return (false)};
				if (!valSelec(document.forma.TCPU)){return (false)};
				if (!valSelec(document.forma.TPPU)){return (false)};
			};
			if ((document.forma.Intermodal.value!="") && (document.forma.Intermodal.value>0)) {
				if (!valSelec(document.forma.CIM)){return (false)};
				if (!valSelec(document.forma.TCIM)){return (false)};
				if (!valSelec(document.forma.TPIM)){return (false)};
			};
			if ((document.forma.SedFilingFee.value!="") && (document.forma.SedFilingFee.value>0)) {
				if (!valSelec(document.forma.CFF)){return (false)};
				if (!valSelec(document.forma.TCFF)){return (false)};
				if (!valSelec(document.forma.TPFF)){return (false)};
			};

            //Rubros de Otros Cargos
			if ((document.forma.OtherChargeName1.value!="")||(document.forma.OtherChargeVal1.value!="")){
				if (!valTxt(document.forma.OtherChargeName1, 2, 5)){return (false)};
				if (!valSelec(document.forma.CO1)){return (false)};
				if (!valTxt(document.forma.OtherChargeVal1, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCO1)){return (false)};
				if (!valSelec(document.forma.TPO1)){return (false)};
			};
			if ((document.forma.OtherChargeName2.value!="")||(document.forma.OtherChargeVal2.value!="")){
				if (!valTxt(document.forma.OtherChargeName2, 2, 5)){return (false)};
				if (!valSelec(document.forma.CO2)){return (false)};
				if (!valTxt(document.forma.OtherChargeVal2, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCO2)){return (false)};
				if (!valSelec(document.forma.TPO2)){return (false)};
			};
			if ((document.forma.OtherChargeName3.value!="")||(document.forma.OtherChargeVal3.value!="")){
				if (!valTxt(document.forma.OtherChargeName3, 2, 5)){return (false)};
				if (!valSelec(document.forma.CO3)){return (false)};
				if (!valTxt(document.forma.OtherChargeVal3, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCO3)){return (false)};
				if (!valSelec(document.forma.TPO3)){return (false)};
			};
			if ((document.forma.OtherChargeName4.value!="")||(document.forma.OtherChargeVal4.value!="")){
				if (!valTxt(document.forma.OtherChargeName4, 2, 5)){return (false)};
				if (!valSelec(document.forma.CO4)){return (false)};
				if (!valTxt(document.forma.OtherChargeVal4, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCO4)){return (false)};
				if (!valSelec(document.forma.TPO4)){return (false)};
			};
			if ((document.forma.OtherChargeName5.value!="")||(document.forma.OtherChargeVal5.value!="")){
				if (!valTxt(document.forma.OtherChargeName5, 2, 5)){return (false)};
				if (!valSelec(document.forma.CO5)){return (false)};
				if (!valTxt(document.forma.OtherChargeVal5, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCO5)){return (false)};
				if (!valSelec(document.forma.TPO5)){return (false)};
			};
			if ((document.forma.OtherChargeName6.value!="")||(document.forma.OtherChargeVal6.value!="")){
				if (!valTxt(document.forma.OtherChargeName6, 2, 5)){return (false)};
				if (!valSelec(document.forma.CO6)){return (false)};
				if (!valTxt(document.forma.OtherChargeVal6, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCO6)){return (false)};
				if (!valSelec(document.forma.TPO6)){return (false)};
			};

			//Rubros del Carrier
			if ((document.forma.AdditionalChargeName3.value!="")||(document.forma.AdditionalChargeVal3.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName3, 2, 5)){return (false)};
				if (!valSelec(document.forma.CC1)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal3, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCC1)){return (false)};
				if (!valSelec(document.forma.TPC1)){return (false)};
			};
			if ((document.forma.AdditionalChargeName4.value!="")||(document.forma.AdditionalChargeVal4.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName4, 2, 5)){return (false)};
				if (!valSelec(document.forma.CC2)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal4, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCC2)){return (false)};
				if (!valSelec(document.forma.TPC2)){return (false)};
			};
			if ((document.forma.AdditionalChargeName5.value!="")||(document.forma.AdditionalChargeVal5.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName5, 2, 5)){return (false)};
				if (!valSelec(document.forma.CC3)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal5, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCC3)){return (false)};
				if (!valSelec(document.forma.TPC3)){return (false)};
			};
			if ((document.forma.AdditionalChargeName8.value!="")||(document.forma.AdditionalChargeVal8.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName8, 2, 5)){return (false)};
				if (!valSelec(document.forma.CC4)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal8, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCC4)){return (false)};
				if (!valSelec(document.forma.TPC4)){return (false)};
			};
			//Rubros del Agente
			if ((document.forma.AdditionalChargeName1.value!="")||(document.forma.AdditionalChargeVal1.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName1, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA1)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal1, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA1)){return (false)};
				if (!valSelec(document.forma.TPA1)){return (false)};
			};
			if ((document.forma.AdditionalChargeName2.value!="")||(document.forma.AdditionalChargeVal2.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName2, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA2)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal2, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA2)){return (false)};
				if (!valSelec(document.forma.TPA2)){return (false)};
			};
			if ((document.forma.AdditionalChargeName6.value!="")||(document.forma.AdditionalChargeVal6.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName6, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA3)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal6, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA3)){return (false)};
				if (!valSelec(document.forma.TPA3)){return (false)};
			};
			if ((document.forma.AdditionalChargeName7.value!="")||(document.forma.AdditionalChargeVal7.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName7, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA4)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal7, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA4)){return (false)};
				if (!valSelec(document.forma.TPA4)){return (false)};
			};
			if ((document.forma.AdditionalChargeName9.value!="")||(document.forma.AdditionalChargeVal9.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName9, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA5)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal9, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA5)){return (false)};
				if (!valSelec(document.forma.TPA5)){return (false)};
			};
			if ((document.forma.AdditionalChargeName10.value!="")||(document.forma.AdditionalChargeVal10.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName10, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA6)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal10, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA6)){return (false)};
				if (!valSelec(document.forma.TPA6)){return (false)};
			};
			if ((document.forma.AdditionalChargeName11.value!="")||(document.forma.AdditionalChargeVal11.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName11, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA7)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal11, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA7)){return (false)};
				if (!valSelec(document.forma.TPA7)){return (false)};
			};
			if ((document.forma.AdditionalChargeName12.value!="")||(document.forma.AdditionalChargeVal12.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName12, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA8)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal12, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA8)){return (false)};
				if (!valSelec(document.forma.TPA8)){return (false)};
			};
			if ((document.forma.AdditionalChargeName13.value!="")||(document.forma.AdditionalChargeVal13.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName13, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA9)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal13, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA9)){return (false)};
				if (!valSelec(document.forma.TPA9)){return (false)};
			};
			if ((document.forma.AdditionalChargeName14.value!="")||(document.forma.AdditionalChargeVal14.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName14, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA10)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal14, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA10)){return (false)};
				if (!valSelec(document.forma.TPA10)){return (false)};
			};
			if ((document.forma.AdditionalChargeName15.value!="")||(document.forma.AdditionalChargeVal15.value!="")){
				if (!valTxt(document.forma.AdditionalChargeName15, 2, 5)){return (false)};
				if (!valSelec(document.forma.CA11)){return (false)};
				if (!valTxt(document.forma.AdditionalChargeVal15, 1, 5)){return (false)};
				if (!valSelec(document.forma.TCA11)){return (false)};
				if (!valSelec(document.forma.TPA11)){return (false)};
			};			
			
			if (!valSelec(document.forma.SalespersonID)){return (false)};
			if (!valTxt(document.forma.AgentSignature, 3, 5)){return (false)};
			if (!valTxt(document.forma.AWBDate, 3, 5)){return (false)};
			if (!valTxt(document.forma.AgentContactSignature, 3, 5)){return (false)};
                    

            if (Action == 5)                 
                return true;
            
            if (!confirm( Action == 5 ? 'Confirme enviar datos a Exactus ? ' : 'Confirme Actualizar datos'  )) 
			    return false;

            document.forma.Action.value = Action;
            move();
            document.forma.submit();	

		} else {

            //2018-04-18
            //" & IIf(InStr(1,iIps,Request.ServerVariables("LOCAL_ADDR")),"localhost:3030","10.10.1.21:8181/admin") & "
            //document.forma.Action.value = Action;
            //document.forma.submit();

            if (document.forma.eliminar.value == 0) {
                alert("La Guia no se puede eliminar, porque tiene facturas relacionadas. Si desea realizar la operacion primero debe anular las facturas correspondientes");
				return(false);            
            } else {

                //2018-04
                <% if RoutingID > 0 then %>

                if (child) { //just in case its open
                    child.close();
                }          

                child = window.open(RoutingErrorSite + '?' + RoutingErrorUrl, 'iWinRou', 'location=yes,height=325,width=500,scrollbars=no,status=no,titlebar=no,resizable=no,menubar=no');
                
                var interval = setInterval(function() {
                    try {							
			            console.log(child.closed);						
			            if (child.closed) {
                            console.clear();
                            clearInterval(interval);                  
                            GetAjaxProcess('Utils.asp',RoutingErrorUrl);                            
			            }
                    } catch(e) {
                        // we're here when the child window has been navigated away or closed
                        if (child.closed) {
                            console.clear();
                            clearInterval(interval);
                            //alert("closed");
                            return; 
                        }
                    }
                }, 500);

                <% else %>
                    move();
                    document.forma.Action.value = Action;
                    document.forma.submit();	
                <% end if %>

            }
        }
		 
	 }



<%'Desplegando Otros Cargos del Transportista'
'response.write "var TAXES=" & TaxRate & ";" & vbCrLf
'For i = 0 To CountList2Values-1
'response.write "var _" & CarrierID & "_" & aList2Values(0,i) & "_TFPD='" & aList2Values(3,i) & "';" & vbCrLf
'response.write "var _" & CarrierID & "_" & aList2Values(0,i) & "_TFCS='" & aList2Values(4,i) & "';" & vbCrLf
'response.write "var _" & CarrierID & "_" & aList2Values(0,i) & "_CF='" & aList2Values(5,i) & "';" & vbCrLf
'response.write "var _" & CarrierID & "_" & aList2Values(0,i) & "_SECURITY='" & aList2Values(6,i) & "';" & vbCrLf
'response.write "var _" & CarrierID & "_" & aList2Values(0,i) & "_FUEL='" & aList2Values(7,i) & "';" & vbCrLf
'Next
%>

<%'Desplegando Tarifas del Transportista (Carrier)'
'For i = 0 To CountList5Values-1
'response.write "var _" & aList5Values(0,i) & "_" & aList5Values(1,i) & "_" & aList5Values(2,i) & "_" & aList5Values(3,i) & "='" & aList5Values(4,i) & "';" & vbCrLf
'Next
%>

var FlgTF=0;
var FlgCF=0;
var FlgFS=0;
var FlgSF=0;

//var CommoditiesCode = new Array();
//var CommoditiesName = new Array();
//var CommoditiesType = new Array();
var Currencies = new Array();
var CurrenciesCode = new Array();
var CurrenciesRate = new Array();
var CurrenciesSymbol = new Array();

<%if CarrierRates = "AS AGREED" then%>
var AsAgreed = true;
<%else%>
var AsAgreed = false;
<%end if%>

<%'Desplegando Monedas'
For i = 0 To CountList6Values-1
response.write "Currencies[" & i & "]=" & aList6Values(0,i) & "; CurrenciesCode[" & i & "]='" & aList6Values(1,i) & "'; CurrenciesRate[" & i & "]=" & aList6Values(2,i) & "; CurrenciesSymbol[" & i & "]='" & aList6Values(3,i) & "';" & vbCrLf
Next
%>


<%'Desplegando Datos de commodities'
'For i = 0 To CountList7Values-1
	'response.write "alert('"&aList7Values(1,i)&");"
	'aList7Values(1,i) = Replace(aList7Values(1,i),chr(13)&chr(10)," ")
'	response.write "CommoditiesCode[" & i & "]='" & aList7Values(3,i) & "';	CommoditiesName[" & i & "]='" & aList7Values(1,i) & "';	CommoditiesType[" & i & "]=" & aList7Values(2,i) & ";" & vbCrLf
'Next
%>

function SumVals(obj, destination) {
var Vals = obj.value.split("\r\n");
var TotVals = 0;
var Values = "";
	for (i=0;i<Vals.length;i++) { 
		if (Vals[i] != ""){
			TotVals = TotVals + (Vals[i]*1);
			Values = Values + Vals[i] + "\r\n"; 
		}
	}
	obj.value = Values;
	destination.value = Round(TotVals);
}

function CheckRates(theForm) {
 if (theForm.ChargeableWeights.value != "") {
 	CalcRate(theForm);
 }
}

function CalcRate(theForm) {


try { 


if (!AsAgreed) {
	if (!valSelec(document.forma.AirportDepID)){return (false)};
	if (!valSelec(document.forma.AirportDesID)){return (false)};

	var obj = theForm.ChargeableWeights;
	var obj2 = theForm.Weights;
	var Weights = obj.value.split("\r\n");

	//Excepcion GT para Linea Aerea American Airlines ID = 9
	//if ((document.forma.Countries.value=="GT") && (document.forma.CarrierID[document.forma.CarrierID.selectedIndex].value==9))
	//{
	//	Weights = obj2.value.split("\r\n");
	//}
	//Fin Excepcion

	if (document.forma.WType.selectedIndex==0) {
		var Weights2 = obj2.value.split("\r\n");
	} else {
		var Weights2 = Weights;
	}

	var CommodityTypes = theForm.CommoditiesTypes.value.split(",");
	var Rates = theForm.CarrierRates.value.split("\r\n");
	var Rate;
	var Weight = "";
	var CarrierRates = ""; 
	var CarrierSubTot = ""; 
	var TotCarrierRate = 0; 
	var TotWeightChargeable = 0;
	var TotWeightChargeableCS = 0;
	var TotWeightChargeablePD = 0;
	var RateName = "_" + theForm.CarrierID.value + "_" + theForm.AirportDepID.value + "_" + theForm.AirportDesID.value;
	
    var MinRate = 0;
    
	try {
		MinRate = eval(RateName + "_0");
	} catch(exception) {
        MinRate = 0;
        //2022-03-02
    	/*
        alert("No existen Tarifas para la ruta indicada");
		theForm.AirportDepID.value = "-1";
		theForm.AirportDesID.value = "-1";		
        theForm.CarrierRates.value = "";
        

		theForm.CarrierSubTot.value = 0;
		theForm.TotCarrierRate.value = 0;
		theForm.TotWeightChargeable.value = 0;

		theForm.TerminalFee.value = 0;
		theForm.CustomFee.value = 0;
		theForm.FuelSurcharge.value = 0;
		theForm.SecurityFee.value = 0;
		
		SumOtherCharges(theForm);
		AssignChargeWeight(theForm);
		CalcTax(theForm);
		CalcTot(theForm);

		return false;
        */
	}
    

    var msg;
	for (i=0;i<Weights.length;i++) { 
		if (Weights[i] != "") {
        			
            Weights[i] = parseFloat(Weights[i]);

            Rate = SetRate(Weights[i], RateName);

            //console.log("Rate=(" + Rate + ")");


            try { 
            
                if (Rate != '0') 
                    if (Rate.indexOf(",") > -1) 
                        Rate = Rate.replace(',','.');             

            } catch(exception) { 
                    alert('Rate:' + Rate);
            } 

            
            Rate = parseFloat(Rate);

            Rates[i] = parseFloat(Rates[i]);
            MinRate = parseFloat(MinRate);
       
            //console.log("*********" + i + "***************************************************");			
			//Si se ingresa tarifa manualmente, se utiliza esa en lugar de la tarifa definida en el sistema

			msg = "i=" + i + " - Weights.length=" + Weights.length + " - Rates.length=" + Rates.length;
			msg = msg + " - Rate=" + Rate + " - MinRate=" + MinRate + " - Rates[i]=" + Rates[i];
            //console.log(msg);

			if (i <= Rates.length) {
				if ((Rates[i]!=Rate) && (Rates[i]!=MinRate) && (Rates[i]!="")) {
                    //console.log("Asigno valor, Rate=" + Rate + " Rates[i]=" + Rates[i] );
                    Rate = parseFloat(Rates[i]);
				}
			}

			TempRate = Round(parseFloat(Weights[i]) * parseFloat(Rate));
                                    
            //console.log("TempRate (" + TempRate + ") = Weights[i] (" + Weights[i] + ") * Rate (" + Rate + ")");

			if (parseFloat(MinRate) > parseFloat(TempRate)) {
                //console.log("MinRate (" + MinRate + ") > TempRate (" + TempRate + ") ? Si");                
				CarrierRates = CarrierRates + Round(MinRate) + "\r\n"; 
                //console.log("CarrierRates = CarrierRates + Round(MinRate) (" + CarrierRates + ")");
				CarrierSubTot = CarrierSubTot + Round(MinRate) + "\r\n";
				TotCarrierRate = Round(TotCarrierRate*1 + MinRate*1);
			} else {				
                //console.log("MinRate (" + MinRate + ") > TempRate (" + TempRate + ") ? No");                
                CarrierRates = CarrierRates + Round(Rate) + "\r\n";
                //console.log("CarrierRates = CarrierRates + Round(Rate) (" + CarrierRates + ")");
				CarrierSubTot = CarrierSubTot + Round(TempRate) + "\r\n";
				TotCarrierRate = Round(TotCarrierRate*1 + TempRate*1);
			}
			Weight = Weight + Weights[i] + "\r\n";
            //console.log("CarrierRates = " + CarrierRates + " - TotCarrierRate = " + TotCarrierRate);
		}
	}
	
	for (i=0;i<Weights2.length;i++) { 
		if (Weights2[i] != "") {
			if (CommodityTypes[i]==1) {
				TotWeightChargeablePD = TotWeightChargeablePD + (Weights2[i]*1);
			}
			if (CommodityTypes[i]==2) {
				TotWeightChargeableCS = TotWeightChargeableCS + (Weights2[i]*1);
			} 
			TotWeightChargeable = TotWeightChargeable + (Weights2[i]*1);
		}
	}	


	//Excepcion GT para Linea Aerea American Airlines ID = 9
	//if (document.forma.CarrierID[document.forma.CarrierID.selectedIndex].value!=9)
	//{
	//	obj.value = Weight;
	//}
	//Fin Excepcion
	
	theForm.CarrierRates.value = CarrierRates; 
	theForm.CarrierSubTot.value = CarrierSubTot;
	theForm.TotCarrierRate.value = Round(TotCarrierRate);
	theForm.TotWeightChargeable.value = TotWeightChargeable;

	//alert(i + " - " + Weights.length + " - " + Rates.length + " - " + Rate + " - " + MinRate + " - " + Rates[i]);
	//alert(TotWeightChargeableCS + " - " + theForm.CarrierID.value + "_" + theForm.AirportDepID.value + "_TFCS" + " - " + TotWeightChargeablePD + " - " + theForm.CarrierID.value + "_" + theForm.AirportDepID.value + "_TFPD");
	if (FlgTF == 0) {
		theForm.TerminalFee.value = Round((TotWeightChargeableCS * eval("_" + theForm.CarrierID.value + "_" + theForm.AirportDepID.value + "_TFCS"))+(TotWeightChargeablePD * eval("_" + theForm.CarrierID.value + "_" + theForm.AirportDepID.value + "_TFPD")));
	} else {
		theForm.TerminalFee.value = Round(theForm.TerminalFee.value)
	}
	
    //2016-03-30
    //alert(theForm.CustomFee.value);

	if (FlgCF == 0) {
		theForm.CustomFee.value = Round(eval("_" + theForm.CarrierID.value + "_" + theForm.AirportDepID.value + "_CF"));
	} else {
		theForm.CustomFee.value = Round(theForm.CustomFee.value)
	}

    //alert(theForm.CustomFee.value);

	
	if (FlgFS == 0) {
		theForm.FuelSurcharge.value = Round(TotWeightChargeable * eval("_" + theForm.CarrierID.value + "_" + theForm.AirportDepID.value + "_FUEL"));
	} else {
		theForm.FuelSurcharge.value = Round(theForm.FuelSurcharge.value)
	}
	
	if (FlgSF == 0) {
		theForm.SecurityFee.value = Round(TotWeightChargeable * eval("_" + theForm.CarrierID.value + "_" + theForm.AirportDepID.value + "_SECURITY"));
	} else {
		theForm.SecurityFee.value = Round(theForm.SecurityFee.value)
	}
	
	SumOtherCharges(theForm);
	AssignChargeWeight(theForm);
	CalcTax(theForm);
	CalcTot(theForm);
}


    } catch(exception) { 
            //console.log('CalcRate Error:'); 2022-05-07
            //console.log(exception);
    } 


}

function Round(value){
	var number = (Math.round(value * 100)) / 100;
    return (number == Math.floor(number)) ? number + '.00' : ((number * 10 == Math.floor(number * 10)) ? number + '0' : number);
}

function SetRate(value, RateName){
var Rate = 0;
var DisplayAlert = 0;

<%'Desplegando Tarifas del Transportista (Carrier)'
if CountList4Values >= 0 then
For i = 0 To CountList4Values-1
	select case i
	case 0
		response.write "if (value < " & aList4Values(1,i) & ") {" & vbCrLf & _
					   "	try { Rate = eval(RateName + '_" & aList4Values(0,i) & "') } catch(exception) { DisplayAlert = 1 } " & vbCrLf & _
					   "} "
	case CountList4Values-1
		response.write "else if (value >= " & aList4Values(1,i-1) & " && value < "  & aList4Values(1,i) & ") {" & vbCrLf & _
					   "	try { Rate = eval(RateName + '_" & aList4Values(0,i-1) & "') } catch(exception) { DisplayAlert = 1 } " & vbCrLf & _
					   "} "
		response.write "else if (value >= " & aList4Values(1,i) & ") {" & vbCrLf & _
					   "	try { Rate = eval(RateName + '_" & aList4Values(0,i) & "') } catch(exception) { DisplayAlert = 1 } " & vbCrLf & _

					   "} "
	case else
		if aList4Values(1,i-1) <> aList4Values(1,i) then
		response.write "else if (value >= " & aList4Values(1,i-1) & " && value < "  & aList4Values(1,i) & ") {" & vbCrLf & _
					   "	try { Rate = eval(RateName + '_" & aList4Values(0,i-1) & "') } catch(exception) { DisplayAlert = 1 } " & vbCrLf & _
					   "} "
		else
		response.write "else if (value == " & aList4Values(1,i) & ") {" & vbCrLf & _
					   "	try { Rate = eval(RateName + '_" & aList4Values(0,i) & "') } catch(exception) { DisplayAlert = 1 } " & vbCrLf & _
					   "} "
		end if
	end select
Next
%>
	if (DisplayAlert == 1) {
    	console.log("No existen Tarifas para la ruta indicada.");
		theForm.AirportDepID.value = "-1";
		theForm.AirportDesID.value = "-1";
		theForm.CarrierRates.value = "";

		theForm.CarrierSubTot.value = 0;
		theForm.TotCarrierRate.value = 0;
		theForm.TotWeightChargeable.value = 0;

		theForm.TerminalFee.value = 0;
		theForm.CustomFee.value = 0;
		theForm.FuelSurcharge.value = 0;
		theForm.SecurityFee.value = 0;
		
		SumOtherCharges(theForm);
		AssignChargeWeight(theForm);
		CalcTax(theForm);
		CalcTot(theForm);
	}
<%end if%>	
	return (Rate);
}

function AssignChargeWeight(theForm){
	if (theForm.ValChargeType.value == "1") {
		theForm.TotChargeWeightCollect.value = "";	
		theForm.TotChargeWeightPrepaid.value = theForm.TotCarrierRate.value;
	} else {		
		theForm.TotChargeWeightCollect.value = theForm.TotCarrierRate.value;	
		theForm.TotChargeWeightPrepaid.value = "";
	}
}

function CalcTax(theForm){
 if (!AsAgreed) {
	var Result = '0.00';
	if (theForm.ChargeType.value == "1") {
		if (document.forma.Countries.value!="HN"){
			Result = Round((theForm.TotChargeWeightPrepaid.value*1 + theForm.AnotherChargesAgentPrepaid.value*1 + theForm.AnotherChargesCarrierPrepaid.value*1)*TAXES);
			theForm.TotChargeTaxCollect.value = "";	
			theForm.TotChargeTaxPrepaid.value = Result;
		} else {
			if (theForm.SetTax.checked) {
			Result = Round((theForm.TotChargeWeightPrepaid.value*1)*TAXES);
			theForm.TotChargeTaxCollect.value = "";	
			theForm.TotChargeTaxPrepaid.value = Result;
			} else {
			theForm.TotChargeTaxCollect.value = "";	
			theForm.TotChargeTaxPrepaid.value = '0.00';
			}
		}
	} else {		
		//Result = Round((theForm.TotChargeWeightCollect.value*1 + theForm.PBA.value*1 + theForm.AnotherChargesCarrierCollect.value*1)*TAXES);
		//theForm.TotChargeTaxCollect.value = Result;	
		theForm.TotChargeTaxCollect.value = "";	
		theForm.TotChargeTaxPrepaid.value = "";
	}
	theForm.TAX.value = Result;
 }
};

function SumOtherCharges(theForm){
 if (!AsAgreed) {
 	var Result = Round(theForm.TerminalFee.value*1 + theForm.CustomFee.value*1 + theForm.FuelSurcharge.value*1 + theForm.SecurityFee.value*1);

    if (document.forma.CCBLO1.value==1) {
        Result = Round(Result*1 + theForm.OtherChargeVal1.value*1);
    };
    if (document.forma.CCBLO2.value==1) {
        Result = Round(Result*1 + theForm.OtherChargeVal2.value*1);
    };
    if (document.forma.CCBLO3.value==1) {
        Result = Round(Result*1 + theForm.OtherChargeVal3.value*1);
    };
    if (document.forma.CCBLO4.value==1) {
        Result = Round(Result*1 + theForm.OtherChargeVal4.value*1);
    };
    if (document.forma.CCBLO5.value==1) {
        Result = Round(Result*1 + theForm.OtherChargeVal5.value*1);
    };
    if (document.forma.CCBLO6.value==1) {
        Result = Round(Result*1 + theForm.OtherChargeVal6.value*1);
    };

    if (document.forma.CCBLC1.value==1) {
            Result = Round(Result*1 + theForm.AdditionalChargeVal3.value*1);
    };
    if (document.forma.CCBLC2.value==1) {
            Result = Round(Result*1 + theForm.AdditionalChargeVal4.value*1);
    };
    if (document.forma.CCBLC3.value==1) {
            Result = Round(Result*1 + theForm.AdditionalChargeVal5.value*1);
    };
    if (document.forma.CCBLC4.value==1) {
            Result = Round(Result*1 + theForm.AdditionalChargeVal8.value*1);
    };    
	
    //if ((theForm.ChargeType.value== "1") && (theForm.Countries.value=="GT")) { //1.Prepaid, 2.Collect, si es Prepaid-Export en Guatemala no se suman los valores
	//	var Result2 = Round(theForm.PBA.value*1 + theForm.PickUp.value*1 + theForm.Intermodal.value*1 + theForm.SedFilingFee.value*1);
	//} else {
	if (theForm.ChargeType.value == "2") { //1.Prepaid, 2.Collect
		var Result2 = Round(theForm.PBA.value*1 + theForm.PickUp.value*1 + theForm.Intermodal.value*1 + theForm.SedFilingFee.value*1);
        if (document.forma.CCBLA1.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal1.value*1);
        };
        if (document.forma.CCBLA2.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal2.value*1);
        };
        if (document.forma.CCBLA3.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal6.value*1);
        };
        if (document.forma.CCBLA4.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal7.value*1);
        };  
        if (document.forma.CCBLA5.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal9.value*1);
        };   
        if (document.forma.CCBLA6.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal10.value*1);
        };  
        if (document.forma.CCBLA7.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal11.value*1);
        };   
        if (document.forma.CCBLA8.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal12.value*1);
        };   
        if (document.forma.CCBLA9.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal13.value*1);
        };  
        if (document.forma.CCBLA10.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal14.value*1);
        }; 
        if (document.forma.CCBLA11.value==1) {
            Result2 = Round(Result2*1 + theForm.AdditionalChargeVal15.value*1);
        };

	} else {
		var Result2 = Round(theForm.PickUp.value*1 + theForm.Intermodal.value*1 + theForm.SedFilingFee.value*1);
	}	
	
	if (theForm.OtherChargeType.value == "1") { //1.Prepaid, 2.Collect
		theForm.AnotherChargesCarrierCollect.value = "";
		theForm.AnotherChargesCarrierPrepaid.value = Result;
		theForm.AnotherChargesAgentCollect.value = "";	
		theForm.AnotherChargesAgentPrepaid.value = Result2;
	} else {		
		theForm.AnotherChargesCarrierCollect.value = Result;	
		theForm.AnotherChargesCarrierPrepaid.value = "";
		theForm.AnotherChargesAgentCollect.value = Result2;
		theForm.AnotherChargesAgentPrepaid.value = "";
	}
 }	
};

function ValPBA(theForm){
 if (!AsAgreed) {
	if (theForm.PBA.value < <%=Session("PBAValue")%>){
		alert("El PBA no puede ser menor de <%=Session("PBAValue")%>");
		theForm.PBA.focus();		
	} else {
		SumOtherCharges(theForm);
		CalcTax(theForm);
		CalcTot(theForm);
	}
 }	
}

function CalcTot(theForm){
 if (!AsAgreed) {
	var Result = 0;
	if (theForm.ChargeType.value == "1") {
		Result = Round(theForm.TotChargeWeightPrepaid.value*1 + theForm.TotChargeTaxPrepaid.value*1 + theForm.AnotherChargesCarrierPrepaid.value*1 + theForm.AnotherChargesAgentPrepaid.value*1);
		theForm.TotCollect.value = "";	
		theForm.TotPrepaid.value = Result;
	} else {		
		Result = Round(theForm.TotChargeWeightCollect.value*1 + theForm.TotChargeTaxCollect.value*1 + theForm.AnotherChargesCarrierCollect.value*1 + theForm.AnotherChargesAgentCollect.value*1);
		theForm.TotCollect.value = Result;	
		theForm.TotPrepaid.value = "";
	}
 } else {
	if (theForm.ChargeType.value == "1") {
		theForm.TotCollect.value = "";	
		theForm.TotPrepaid.value = "AS AGREED";
		theForm.AnotherChargesAgentPrepaid.value = "AS AGREED";
		theForm.AnotherChargesAgentCollect.value = "";
		theForm.AnotherChargesCarrierPrepaid.value = "AS AGREED";
		theForm.AnotherChargesCarrierCollect.value = "";		
	} else {		
		theForm.TotCollect.value = "AS AGREED";	
		theForm.TotPrepaid.value = "";
		theForm.AnotherChargesAgentPrepaid.value = "";
		theForm.AnotherChargesAgentCollect.value = "AS AGREED";
		theForm.AnotherChargesCarrierPrepaid.value = "";
		theForm.AnotherChargesCarrierCollect.value = "AS AGREED";		
	} 
 }
};

function As_Agreed(obj){
	if (obj.checked) {
		AsAgreed = true;
		document.forma.CarrierRates.value = "AS AGREED";
		document.forma.CarrierSubTot.value = "AS AGREED";
		document.forma.TotCarrierRate.value = "AS AGREED";
		/*
		document.forma.TotChargeWeightPrepaid.value = "";
		document.forma.TotChargeWeightCollect.value = "";
		document.forma.TotChargeValuePrepaid.value = "";
		document.forma.TotChargeValueCollect.value = "";
		document.forma.TotChargeTaxPrepaid.value = "";
		document.forma.TotChargeTaxCollect.value = "";
		document.forma.AnotherChargesAgentPrepaid.value = "";
		document.forma.AnotherChargesAgentCollect.value = "";
		document.forma.AnotherChargesCarrierPrepaid.value = "";
		document.forma.AnotherChargesCarrierCollect.value = "";
		document.forma.TerminalFee.value = "";
		document.forma.CustomFee.value = "";
		document.forma.FuelSurcharge.value = "";
		document.forma.SecurityFee.value = "";
		document.forma.PBA.value = "";
		document.forma.TAX.value = "";
		document.forma.PickUp.value = "";
		document.forma.Intermodal.value = "";
		document.forma.SedFilingFee.value = "";
		document.forma.AdditionalChargeVal1.value = "";
		document.forma.AdditionalChargeVal2.value = "";		
		document.forma.AdditionalChargeVal3.value = "";
		document.forma.AdditionalChargeVal4.value = "";		
        */
		AssignChargeWeight(document.forms[0]);
		CalcTot(document.forms[0]);
	} else {
		AsAgreed = false;
		document.forma.CarrierRates.value = "";
		document.forma.CarrierSubTot.value = "";
		document.forma.TotCarrierRate.value = "";
		document.forma.TotChargeWeightPrepaid.value = "";
		document.forma.TotChargeWeightCollect.value = "";
		document.forma.PBA.value = Round(<%=PBA%>);
		AssignChargeWeight(document.forms[0]);
		CalcTot(document.forms[0]);		
	}
}

function SetXchange(obj){
	for (j=0;j<=Currencies.length;j++) { 
		if (Currencies[j]==obj.value){
			document.forma.AccountInformation.value = 'ROE 1 ' + CurrenciesCode[j] + ' X ' + CurrenciesSymbol[j] + CurrenciesRate[j];
		}
	}	
}

var numb = "0123456789./\r/\n";
function res(t,v){
	var w = "";
	for (i=0; i < t.value.length; i++) {
	x = t.value.charAt(i);
	if (v.indexOf(x,0) != -1)
		w += x;
	}
	t.value = w;
}

function DelAgentCharge(Pos1,Pos2) {
    if (confirm('¿ Confirme Borrar Este Cargo ?')) {
	    document.forma.elements["A"+Pos1].value=0;
	    document.forma.elements["A"+Pos1+"_Tarifa"].value='';
	    document.forma.elements["A"+Pos1+"_TarifaTipo"].value='';
	    document.forma.elements["A"+Pos1+"_Regimen"].value='';
	    document.forma.elements["VA"+Pos1].value=0;
	    document.forma.elements["CA"+Pos1].value=-1;
	    document.forma.elements["AdditionalChargeName"+Pos2].value='';
	    document.forma.elements["AdditionalChargeVal"+Pos2].value='';
	    document.forma.elements["TCA"+Pos1].value='-1';
	    document.forma.elements["TPA"+Pos1].value='-1';
	    document.forma.elements["SVNA"+Pos1].value='';
	    document.forma.elements["SVIA"+Pos1].value=0;
	    document.forma.elements["INVA"+Pos1].value='0';
        document.forma.elements["CCBLA"+Pos1].value='1';
    }
	return false; 
 }

function DelCarrierCharge(Pos1,Pos2) {
    if (confirm('¿ Confirme Borrar Este Cargo ?')) {
	    document.forma.elements["C"+Pos1].value=0;
	    document.forma.elements["C"+Pos1+"_Tarifa"].value='';
	    document.forma.elements["C"+Pos1+"_TarifaTipo"].value='';
	    document.forma.elements["C"+Pos1+"_Regimen"].value='';
	    document.forma.elements["VC"+Pos1].value=0;
	    document.forma.elements["CC"+Pos1].value=-1;
	    document.forma.elements["AdditionalChargeName"+Pos2].value='';
	    document.forma.elements["AdditionalChargeVal"+Pos2].value='';
	    document.forma.elements["TCC"+Pos1].value='-1';
	    document.forma.elements["TPC"+Pos1].value='-1';
	    document.forma.elements["SVNC"+Pos1].value='';
	    document.forma.elements["SVIC"+Pos1].value=0;
	    document.forma.elements["INVC"+Pos1].value='0';
        document.forma.elements["CCBLC"+Pos1].value='1';
    }
	return false; 
 }


 function DelOtherCharge(Pos1,Pos2) {
    if (confirm('¿ Confirme Borrar Este Cargo ?')) {
	    document.forma.elements["O"+Pos1].value=0;
	    document.forma.elements["O"+Pos1+"_Tarifa"].value='';
	    document.forma.elements["O"+Pos1+"_TarifaTipo"].value='';
	    document.forma.elements["O"+Pos1+"_Regimen"].value='';
	    document.forma.elements["VO"+Pos1].value=0;
	    document.forma.elements["CO"+Pos1].value=-1;
	    document.forma.elements["OtherChargeName"+Pos2].value='';
	    document.forma.elements["OtherChargeVal"+Pos2].value='';
	    document.forma.elements["TCO"+Pos1].value='-1';
	    document.forma.elements["TPO"+Pos1].value='-1';
	    document.forma.elements["SVNO"+Pos1].value='';
	    document.forma.elements["SVIO"+Pos1].value=0;
	    document.forma.elements["INVO"+Pos1].value=0;
        document.forma.elements["CCBLO"+Pos1].value='1';
    }
	return false; 
}


var AgentsPos = new Array();
var CarriersPos = new Array();
			
	AgentsPos[1] = 1;
	AgentsPos[2] = 2;
	AgentsPos[3] = 6;
	AgentsPos[4] = 7;
	AgentsPos[5] = 9;
	AgentsPos[6] = 10;
	AgentsPos[7] = 11;
	AgentsPos[8] = 12;
	AgentsPos[9] = 13;
	AgentsPos[10] = 14;
	AgentsPos[11] = 15; 

	CarriersPos[1] = 3;
	CarriersPos[2] = 4;
	CarriersPos[3] = 5;
	CarriersPos[4] = 8;

function CheckAgentDoble(Pos) {
	//Custom Fee		14	3	INVCF
	if ((document.forma.elements["SVIA"+Pos].value==3) && (document.forma.elements["A"+Pos].value==14) && (document.forma.elements["INVCF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelAgentCharge(Pos,AgentsPos[Pos]);
		return (false);
	}
	//Terminal Fee		15	3	INVTF
	if ((document.forma.elements["SVIA"+Pos].value==3) && (document.forma.elements["A"+Pos].value==15) && (document.forma.elements["INVTF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelAgentCharge(Pos,AgentsPos[Pos]);
		return (false);
	}
	//Air Freight		11	3	INVAF
	if ((document.forma.elements["SVIA"+Pos].value==3) && (document.forma.elements["A"+Pos].value==11) && (document.forma.elements["INVAF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelAgentCharge(Pos,AgentsPos[Pos]);
		return (false);
	}
	//Fuel Surcharge	12	3	INVFS
	if ((document.forma.elements["SVIA"+Pos].value==3) && (document.forma.elements["A"+Pos].value==12) && (document.forma.elements["INVFS"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelAgentCharge(Pos,AgentsPos[Pos]);
		return (false);
	}
	//Security Fee		13	3	INVSF
	if ((document.forma.elements["SVIA"+Pos].value==3) && (document.forma.elements["A"+Pos].value==13) && (document.forma.elements["INVSF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelAgentCharge(Pos,AgentsPos[Pos]);
		return (false);
	}
	//Pick Up			31	5	INVPU
	if ((document.forma.elements["SVIA"+Pos].value==5) && (document.forma.elements["A"+Pos].value==31) && (document.forma.elements["INVPU"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelAgentCharge(Pos,AgentsPos[Pos]);
		return (false);
	}
	//Sed Filing Fee	38	3	INVFF
	if ((document.forma.elements["SVIA"+Pos].value==3) && (document.forma.elements["A"+Pos].value==38) && (document.forma.elements["INVFF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelAgentCharge(Pos,AgentsPos[Pos]);
		return (false);
	}
	//Intermodal		115	5	INVIM
	if ((document.forma.elements["SVIA"+Pos].value==5) && (document.forma.elements["A"+Pos].value==115) && (document.forma.elements["INVIM"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelAgentCharge(Pos,AgentsPos[Pos]);
		return (false);
	}
	//PBA				116	3	INVPB
	if ((document.forma.elements["SVIA"+Pos].value==3) && (document.forma.elements["A"+Pos].value==116) && (document.forma.elements["INVPB"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelAgentCharge(Pos,AgentsPos[Pos]);
		return (false);
	}
	//Otros Cargos Agente
	for (i=1; i<=11;i++) {
		if  (i!= Pos) {
			if ((document.forma.elements["SVIA"+i].value==document.forma.elements["SVIA"+Pos].value) && 
			(document.forma.elements["A"+i].value==document.forma.elements["A"+Pos].value) &&
			(document.forma.elements["INVA"+i].value=='0')) {
				alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
				DelAgentCharge(Pos,AgentsPos[Pos]);
				return (false);
			}
		}		
	}
	//Otros Cargos Transportista
	for (i=1; i<=4;i++) {
		  if ((document.forma.elements["SVIC"+i].value==document.forma.elements["SVIA"+Pos].value) && 
		  (document.forma.elements["C"+i].value==document.forma.elements["A"+Pos].value) &&
		  (document.forma.elements["INVC"+i].value=='0')) {
			  alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
			  DelAgentCharge(Pos,AgentsPos[Pos]);
			  return (false);
		  }
	}

	//Otros Cargos
	for (i=1; i<=6;i++) {
		if ((document.forma.elements["SVIO"+i].value==document.forma.elements["SVIA"+Pos].value) && 
		(document.forma.elements["O"+i].value==document.forma.elements["A"+Pos].value) &&
		(document.forma.elements["INVO"+i].value=='0')) {
			alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
			DelAgentCharge(Pos,AgentsPos[Pos]);
			return (false);
		}
	}

}

function CheckCarrierDoble(Pos) {
	//Custom Fee		14	3	INVCF
	if ((document.forma.elements["SVIC"+Pos].value==3) && (document.forma.elements["C"+Pos].value==14) && (document.forma.elements["INVCF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelCarrierCharge(Pos,CarriersPos[Pos]);
		return (false);
	}
	//Terminal Fee		15	3	INVTF
	if ((document.forma.elements["SVIC"+Pos].value==3) && (document.forma.elements["C"+Pos].value==15) && (document.forma.elements["INVTF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelCarrierCharge(Pos,CarriersPos[Pos]);
		return (false);
	}
	//Air Freight		11	3	INVAF
	if ((document.forma.elements["SVIC"+Pos].value==3) && (document.forma.elements["C"+Pos].value==11) && (document.forma.elements["INVAF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelCarrierCharge(Pos,CarriersPos[Pos]);
		return (false);
	}
	//Fuel Surcharge	12	3	INVFS
	if ((document.forma.elements["SVIC"+Pos].value==3) && (document.forma.elements["C"+Pos].value==12) && (document.forma.elements["INVFS"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelCarrierCharge(Pos,CarriersPos[Pos]);
		return (false);
	}
	//Security Fee		13	3	INVSF
	if ((document.forma.elements["SVIC"+Pos].value==3) && (document.forma.elements["C"+Pos].value==13) && (document.forma.elements["INVSF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelCarrierCharge(Pos,CarriersPos[Pos]);
		return (false);
	}
	//Pick Up			31	5	INVPU
	if ((document.forma.elements["SVIC"+Pos].value==5) && (document.forma.elements["C"+Pos].value==31) && (document.forma.elements["INVPU"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelCarrierCharge(Pos,CarriersPos[Pos]);
		return (false);
	}
	//Sed Filing Fee	38	3	INVFF
	if ((document.forma.elements["SVIC"+Pos].value==3) && (document.forma.elements["C"+Pos].value==38) && (document.forma.elements["INVFF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelCarrierCharge(Pos,CarriersPos[Pos]);
		return (false);
	}
	//Intermodal		115	5	INVIM
	if ((document.forma.elements["SVIC"+Pos].value==5) && (document.forma.elements["C"+Pos].value==115) && (document.forma.elements["INVIM"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelCarrierCharge(Pos,CarriersPos[Pos]);
		return (false);
	}
	//PBA				116	3	INVPB
	if ((document.forma.elements["SVIC"+Pos].value==3) && (document.forma.elements["C"+Pos].value==116) && (document.forma.elements["INVPB"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelCarrierCharge(Pos,CarriersPos[Pos]);
		return (false);
	}
	//Otros Cargos Agente
	for (i=1; i<=11;i++) {
		  if ((document.forma.elements["SVIA"+i].value==document.forma.elements["SVIC"+Pos].value) && 
		  (document.forma.elements["A"+i].value==document.forma.elements["C"+Pos].value) &&
		  (document.forma.elements["INVA"+i].value=='0')) {
			  alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
			  DelCarrierCharge(Pos,CarriersPos[Pos]);
			  return (false);
		  }
	}
	//Otros Cargos Transportista
	for (i=1; i<=4;i++) {
		if  (i!= Pos) {
			if ((document.forma.elements["SVIC"+i].value==document.forma.elements["SVIC"+Pos].value) && 
			(document.forma.elements["C"+i].value==document.forma.elements["C"+Pos].value) &&
			(document.forma.elements["INVC"+i].value=='0')) {
				alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
				DelCarrierCharge(Pos,CarriersPos[Pos]);
				return (false);
			}
		}
	}

	//Otros Cargos
	for (i=1; i<=6;i++) {
		if ((document.forma.elements["SVIO"+i].value==document.forma.elements["SVIC"+Pos].value) && 
		(document.forma.elements["O"+i].value==document.forma.elements["C"+Pos].value) &&
		(document.forma.elements["INVO"+i].value=='0')) {
			alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
			DelCarrierCharge(Pos,CarriersPos[Pos]);
			return (false);
		}
	}

}

function CheckOtherDoble(Pos) {
	//Air Freight		11	3	INVAF
	if ((document.forma.elements["SVIO"+Pos].value==3) && (document.forma.elements["O"+Pos].value==11) && (document.forma.elements["INVAF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelOtherCharge(Pos,Pos);
		return (false);
	}
	//Fuel Surcharge	12	3	INVFS
	if ((document.forma.elements["SVIO"+Pos].value==3) && (document.forma.elements["O"+Pos].value==12) && (document.forma.elements["INVFS"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelOtherCharge(Pos,Pos);
		return (false);
	}
	//Security Fee		13	3	INVSF
	if ((document.forma.elements["SVIO"+Pos].value==3) && (document.forma.elements["O"+Pos].value==13) && (document.forma.elements["INVSF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelOtherCharge(Pos,Pos);
		return (false);
	}
	//Pick Up			31	5	INVPU
	if ((document.forma.elements["SVIO"+Pos].value==5) && (document.forma.elements["O"+Pos].value==31) && (document.forma.elements["INVPU"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelOtherCharge(Pos,Pos);
		return (false);
	}
	//Sed Filing Fee	38	3	INVFF
	if ((document.forma.elements["SVIO"+Pos].value==3) && (document.forma.elements["O"+Pos].value==38) && (document.forma.elements["INVFF"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelOtherCharge(Pos,Pos);
		return (false);
	}
	//Intermodal		115	5	INVIM
	if ((document.forma.elements["SVIO"+Pos].value==5) && (document.forma.elements["O"+Pos].value==115) && (document.forma.elements["INVIM"].value=='0')) {
		alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
		DelOtherCharge(Pos,Pos);
		return (false);
	}
	//Otros Cargos Agente
	for (i=1; i<=11;i++) {
		  if ((document.forma.elements["SVIA"+i].value==document.forma.elements["SVIO"+Pos].value) && 
		  (document.forma.elements["A"+i].value==document.forma.elements["O"+Pos].value) &&
		  (document.forma.elements["INVA"+i].value=='0')) {
			  alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
			  DelOtherCharge(Pos,Pos);
			  return (false);
		  }
	}
	//Otros Cargos Transportista
	for (i=1; i<=4;i++) {
		  if ((document.forma.elements["SVIC"+i].value==document.forma.elements["SVIO"+Pos].value) && 
		  (document.forma.elements["C"+i].value==document.forma.elements["O"+Pos].value) &&
		  (document.forma.elements["INVC"+i].value=='0')) {
			  alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
			  DelOtherCharge(Pos,Pos);
			  return (false);
		  }
	}
	//Otros Cargos
	for (i=1; i<=6;i++) {
		if  (i!= Pos) {
			if ((document.forma.elements["SVIO"+i].value==document.forma.elements["SVIO"+Pos].value) && 
			(document.forma.elements["O"+i].value==document.forma.elements["O"+Pos].value) &&
			(document.forma.elements["INVO"+i].value=='0')) {
				alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
				DelOtherCharge(Pos,Pos);
				return (false);
			}
		}
	}
}

function CheckDoble(RubName, RubName2) {
	var ServiceID=0;
	var RubID=0;
	
	//Custom Fee		14	3	INVCF
	if (RubName=="CustomFee") {
		ServiceID=3;
		RubID=14;
	}
	//Terminal Fee		15	3	INVTF
	if (RubName=="TerminalFee") {
		ServiceID=3;
		RubID=15;
	}
	//PBA				116	3	INVPB
	if (RubName=="PBA") {
		ServiceID=3;
		RubID=116;
	}
	//Air Freight		11	3	INVAF
	if (RubName=="AirFreight") {
		ServiceID=3;
		RubID=11;
	}
	//Fuel Surcharge	12	3	INVFS
	if (RubName=="FuelSurcharge") {
		ServiceID=3;
		RubID=12;
	}
	//Security Fee		13	3	INVSF
	if (RubName=="SecurityFee") {
		ServiceID=3;
		RubID=13;
	}
	//Pick Up			31	5	INVPU
	if (RubName=="PickUp") {
		ServiceID=5;
		RubID=31;
	}
	//Sed Filing Fee	38	3	INVFF
	if (RubName=="SedFilingFee") {
		ServiceID=3;
		RubID=38;
	}	
	//Intermodal		115	5	INVIM
	if (RubName=="Intermodal") {
		ServiceID=5;
		RubID=115;
	}
	
	//Otros Cargos Agente
	for (i=1; i<=11;i++) {
		  if ((document.forma.elements["SVIA"+i].value==ServiceID) && 
		  (document.forma.elements["A"+i].value==RubID) &&
		  (document.forma.elements["INVA"+i].value=='0')) {
			  alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
			  document.forma.elements["C"+RubName2].value=-1;
			  document.forma.elements[RubName].value='';
			  document.forma.elements["TC"+RubName2].value=-1;
			  document.forma.elements["TP"+RubName2].value=-1
			  return (false);
		  }
	}
	//Otros Cargos Transportista
	for (i=1; i<=4;i++) {
		  if ((document.forma.elements["SVIC"+i].value==ServiceID) && 
		  (document.forma.elements["C"+i].value==RubID) &&
		  (document.forma.elements["INVC"+i].value=='0')) {
			  alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
			  document.forma.elements["C"+RubName2].value=-1;
			  document.forma.elements[RubName].value='';
			  document.forma.elements["TC"+RubName2].value=-1;
			  document.forma.elements["TP"+RubName2].value=-1
			  return (false);
		  }
	}

	//Otros Cargos
	for (i=1; i<=6;i++) {
		if  (i!= Pos) {
			if ((document.forma.elements["SVIO"+i].value==document.forma.elements["SVIO"+Pos].value) && 
			(document.forma.elements["O"+i].value==document.forma.elements["O"+Pos].value) &&
			(document.forma.elements["INVO"+i].value=='0')) {
				alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
				DelOtherCharge(Pos,Pos);
				return (false);
			}
		}
	}
}

function AddCharge2(ChargeName, ChargePos, ChargeMoneda, servicio, rubro) {

    return false;

}

function AddCharge(ChargeName, ChargePos, ChargeMoneda) {
    var servicio = document.forms[0].elements['SVI' + ChargePos].value;
    var rubro = document.forms[0].elements[ChargePos].value;
    AddCharge2(ChargeName, ChargePos, ChargeMoneda, servicio, rubro); 
}

function Open(ObjectID, TipoCarga, AWBType){
	
    if (child) { //just in case its open
        child.close();
	}

    if (!valSelec(document.forma.TipoCarga2)) { return (false) };

    child = window.open('AwbCharges.asp?OID=' + ObjectID + '&TC=' + TipoCarga + '&AT=' + AWBType, 'Cargos', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1240,height=450,top=100,left=50');

    var interval = setInterval(function () {
        try {
            console.log(child.closed);
            if (child.closed) {
                console.clear();
                clearInterval(interval);

                location.reload();
            }
        } catch (e) {
            // we're here when the child window has been navigated away or closed
            if (child.closed) {
                console.clear();
                clearInterval(interval);
                location.reload();
            }
        }
	}, 500);

	return false;
}

</script>
<body>

<% 
'response.write "(" & TerminalFee_Routing & ")(" & TerminalFee & ")(" & Request.Form("CTF") & ")(" & Request.Form("TCTF") & ")(" & Request.Form("TPTF") & ")<br>" 
'response.write "(" & FuelSurcharge_Routing & ")(" & FuelSurcharge & ")(" & Request.Form("CFS") & ")(" & Request.Form("TCFS") & ")(" & Request.Form("TPFS") & ")<br>" 
%>


<div id="myProgress">
  <div id="myBar">10%</div>
</div>


<%if BAWResult <> "" then %>
    <div class=label><font color=<%if InStr(BAWResult,"Exitosamente") then %>blue<%else %>red<%end if %>><%=Replace(BAWResult,"\n","<br>")%></font></div>
<%end if %>
<form name="forma" action="InsertData.asp" method="post">

	<INPUT name="Peso2Anterior"        type=hidden value="<%=Peso2Anterior%>">			   	
	<INPUT name="Action"        type=hidden value=0>			   	
	<INPUT name="RoutingID" 	type=hidden value="<%=RoutingID%>">
	<INPUT name="Seguro" 		type=hidden value="<%=Seguro%>">    
	<INPUT name="routing_seg" 	type=hidden value="<%=routing_seg%>"> 
    <INPUT name="routing_adu" 	type=hidden value="<%=routing_adu%>"> 
    <INPUT name="routing_ter" 	type=hidden value="<%=routing_ter%>"> 
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
	<INPUT name="AT" type=hidden value="<%=AwbType%>">
	<INPUT name="Closed" type=hidden value="<%=Closed%>">
	<INPUT name="ShipperID" type=hidden value="<%=ShipperID%>">
	<INPUT name="ConsignerID" type=hidden value="<%=ConsignerID%>">
	<INPUT name="ShipperColoader" type=hidden value="<%=ShipperColoader%>">
	<INPUT name="ConsignerColoader" type=hidden value="<%=ConsignerColoader%>">
    <INPUT name="AgentID" type=hidden value="<%=AgentID%>">
	<INPUT name="AgentNeutral" type=hidden value="<%=AgentNeutral%>">
	<INPUT name="ShipperAddrID" type=hidden value="<%=ShipperAddrID%>">
	<INPUT name="ConsignerAddrID" type=hidden value="<%=ConsignerAddrID%>">
	<INPUT name="AgentAddrID" type=hidden value="<%=AgentAddrID%>">	
    <INPUT name="id_coloader" type=hidden value="<%=id_coloader%>">
	<INPUT name="RAirportDepID" type=hidden value="">
	<INPUT name="RAirportDesID" type=hidden value="">

	<INPUT name="Main" type=hidden value="1">
	<INPUT name="Movimiento" type=hidden value="<%=IIf(AWBType = 1,"EXPORT","IMPORT")%>">
		
	<INPUT name="VAF" type=hidden value="0">
	<INPUT name="VTF" type=hidden value="0">
	<INPUT name="VCF" type=hidden value="0">
	<INPUT name="VFS" type=hidden value="0">
	<INPUT name="VSF" type=hidden value="0">
	<INPUT name="VPB" type=hidden value="0">
	<INPUT name="VPU" type=hidden value="0">
	<INPUT name="VIM" type=hidden value="0">
	<INPUT name="VFF" type=hidden value="0">
	
	<INPUT name="VO1" type=hidden value="0">
	<INPUT name="VO2" type=hidden value="0">
	<INPUT name="VO3" type=hidden value="0">
	<INPUT name="VO4" type=hidden value="0">
	<INPUT name="VO5" type=hidden value="0">
	<INPUT name="VO6" type=hidden value="0">

	<INPUT name="VA1" type=hidden value="0">
	<INPUT name="VA2" type=hidden value="0">
	<INPUT name="VA3" type=hidden value="0">
	<INPUT name="VA4" type=hidden value="0">
	<INPUT name="VA5" type=hidden value="0">
	<INPUT name="VA6" type=hidden value="0">
	<INPUT name="VA7" type=hidden value="0">
	<INPUT name="VA8" type=hidden value="0">
	<INPUT name="VA9" type=hidden value="0">
	<INPUT name="VA10" type=hidden value="0">
	<INPUT name="VA11" type=hidden value="0">
	<INPUT name="VC1" type=hidden value="0">
	<INPUT name="VC2" type=hidden value="0">
	<INPUT name="VC3" type=hidden value="0">
	<INPUT name="VC4" type=hidden value="0">

	<INPUT name="NO1" type=hidden value="">
	<INPUT name="NO2" type=hidden value="">
	<INPUT name="NO3" type=hidden value="">
	<INPUT name="NO4" type=hidden value="">
	<INPUT name="NO5" type=hidden value="">
	<INPUT name="NO6" type=hidden value="">

	<INPUT name="NA1" type=hidden value="">
	<INPUT name="NA2" type=hidden value="">
	<INPUT name="NA3" type=hidden value="">
	<INPUT name="NA4" type=hidden value="">
	<INPUT name="NA5" type=hidden value="">
	<INPUT name="NA6" type=hidden value="">
	<INPUT name="NA7" type=hidden value="">
	<INPUT name="NA8" type=hidden value="">
	<INPUT name="NA9" type=hidden value="">
	<INPUT name="NA10" type=hidden value="">
	<INPUT name="NA11" type=hidden value="">
	<INPUT name="NC1" type=hidden value="">
	<INPUT name="NC2" type=hidden value="">
	<INPUT name="NC3" type=hidden value="">
	<INPUT name="NC4" type=hidden value="">
    <INPUT name="ClientCollectID" type=hidden value="<%=ClientCollectID%>">
    <INPUT name="ClientsCollect" type=hidden value="<%=ClientsCollect%>">
    
	<INPUT name="ItemCurrs" 			type=hidden value="<%=ItemCurrs%>">
	<INPUT name="ItemIDs" 				type=hidden value="<%=ItemIDs%>">
	<INPUT name="ItemVals" 				type=hidden value="<%=ItemVals%>">
	<INPUT name="ItemLocs" 				type=hidden value="<%=ItemLocs%>">
	<INPUT name="ItemNames" 			type=hidden value="<%=ItemNames%>">
	<INPUT name="ItemOVals" 			type=hidden value="<%=ItemOVals%>">
	<INPUT name="ItemPPCCs" 			type=hidden value="<%=ItemPPCCs%>">
	<INPUT name="ItemServIDs" 			type=hidden value="<%=ItemServIDs%>">
	<INPUT name="ItemServNames" 		type=hidden value="<%=ItemServNames%>">
    <INPUT name="ItemInvoices" 			type=hidden value="<%=ItemInvoices%>">
    <INPUT name="ItemCalcInBls" 		type=hidden value="<%=ItemCalcInBls%>">
    <INPUT name="ItemIntercompanyIDs" 	type=hidden value="<%=ItemIntercompanyIDs%>">
    <INPUT name="CantItems" 			type=hidden value="<%=CantItems%>">	
	<INPUT name="ItemNames_Routing" 	type=hidden value="<%=ItemNames_Routing%>">	
    
    <INPUT name="CarrierIDAnt" type=hidden value="<%=CarrierIDAnt%>">
    <INPUT name="AWBNumberAnt" type=hidden value="<%=AWBNumberAnt%>">
    <INPUT name="HAWBNumberAnt" type=hidden value="<%=HAWBNumberAnt%>">
    <INPUT type=hidden id="CallRouting" name="CallRouting" value="">
    <INPUT name="No" id="No" type=hidden value="<%=No%>">

    
    <% if Request("Country2") <> "" then %>
        <INPUT name="Country2" type=hidden value="<%=Request("Country2")%>">
    <% end if %>

    <% if Request("Transportista2") <> "" then %>
        <INPUT name="Transportista2" type=hidden value="<%=Request("Transportista2")%>">
    <% end if %>

    <% if Request("AirportDepID2") <> "" then %>
        <INPUT name="AirportDepID2" type=hidden value="<%=Request("AirportDepID2")%>">
    <% end if %>

    <% if Request("AirportDesID2") <> "" then %>
        <INPUT name="AirportDesID2" type=hidden value="<%=Request("AirportDesID2")%>">
    <% end if %>

    <% if Request("BtnMaster2") <> "" then %>
        <INPUT name="BtnMaster2" type=hidden value="<%=Request("BtnMaster2")%>">
    <% end if %>

    <% if Request("AWBNumber2") <> "" then %>
        <INPUT name="AWBNumber2" type=hidden value="<%=Request("AWBNumber2")%>">
    <% end if %>

    <% if Request("HAWBNumber2") <> "" then %>
        <INPUT name="HAWBNumber2" type=hidden value="<%=Request("HAWBNumber2")%>">
    <% end if %>

    <% if Request("BtnReplica2") <> "" then %>
        <INPUT name="BtnReplica2" type=hidden value="<%=Request("BtnReplica2")%>">
    <% end if %>

    <% if Request("awb_frame2") <> "" then %>
        <INPUT name="awb_frame2" type=hidden value="<%=Request("awb_frame2")%>">
    <% end if %>
    



<script type="text/javascript">
    function onChangeTransportista() {
        if ('<%=Countries%>' == 'GT') {
            //document.forma.AWBNumber.style.color = 'silver';            
            //document.forma.AWBNumber.value = '#**#';            
            //document.forma.HAWBNumberAnt.value = '';
            document.forma.TipoMaster.value = 'Ultimo';
            document.forma.TipoHouse.value = '';
        }
        document.forma.AWBNumber.value = '';
        document.forma.HAWBNumber.value = '';
        move();
        document.forma.submit();
    }
</script>




<table width="841" border="1" cellpadding="2" cellspacing="0" align="center">

<tr><td colspan="20" align=center>

    <table width=100% border=0>
    <tr>
    <td width=20%>
    

        <%  'response.write "(" & ObjectID & ")(" & Request("ObjectID2") & ")(" & Request("awb_frame2") & ")"
        
            if ObjectID = 0 then                
                ObjectID2 = CheckNum(Request("ObjectID2")) 
            else
                ObjectID2 = ObjectID                 
            end if
        %>

        <% if ObjectID2 > 0 then 'and ((AWBNumber <> "" and HAWBNumber = "") or (AWBNumber <> "" and HAWBNumber = AWBNumber))then '2018-01-09 %>

            <% 
            if Request("vars") = "" then                 
                ObjectID2 = "OID=" & ObjectID & "&CD=" & CreatedDate & "&CT=" & CreatedTime
            else
                ObjectID2 = Request("vars")
            end if
            %>

            <!-- <input type=button value="<<Frame Principal" class="Boton cBlue" onclick="window.location.href='InsertData.asp?<%=ObjectID2%>&GID=1&AT=<%=AWBType%>&awb_frame2=2'"/> -->
            
            <button value="Frame Principal" class="Boton2 cBlue" onclick="move();window.location.href='InsertData.asp?<%=ObjectID2%>&GID=1&AT=<%=AWBType%>&awb_frame2=2';return false;" title="Retorna Frame Principal"><img src="img/glyphicons_435_undo.png" /> Frame Principal</button>

            		
        <% else %>

            <% if Request("awb_frame2") <> "" then %>

            <!-- <input type=button value="<<Frame Principal" class="Boton cBlue" onclick="window.history.back();"/> -->

            <button value="Frame Principal" class="Boton2 cBlue" onclick="window.history.back();return false;" title="Retorna Frame Principal"><img src="img/glyphicons_435_undo.png" /> Frame Principal</button>


            <% end if %>

        <% end if %>



        <% if ObjectID = 0 And RoutingID > 0 then %>
            <input type=button title="Registro de Routings entregados con errores"
            onclick="window.open(RoutingErrorSite + '?' + RoutingErrorUrl, 'iWinRou', 'location=yes,height=325,width=500,scrollbars=no,status=no,titlebar=no,resizable=no,menubar=no');"
            value="ROUTING A REPORTAR" >
        <% end if %>
  
    </td>
      
    <td width=33% align=center>

        <font color=navy face="Arial"><%=ObjectID & " " & IIF(esMaster,"MASTER","HOUSE")%> :: <%=IIF(No = 1,"Master-Hija",IIF(No > 1,"Hija",""))%> :: <%=IIf(AWBType = 1,"EXPORT","IMPORT")%> :: </font> <INPUT name="Countries" type=text value="<%=Countries%>" readonly size=4 style="border:1px;color:Navy;font-size:16px;"> &nbsp;<br />&nbsp; <INPUT name="file" type=hidden value="<%=file%>" readonly size=30 > 
                 <font color=navy face="Arial" style="font-size:12px"><%=Iif(facturar_a_nombre <> "","FACTURAR A : " & facturar_a & " - " & facturar_a_nombre,"") %></font>       
    </td>




    <td width=10% align=right>

        <!-- <Input type=hidden name=replica value="Directo" readonly size=8> -->

        <!-- esta parte se habilitara despues 2016-12-19 //// 2017-04-19 se publica a resto de la region -->
        <span style="margin:2px;border:1px solid yellow;background:silver">    
            <!--
            <span class="style4">    
                <span style="font-size:10px; font-family:verdana; color:yellow">NEW</span>
                Consolidado o Directo        
            </span>
            -->

            <% if replica <> "" then %>
                <Input Type=Text name=replica value="<%=replica%>" readonly style="border:0px;width:auto"  >
            <% else %>
                <select name="replica" id="replica">
                <option value="">Seleccione</option>
                <option value="Consolidado">Consolidado</option>
                <option value="Directo">Directo</option>
                </select>                
            <% end if %>

        </span> 
    </td>

    </tr>
    </table>

</td></tr>
<td class="style4" align="right" colspan="3">Transportista</td>
    <td class="style4" align="left" bgcolor="#999999" colspan="5">


    <% if Request("Transportista2") <> "" then %> 
        <INPUT name="CarrierID" id="Transportista" type=text value="<%=Request("Transportista2")%>" readonly class="readonly" size=2>
    <% end if %> 

	<select name="CarrierID<% if Request("Transportista2") <> "" then response.write "22" end if %>" onChange="onChangeTransportista()" id="Transportista<% if Request("Transportista2") <> "" then response.write "22" end if %>" style="width:200px" <% if Request("Transportista2") <> "" then %> disabled class="readonly"  <% else %> class="style10" <% end if %> >
	<option value="-1">Seleccionar</option>
	<%
        For i = 0 To CountList1Values-1
    		if aList1Values(0,i) = CarrierID then
	    		CarrierName = aList1Values(1,i)
    			'Countries = aList1Values(2,i)
	    	end if
	%>
	<option value="<%=aList1Values(0,i)%>"><%response.write TranslateCompany(aList1Values(2,i)) & " - " & aList1Values(1,i)%></option>
	<%
   		Next
	%>
	</select>


    <% if Request("Transportista2") <> "" then response.write "<script>selecciona('forma.CarrierID22','" & Request("Transportista2") & "');</script>" end if %>


	
    <INPUT name="TipoMaster" type=hidden value="<%=TipoMaster%>">
    <INPUT name="TipoHouse" type=hidden value="<%=TipoHouse%>">

	</td>
	<td class="style4" align="right">ROUTING</td>
    <td class="style4" align="left" bgcolor="#999999" colspan="2">
	<input name="Routing" value="<%=Routing%>" type="text" size="20" readonly class="readonly">
	
<%        
        'response.write "(" & Request("Routing2") & ")(" & CheckNum(ConsignerID) & ")<BR>" 'and couStr = "1" 

        if Request("Routing2") <> "" AND Request("Routing2") <> "NINGUNO" AND CheckNum(ConsignerID) = 0 then 'si esta habilitado en awb_new.asp el seleccionar routing line 192

            'response.write "SI ENTRO"
            
            response.write "<body onload=window.open('Search_AWBData.asp?GID=17&AT=" & AWBType & "&routing=" & Request("Routing2") & "','AWBData','height=300,width=600,menubar=0,resizable=1,scrollbars=1,toolbar=0');>"

        else 

            'response.write "NO ENTRO"
        end if
%>
		
	<% if Action = 2 And CheckNum(RoutingID) > 0 AND CheckNum(ConsignerID) > 0 Then %>
	
	<% Else %>


        <% 
        couStr = "0"                       
        
        'response.write "[" & Request("Routing2") & "]<br>"

        'response.write "[" & AWBNumber & "][" & HAWBNumber & "][" & replica & "]<br>"

        if AWBNumber <> "" and HAWBNumber <> "" and replica = "Consolidado" then 'es house
            couStr = "1"
        end if
        
        if AWBNumber <> "" and HAWBNumber = AWBNumber and replica = "Directo" then 'es master          
            couStr = "1"
        end if 
        'couStr = "1"

        if ObjectID > 0 and CheckNum(RoutingID) > 0 AND CheckNum(ConsignerID) > 0 then
            couStr = "0"
        end if
        %>

        <% if couStr = "1" then  %>
		<a href="#" onClick="Javascript:window.open('Search_AWBData.asp?GID=17&AT=<%=AWBType%>','AWBData','height=300,width=600,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
        <% end if %>

	<% End If %>
	
    <%if RoutingID<>0 then %>
		&nbsp;&nbsp;&nbsp;<a href="#" onClick="Javascript:window.open('http://10.10.1.20/ventasV2/vendedores/routing_ver.php?id_routing=<%=RoutingID %>', 'routing_ver', 'height=600, width=700, menubar=0, resizable=1, scrollbars=1, toolbar=0');return (false);" class="menu"><font color="FFFFFF"><b>Ver RO</b></font></a>
    <%end if %>
	</td>


    <td class="style4" align="right">    

    Tarifa Minima
        
    <input type="checkbox" name="iMinimo" value="<%=iMinimo%>" <% if iMinimo <> "" and iMinimo <> "1" then %> checked <% end if %> 
    
    <%if RoutingID<>0 then  response.write " disabled " end if %>
    
    >
    
    <%if RoutingID<>0 then %> 

    <input type="hidden" value="<%=iMinimo%>" name="iMinimo" />
    
    <%end if %>

    </td>

 </tr>
  <tr>
    <td class="style4" align="right" colspan="3">No. de Cuenta del Embarcador</td>
    <td class="style4" align="left" bgcolor="#999999" colspan="4">
	<input class="style10" name="AccountShipperNo" value="<%=AccountShipperNo%>" type="text" id="No. de Cuenta del Embarcador" readonly="readonly"></td>
	<td class="style4" align="right" colspan="2">MASTER AWB</td>
    <td class="style4" align="left" bgcolor="#999999" colspan="2".>
	    <input name="AWBNumber" value="<%=AWBNumber%>" id="No. de Master AWB" type="text" size="20" <% if Request("AWBNumber2") <> "" then %> readonly class="readonly" <%else%> class="style10" <%end if%> >
        <% if (CountTableValues = -1 or CarrierID <> CarrierIDAnt) and Countries = "GT" then %>  <br />      
        <!--<input type=submit name="BtnHouse" value="Nuevo" style="font-size:10px;padding:0px">
        <input type=submit name="BtnHouse" value="Ultimo" style="font-size:10px;padding:0px">
        <input type=submit name="BtnHouse" value="Espacio" style="font-size:10px;padding:0px;display:none" disabled> 2017-12-14 -->
        <% end if %>
	</td>
    <td class="style4" align="right">As Agreed<input type="checkbox" onClick="javascript:As_Agreed(this);" name="Agreed" <%if CarrierRates = "AS AGREED" then%> checked <%end if%>></td>
 </tr>
  <tr>
    <td class="style4" align="center" colspan="7">Nombre y Direccion del Embarcador</td>
	<td class="style4" align="right" colspan="2">House AWB</td>
    <td class="style4" align="left" bgcolor="#999999" colspan="2">
        
	    <input name="HAWBNumber" value="<%=HAWBNumber%>" id="No. de House AWB" type="text" size="20" <% if Request("AWBNumber2") <> "" then %> readonly class="readonly" <%else%> class="style10" <%end if%> >
<!--        
        <input name="HAWBNumber" value="<%=HAWBNumber%>" id="No. de House AWB" type="text" size="20"   
            <% if Countries = "GT" then   '1 export 2 import  %>
            <% if Request.Form("BtnHouse") = "Manual" then %>    
                <% if (CountTableValues = -1 or CarrierID <> CarrierIDAnt) and Countries = " " then %>    
                    class="style10"
                <% else %>                
                    readonly class="readonly"
                <% end if %>
            <% else %>                
                readonly class="readonly"
            <% end if %>
            <% end if %>
        >
-->
        <% if (CountTableValues = -1 or CarrierID <> CarrierIDAnt) and Countries = "GT" then %>    <br />  
	    <!-- <input type=submit name="BtnHouse" value="Asignar" style="font-size:10px;padding:0px;" >
        <input type=submit name="BtnHouse" value="Directo" style="font-size:10px;padding:0px;" >
        <input type=submit name="BtnHouse" value="Manual" style="font-size:10px;padding:0px;" > 2017-12-14 -->
        <% end if %>
    </td>
	<td class="style4" align="right" colspan="1">Display&nbsp;Master<input type="checkbox" name="DisplayNumber" <%if DisplayNumber = 1 then%> checked <%end if%> ></td>
  </tr>
  <tr>
    <td class="style4" align="center" bgcolor="#999999" colspan="7"><textarea class="style10" name="ShipperData" rows="5" cols="70" id="Nombre y Direccion del Embarcador"  readonly="readonly"><%=ShipperData%></textarea>
	<table width="100%" cellpadding="0" cellspacing="0">
	<tr>
		<td class="style4" width="80%">&nbsp;</td>
		<td class="style4" align="left">
			<a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF">Nuevo</font></a>
		</td>
		<td class="style4">&nbsp;&nbsp;</td>
		<td class="style4" align="right">
			<a href="#" onClick="Javascript:window.open('Search_AWBData.asp?GID=10','AWBData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
		</td>
	</tr>
	</table>
	</td>
	<td class="style4" align="center" valign="top" bgcolor="#999999" colspan="5" rowspan="2">
        <%if Countries = "GT" then %>
        <table border="1" cellpadding="2" cellspacing="0" width="100%">
        <tr>
        <td bgcolor="#FFFFFF" class="style4" align="right" colspan="2" width="34%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Manifiesto&nbsp;Master</td>
        <td class="style4" align="left" bgcolor="#999999" colspan="2"><input class="style10" name="ManifestNumber" value="<%=ManifestNumber%>" id="ManifiestoMaster" type="text" size="16"></td>
        </tr>
        </table>
        <%end if %>
        <span style="font-size:10px; font-family:verdana; color:yellow"><!--NEW--></span>
        <table style="margin:2px;border:1px solid gray;background:silver"> <!-- 2017-04-19 se publicara sin border yellow -->
            <tr>
                <td class="style4" align="center">Notificar a:</td>
                <td><input id="id_cliente_order" name="id_cliente_order" type="text" readonly  value="<%=id_cliente_order%>" style="background:silver;border:0px;color:gray"/></td>
            </tr>
            <tr>
                <td class="style4" align="center" bgcolor="#999999" colspan="2">                    
                <textarea class="style10" name="id_cliente_orderData" rows="5" cols="70" id="Notificar a:"  readonly="readonly"><%=id_cliente_orderData%></textarea>
                </td>
            </tr>
            <tr>
                <td>

	                <table width="100%" cellpadding="0" cellspacing="0">
	                <tr>
		                <td class="style4" width="80%">&nbsp;</td>
		                <td class="style4" align="left">
			                <a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF">Nuevo</font></a>
		                </td>
		                <td class="style4">&nbsp;&nbsp;</td>
		                <td class="style4" align="right">
                            <a href="#" onClick="Javascript:window.open('Search_AWBData.asp?GID=23','AWBData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
		                </td>
	                </tr>
	                </table>

                </td>
            </tr>
        </table>
       
	</td>
  </tr>
  <tr>
    <td class="style4" align="right" colspan="3">No. de Cuenta del Destinatario </td>
    <td class="style4" align="left" bgcolor="#999999" colspan="4"><input class="style10" name="AccountConsignerNo" value="<%=AccountConsignerNo%>" type="text" id="No. de Cuenta del Destinatario"  readonly="readonly"></td>
  </tr>
  <tr>
    <td class="style4" align="center" colspan="7">Nombre y Direccion del Destinatario </td>
    <% 'NUEVO COLOADER 2015-05-27 %>
    <td class="style4" align="center" colspan="5">Nombre y Direccion del CoLoader</td>
  </tr>
  <tr>
    <td class="style4" align="center" bgcolor="#999999" colspan="7">
	<textarea class="style10" rows="5" name="ConsignerData" id="Nombre y Direccion del Destinatario"  cols="70"  readonly="readonly"><%=ConsignerData%></textarea>
	<table width="100%" cellpadding="0" cellspacing="0">
	<tr>
		<td class="style4" width="80%">&nbsp;</td>
		<td class="style4" align="left">
			<a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF">Nuevo</font></a>
		</td>
		<td class="style4">&nbsp;&nbsp;</td>
		<td class="style4" align="right">
			<a href="#" onClick="Javascript:window.open('Search_AWBData.asp?GID=7','AWBData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
		</td>
	</tr>
	</table>
	</td>

    <% 'NUEVO COLOADER 2015-05-27 %>
    <td class="style4" align="center" bgcolor="#999999" colspan="7">
	<textarea class="style10" rows="5" name="ColoaderData" id="Nombre y Direccion del Coloader"  cols="70"  readonly="readonly"><%=ColoaderData%></textarea>
	<table width="100%" cellpadding="0" cellspacing="0">
	<tr>
		<td class="style4" width="80%">&nbsp;</td>
		<td class="style4" align="left">
			<a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF">Nuevo</font></a>
		</td>
		<td class="style4">&nbsp;&nbsp;</td>
		<td class="style4" align="right">
			<a href="#" onClick="Javascript:window.open('Search_AWBData.asp?GID=22','AWBData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
		</td>
	</tr>
	</table>
	</td>


  </tr>
  <tr>
    <td class="style4" align="center" colspan="7">Agente del Transportista Emisor, Nombre y Ciudad </td>
	<td class="style4" align="center" colspan="7">Informaci&oacute;n Contable</td>
  </tr>
  <tr>
    <td class="style4" align="center" bgcolor="#999999" colspan="7"><textarea class="style10" rows="5" name="AgentData" cols="70" id="Agente del Transportista Emisor, Nombre y Ciudad"  readonly="readonly"><%=AgentData%></textarea>
	<table width="100%" cellpadding="0" cellspacing="0">
	<tr>
		<td class="style4" width="80%">&nbsp;</td>
		<td class="style4" align="left">
			<a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF">Nuevo</font></a>
		</td>
		<td class="style4">&nbsp;&nbsp;</td>
		<td class="style4" align="right">
			<a href="#" onClick="Javascript:window.open('Search_AWBData.asp?GID=8','AWBData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
		</td>
	</tr>
	</table>
	</td>
	<td class="style4" align="center" colspan="7" bgcolor="#999999" valign="top"><textarea class="style10" rows="5" name="AccountInformation" cols="70" id="Informaci&oacute;n Contable"><%=AccountInformation%></textarea></td>
  </tr>
  <tr>
    <td class="style4" align="center" colspan="3">Codigo IATA del Agente </td>
    <td class="style4" align="center" colspan="4">No. de Cuenta del Agente</td>
    <td class="style4" align="center" rowspan="2">Viaje</td>
    <td class="style4" align="left" bgcolor="#999999" colspan="6" rowspan="2"><input class="style10" name="Voyage" value="<%=Voyage%>" type="text" id="Viaje" onKeyUp="res(this,numb);"></td>
  </tr>
  <tr>
    <td class="style4" align="center" bgcolor="#999999" colspan="3"><input class="style10" name="IATANo" value="<%=IATANo%>" type="text" id="Codigo IATA del Agente"  readonly="readonly"></td>
    <td class="style4" align="center" bgcolor="#999999" colspan="4"><input class="style10" name="AccountAgentNo" value="<%=AccountAgentNo%>" id="No. de Cuenta del Agente" type="text"  readonly="readonly"></td>
  </tr>
  <tr>
    <td class="style4" align="center" colspan="3">Aeropuerto de Salida</td>
    <td class="style4" align="center" colspan="4">Ruta Solicitada</td>
    <td class="style4" align="center" bgcolor="#999999" colspan="7" rowspan="2"></td>
  </tr>
  <tr>
    <td class="style4" align="center" bgcolor="#999999" colspan="3">

    <% if Request("AirportDepID2") <> "" then %> 
        <INPUT name="AirportDepID" id="Aeropuerto Salida" type=text value="<%=Request("AirportDepID2")%>" readonly class="readonly" size=2>
    <% end if %> 

	<select name="AirportDepID<% if Request("AirportDepID2") <> "" then response.write "22" end if %>" id="Aeropuerto Salida<% if Request("AirportDepID2") <> "" then response.write "22" end if %>" style="width:200px" <% if Request("AirportDepID2") <> "" then %> disabled class="readonly" <% else %> class="style10"  <% end if %>>
	<option value="-1">Seleccionar</option>
	<%
		For i = 0 To CountList2Values-1
	%>
	<option value="<%=aList2Values(0,i)%>"  <%=IIf(CheckNum(Request("AirportDepID2"))=CheckNum(aList2Values(0,i))," selected ","")%>  ><%response.Write aList2Values(2,i) & " - " & aList2Values(1,i)%></option>
	<%
   		Next
	%>
	</select>

    <% 'if Request("AirportDepID2") <> "" then response.write "<script>selecciona('forma.AirportDepID22','" & Request("AirportDepID2") & "');</script>" end if %>

    </td>
    <td class="style4" align="center" bgcolor="#999999" colspan="4"><input class="style10" name="RequestedRouting" value="<%=RequestedRouting%>" id="Ruta Solicitada" type="text"></td>
  </tr>
  
  <tr>
    <td class="style4" align="center">A</td>
    <td class="style4" align="center">1er. Transportista</td>
    <td class="style4" align="center">A </td>
    <td class="style4" align="center">Por </td>
    <td class="style4" align="center">A </td>
    <td class="style4" align="center">Por </td>
    <td class="style4" align="center">Moneda </td>
    <td class="style4" align="center">Codigo<br>
      Cargos </td>
    <td class="style4" align="center">Peso/Valor </td>
    <td class="style4" align="center">Otros </td>
    <td class="style4" align="center">Valor Declarado<br>
      para Transporte </td>
    <td class="style4" align="center">Valor Declarado<br>
      Aduana </td>
  </tr>
  <tr bgcolor="#999999">
    <td class="style4" align="center"><input class="style10" name="AirportToCode1" value="<%=AirportToCode1%>" id="A" type="text" size="6"></td>
    <td class="style4" align="center" bgcolor="#FFFFFF">
	<%=CarrierName%>
    </td>
    <td class="style4" align="center"><input class="style10" name="AirportToCode2" value="<%=AirportToCode2%>" id="A" type="text" size="6"></td>
    <td class="style4" align="center"><input class="style10" name="CarrierCode2" value="<%=CarrierCode2%>" id="Por" type="text" size="6">
    </td>
    <td class="style4" align="center"><input class="style10" name="AirportToCode3" value="<%=AirportToCode3%>" id="A" type="text" size="6">
    </td>
    <td class="style4" align="center"><input class="style10" name="CarrierCode3" value="<%=CarrierCode3%>" id="Por" type="text" size="6">
    </td>


    <td class="style4" align="center">
	<select class="style10" name="CurrencyID" id="Moneda" <%if Countries  <> "HN" then%>onChange="javascript:SetXchange(this);"<%end if%>>
	<option value="-1">Seleccionar</option>
	<%
		For i = 0 To CountList6Values-1
	%>
	<option value="<%=aList6Values(0,i)%>"><%=aList6Values(1,i)%></option>
	<%
   		Next
	%>
	</select>
    </td>
    <td class="style4" align="center">
	<select class="style10" name="ChargeType" id="Codigo Cargos" onChange="javascript:SumOtherCharges(document.forms[0]);AssignChargeWeight(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
	<option value="1">PPD
	<option value="2">COLL
	</select>
    </td>
    <td class="style4" align="center">
	<select class="style10" name="ValChargeType" id="Peso / Valor" onChange="javascript:SumOtherCharges(document.forms[0]);AssignChargeWeight(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
	<option value="1">PPD
	<option value="2">COLL
	</select>
    </td>
    <td class="style4" align="center">
	<select class="style10" name="OtherChargeType" id="Otros" onChange="javascript:SumOtherCharges(document.forms[0]);AssignChargeWeight(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
	<option value="1">PPD
	<option value="2">COLL
	</select>
    </td>
    <td class="style4" align="center"><input class="style10" name="DeclaredValue" value="<%=DeclaredValue%>" id="Valor Declarado Transporte" type="text" size="10">
    </td>
    <td class="style4" align="center"><input class="style10" name="AduanaValue" value="<%=AduanaValue%>" id="Valor Declarado Aduana" type="text" size="10">
    </td>
  </tr>
  <tr>
    <td class="style4" colspan="8" align="center">Aeropuerto Destino </td>
    <td class="style4" align="center">Vuelo<br>
      Fecha </td>
    <td class="style4" align="center">Vuelo<br>
      Fecha </td>
    <td class="style4" colspan="2">Valor Asegurado </td>
  </tr>
  <tr bgcolor="#999999">
    <td class="style4" colspan="8" align="center">

    <% if Request("AirportDesID2") <> "" then %> 
        <INPUT name="AirportDesID" id="Aeropuerto Destino" type=text value="<%=Request("AirportDesID2")%>" readonly class="readonly" size=2>
    <% end if %> 

	<select name="AirportDesID<% if Request("AirportDesID2") <> "" then response.write "22" end if %>" onChange="CheckRates(document.forms[0]);" id="Aeropuerto Destino<% if Request("AirportDesID2") <> "" then response.write "22" end if %>" <% if Request("AirportDesID2") <> "" then %> disabled class="readonly" <% else %> class="style10" <% end if %> >>
	<option value="-1">Seleccionar</option>
	<%
		For i = 0 To CountList3Values-1
	%>
	<option value="<%=aList3Values(0,i)%>"><%response.write aList3Values(2,i) & " - " & aList3Values(1,i)%></option>
	<%
   		Next
	%>
	</select>

    <% if Request("AirportDesID2") <> "" then response.write "<script>selecciona('forma.AirportDesID22','" & Request("AirportDesID2") & "');</script>" end if %>

    </td>
    <td class="style4"><input class="style10" type="text" size="10" name="FlightDate1" value="<%=FlightDate1%>" id="Vuelo Fecha">
    </td>
    <td class="style4"><input class="style10" type="text" size="10" name="FlightDate2" value="<%=FlightDate2%>" id="Vuelo Fecha">
    </td>
    <td class="style4" colspan="2"><input class="style10" type="text" size="10" name="SecuredValue" value="<%=SecuredValue%>" id="Valor Asegurado">
    </td>
  </tr>
  <tr>
    <td class="style4" colspan="7" align="center">Informaci&oacute;n Manejo</td>
    <td class="style4" colspan="5" align="center">Observaci&oacute;n</td>
  </tr>
  <tr bgcolor="#999999">
    <td height="71" colspan="7" align="center" class="style4"><textarea class="style10" name="HandlingInformation" id="Informacion Manejo" rows="4" cols="70"><%=HandlingInformation%></textarea></td>
    <td class="style4" colspan="5" align="center"><textarea class="style10" name="Observations" id="Observaciones" rows="4" cols="67"><%=Observations%></textarea>
    </td>
  </tr>
  <tr>
    <td class="style4" colspan="3" align="center">Invoice</td>
    <td class="style4" colspan="4" align="center">Lic. de Exportacion</td>
    <td class="style4" colspan="5" align="center">Instrucciones</td>
  </tr>
  <tr bgcolor="#999999">
    <td colspan="3" align="center" class="style4"><input type="text" name="Invoice" id="Invoice" value="<%=Invoice%>" class="style10"></td>
	<td colspan="4" align="center" class="style4"><input type="text" name="ExportLic" value="<%=ExportLic%>" id="Lic. de Exportacion" class="style10"></td>
    <td class="style4" colspan="5" align="center"><textarea class="style10" name="Instructions" id="Instrucciones" rows="2" cols="67"><%=Instructions%></textarea>
    </td>
  </tr>
  <tr>
    <td colspan="12" style="border:0px" cellpadding=0 cellspacing=0>
    
        <table width="100%" height="309" border="1" cellpadding="0" cellspacing="0" align="center">
	        <tr>
		        <td width="72" class="style4" align="center">No. Bultos</td>
		        <td width="72" class="style4" align="center">Peso Bruto</td>
		        <td width="39" class="style4" align="center">kg / lb</td>
		        <td width="68" class="style4" align="center">Commodity Item No. </td>
		        <td width="164" class="style4" align="center">Peso a Cobrar</td>
		        <td width="87" class="style4" align="center">Tarifa / Cargo</td>
		        <td width="73" class="style4" align="center">Total</td>
		        <td width="249" class="style4" align="center">Naturaleza y Cantidad<br>de la Mercancia</td>		
	        </tr>
	        <tr bgcolor="#999999">
	            <td height="279" valign="top" class="style4">
			        <textarea class="style10" cols="15" rows="20" wrap="off"  name="NoOfPieces" id="Numero de Bultos" onBlur="javascript:SumVals(this, document.forms[0].TotNoOfPieces);" onKeyUp="res(this,numb);"><%=NoOfPieces%></textarea>
			        <br>
			        <input class="style10" name="TotNoOfPieces" value="<%=TotNoOfPieces%>"  type="text" size="15" readonly>
		        </td>
		        <td class="style4" valign="top">
			        <textarea class="style10"  cols="15" rows="20" wrap="off" name="Weights" id="Peso Bruto"  onBlur="javascript:SumVals(this, document.forms[0].TotWeight);" onKeyUp="res(this,numb);"><%=Weights%></textarea><br><input class="style10"  type="text" size="15" name="TotWeight" value="<%=TotWeight%>" readonly>
		        </td>
		        <td class="style4" valign="top">
			        <textarea class="style10"  cols="4" rows="20" wrap="off" name="WeightsSymbol" id="Simbolo de Peso"><%=WeightsSymbol%></textarea>
		        </td>
		        <td class="style4" valign="top">
			        <textarea class="style10"  cols="6" rows="20" wrap="off" name="Commodities" id="Codigo Producto" onKeyUp="res(this,numb);" readonly><%=Commodities%></textarea>
			        <input type="hidden" name="CommoditiesTypes" value="<%=CommoditiesTypes%>">
			        <div align="right">
<!-- 2022-04-20			
			        <input name="TotCarrierRate_Routing" value="<%=TotCarrierRate_Routing%>" type="hidden" size="2" readonly>
-->			
			        <% If TotCarrierRate_Routing = "1" Then %>
			
			
			        <% else %>
		
				        <a href="#" onClick="Javascript:window.open('Search_AWBData.asp?GID=11','AWBData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
				
			        <% end If %>
				
			        </div>
		        </td>
		        <td class="style4" valign="top">
			        <input type="hidden" name="TotWeightChargeable" value="<%=TotWeightChargeable%>">		
			        <textarea class="style10"  cols="15" rows="20" wrap="off" name="ChargeableWeights" id="Peso a Cobrar" onBlur="javascript:CalcRate(document.forms[0]);" onChange="javascript:RecalculoPeso(document.forms[0]);" onKeyUp="res(this,numb);"><%=ChargeableWeights%></textarea>            
<!-- 2022-04-20			
                    Air Freight <span class=ids>11</span>
-->
		        </td>
		        <td class="style4" valign="top" align="right">
			        <textarea class="style10"  cols="14" rows="20" wrap="off" name="CarrierRates" onBlur="javascript:CalcRate(document.forms[0]);"><%=CarrierRates%></textarea>
			        
<!-- 2022-04-20			                  
                    <br>
-->
			
			        <% If TotCarrierRate_Routing = "1" Then %>
<!-- 2022-04-20									
				        <input type="text" size="5" class="style10" name="CAF" value="<%=Request.Form("CAF")%>" id="Tipo Moneda de Air Freight" readonly>
-->				
			        <% else %>
<!-- 2022-04-20			
				        <select class="style10" name="CAF" id="Tipo Moneda de Air Freight">
				        <option value="-1">---</option>
				        <%=Currencies%>
				        </select>
-->				
			        <% end If %>
			
		        </td>
		        <td class="style4" valign="top">
			        <textarea class="style10"  cols="15" rows="20" wrap="off" name="CarrierSubTot" readonly><%=CarrierSubTot%></textarea>
<!-- 2022-04-20			
			        <br>
			        <input class="style10" name="TotCarrierRate" value="<%=TotCarrierRate%>" type="text" size="15" readonly onBlur="CheckDoble('AirFreight','AF');">
-->
		        </td>
		        <td class="style4" valign="top" align="left">
			        <textarea class="style10"  cols="44" rows="20" wrap="off" name="NatureQtyGoods" id="Naturaleza y Cantidad de la Mercancia"><%=NatureQtyGoods%></textarea>
<!-- 2022-04-20			
			        <br>
-->			
                    <% '="(" & TotCarrierRate_Routing & ")" %>

			        <% If TotCarrierRate_Routing = "99" Then %>
<!-- 2022-04-20			
				        <input type="hidden" size="5" class="style10" name="TCAF" value="<%=Request.Form("TCAF")%>" id="Hidden1" readonly>
                        <input type="text" size="5" class="style10" name="TCAF_copy" value="<%=IntLoc(CheckNum(Request.Form("TCAF")))%>"  readonly>
-->
			        <% else %>
<!-- 2022-04-20			
				        <select class="style10" name="TCAF" id="Tipo de Cobro de Air Freight">
				        <option value="-1">---</option>
				        <option value="0">INT</option>
				        <option value="1">LOC</option>
				        </select>
-->
			        <% end If %>

				        &nbsp;
			
			        <% If TotCarrierRate_Routing = "1" Then %>		
<!-- 2022-04-20			
				        <input type="hidden" size="5" class="style10" name="TPAF" value="<%=Request.Form("TPAF")%>" id="Hidden2" readonly>
                        <input type="text" size="5" class="style10" name="TPAF_copy" value="<%=PrepColl(CheckNum(Request.Form("TPAF")))%>"  readonly>				
-->
			        <% else %>
<!-- 2022-04-20			
			            <select class="style10" name="TPAF" id="Forma de Pago de Air Freight">
			            <option value="-1">---</option>
			            <option value="0">PREP</option>
			            <option value="1">COLL</option>
			            </select>
-->
			        <% end If %>			

<!-- 2022-04-20			
			        <input type="hidden" name="INVAF" value="0">
-->                                
		        </td> 

	        </tr>
        </table>


    </td>
  </tr>


  <!---- //////////////////////////////////////// CARGOS ///////////////////////////////////////////////////////////  --->


  <tr>
    <td colspan="12" style="border:0px" cellpadding=0 cellspacing=0>

		<table width="100%" border="1" cellpadding="2" cellspacing="0" align="center">
		  <tr>
			<td width="168" align="center" class="style4">PAGADO / Prepaid</td>
			<td width="165" align="center" class="style4">DEBIDO / Collect</td>
			<td width="488" align="center" class="style4">Tipo Carga&nbsp;
			<select class="style10" name="TipoCarga2" id="Tipo de Carga">
			<option value="-1">---</option>
			<%
                        
OpenConn3 Conn

    QuerySelect = "SELECT ""tpt_codigo"", ""tpt_descripcion"", ""tpt_pk"" FROM ""ti_pricing_tipo"" WHERE ""tpt_tipo"" = 'TIPO_CARGA' AND ""tpt_tps_fk"" = '1' ORDER BY ""tpt_descripcion"""
	Set rs = Conn.Execute(QuerySelect)
	Do While Not rs.EOF
		response.write "<option value=" & rs("tpt_codigo") & ">" & rs(0) & "</option>"
		rs.MoveNext
	Loop

CloseOBJs rs, Conn                        
            %>
			</select>	

                &nbsp;&nbsp;&nbsp;&nbsp;
                <INPUT type=button onClick="Javascript:return Open(<%=ObjectID%>, '<%=TipoCarga2%>', <%=AWBType%>)" value="&nbsp;&nbsp;Charges&nbsp;&nbsp;" class="Boton cRed">
					            
            </td>

		  </tr>
		 <tr>
			<td class="style4" align="center" colspan="2">Cargos por Peso</td>
			<td class="style4" align="center">Cargos</td>
		 </tr>
		  <tr>
			<td class="style4" align="center" bgcolor="#999999">
					<input type="text" class="style10" name="TotChargeWeightPrepaid" value="<%=TotChargeWeightPrepaid%>" readonly size="13">
			</td>
			<td class="style4" align="center" bgcolor="#999999">
					<input type="text" class="style10" name="TotChargeWeightCollect" value="<%=TotChargeWeightCollect%>" readonly size="13">
			</td>
			
			<td width="488" rowspan="11" align="center" class="style4">
			<table width="100%" border="0">
			<tr>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Rubro</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Tarifa</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Tarifa Tipo</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Regimen</font>
				</td>

				<td align="center" class="style4">
				<font class="style8">Moneda</font>
				</td>
				<td align="center" class="style4">
				<font class="style8">Monto</font>
				</td>
				<td align="center" class="style4">
				<font class="style8">Int/Loc</font>
				</td>
				<td align="center" class="style4">
				<font class="style8">CC/PP</font>
				</td>
				<td></td>

			</tr>

<!-- 2022-04-20			se traslada para aca el codigo de air freight por orden y mejor entendimiento-->

			<tr>
				<td align="right" class="style4">               
                <%=IIf(replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija","<a href=# onClick=""Javascript:AddCharge2('TotCarrierRate_Routing','AF','CAF',3,11);return false;"">Air Freight","Air Freight&nbsp")%> 
                <span class=ids>11</span></td>

                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="AF_Tarifa" value="" id="Text47" readonly>
				</td>
                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="AF_TarifaTipo" value="" id="Text50" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="AF_Regimen" value="" id="Text48" readonly>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				
					<input name="TotCarrierRate_Routing" value="<%=TotCarrierRate_Routing%>" type="hidden" size="2" readonly>
				
					<% If TotCarrierRate_Routing = "1" Then %>
						
						<input type="text" size="5" class="style10" name="CAF" value="<%=Request.Form("CAF")%>" id="Text49" readonly>
						
					<% else %>

						<select class="style10" name="CAF" id="Tipo Moneda de Air Freight">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
						
					<% end If %>
					
				</td>
				<td align="center" bgcolor="#999999" class="style4">
					<span class="style15">
					
					<% If TotCarrierRate_Routing = "1" Then %>
					
									
						<input type="text" size="8" class="style10" id="Air Freight" name="TotCarrierRate" value="<%=TotCarrierRate%>" readonly>
						
					<% else %>
					
                        <input class="style10" name="TotCarrierRate" id="Air Freight" value="<%=TotCarrierRate%>" type="text" size="8" readonly onBlur="CheckDoble('AirFreight','AF');">

					<% end If %>

					</span>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				
					<% If TotCarrierRate_Routing = "99" Then %>				
						<input type="hidden" size="5" class="style10" name="TCAF" value="<%=Request.Form("TCAF")%>" id="Hidden3" readonly>
						<input type="text" size="5" class="style10" name="TCAF_copy" value="<%=IntLoc(CheckNum(Request.Form("TCAF")))%>"  readonly>				
					<% else %>
						<select class="style10" name="TCAF" id="Tipo de Cobro de Air Freight">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>						
					<% end If %>
					
				</td>
				<td align="right" class="style4" bgcolor="#999999">

					<% If TotCarrierRate_Routing = "1" Then %>				
						<input type="hidden" size="5" class="style10" name="TPAF" value="<%=Request.Form("TPAF")%>" id="Hidden4" readonly>
						<input type="text" size="5" class="style10" name="TPAF_copy" value="<%=PrepColl(CheckNum(Request.Form("TPAF")))%>"  readonly>				
					<% else %>
						<select class="style10" name="TPAF" id="Forma de Pago de Air Freight">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>				
					<% end If %>

					<input type="hidden" name="INVAF" value="0">
				</td>  
			</tr>

			<tr>
				<td align="right" class="style4">
                <%=IIf(replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija","<a href=# onClick=""Javascript:AddCharge2('TerminalFee_Routing','TF','CTF',3,15);return false;"">Terminal Fee&nbsp","Terminal Fee&nbsp")%> 
                <span class=ids>15</span></td>

                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="TF_Tarifa" value="" id="Text31" readonly>
				</td>
                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="TF_TarifaTipo" value="" id="Text51" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="TF_Regimen" value="" id="Text32" readonly>
				</td>


				<td align="right" class="style4" bgcolor="#999999">
				
					<input name="TerminalFee_Routing" value="<%=TerminalFee_Routing%>" type="hidden" size="2" readonly>
				
					<% If TerminalFee_Routing = "1" Then %>
						
						<input type="text" size="5" class="style10" name="CTF" value="<%=Request.Form("CTF")%>" id="Tipo Moneda de Terminal Fee" readonly>
						
					<% else %>

						<select class="style10" name="CTF" id="Tipo Moneda de Terminal Fee">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
						
					<% end If %>
					
				</td>
				<td align="center" bgcolor="#999999" class="style4">
					<span class="style15">
					
					<% If TerminalFee_Routing = "1" Then %>
					
									
						<input type="text" size="8" class="style10" id="Terminal Fee" name="TerminalFee" value="<%=TerminalFee%>" readonly>
						
					<% else %>
					
						<input type="text" size="8" class="style10" id="TerminalFee" name="TerminalFee" value="<%=TerminalFee%>" onKeyUp="res(this,numb);FlgTF=SetFlag(this);" onBlur="javascript:CalcRate(document.forms[0]);CheckDoble('TerminalFee','TF');">
					<% end If %>

					</span>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				
					<% If TerminalFee_Routing = "99" Then %>				
						<input type="hidden" size="5" class="style10" name="TCTF" value="<%=Request.Form("TCTF")%>" id="Tipo de Cobro de Terminal Fee" readonly>
						<input type="text" size="5" class="style10" name="TCTF_copy" value="<%=IntLoc(CheckNum(Request.Form("TCTF")))%>"  readonly>				
					<% else %>
						<select class="style10" name="TCTF" id="Tipo de Cobro de Terminal Fee">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>						
					<% end If %>
					
				</td>
				<td align="right" class="style4" bgcolor="#999999">

					<% If TerminalFee_Routing = "1" Then %>				
						<input type="hidden" size="5" class="style10" name="TPTF" value="<%=Request.Form("TPTF")%>" id="Forma de Pago de Terminal Fee" readonly>
						<input type="text" size="5" class="style10" name="TPTF_copy" value="<%=PrepColl(CheckNum(Request.Form("TPTF")))%>"  readonly>				
					<% else %>
						<select class="style10" name="TPTF" id="Forma de Pago de Terminal Fee">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>				
					<% end If %>

					<input type="hidden" name="INVTF" value="0">
				</td>  
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DETF" style="VISIBILITY: visible;">
					
					<% If TerminalFee_Routing = "1" Then %>
					
						
					<% else %>

						<a href="#" onClick="document.forma.CTF.value=-1;document.forma.TerminalFee.value='';document.forma.TCTF.value=-1;document.forma.TPTF.value=-1;return false;" class="menu"><font color="FFFFFF">X</font></a>		
						
					<% end If %>

					</div>
				</td>

			</tr>
			<tr>
				<td align="right" class="style4">
                <%=IIf(replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija","<a href=# onClick=""Javascript:AddCharge2('CustomFee_Routing','CF','CCF',3,14);return false;"">Custom Fee&nbsp","Custom Fee&nbsp")%> 
                <span class=ids>14</span></td>

                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="CF_Tarifa" value="" id="Text33" readonly>
				</td>
                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="CF_TarifaTipo" value="" id="Text52" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="CF_Regimen" value="" id="Text34" readonly>
				</td>

				<td align="right" class="style4" bgcolor="#999999">

					<input name="CustomFee_Routing" value="<%=CustomFee_Routing%>" type="hidden" size="2" readonly>
				
					<% If CustomFee_Routing = "1" Then %>			

						<input type="text" size="5" class="style10" name="CCF" value="<%=Request.Form("CCF")%>" id="Tipo Moneda de Fuel Custom Fee" readonly>
					
					<% else %>
									
						<select class="style10" name="CCF" id="Tipo Moneda de Fuel Custom Fee">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
						
					<% end If %>				

				</td>			
				<td align="center" class="style4" bgcolor="#999999">
				
					<% If CustomFee_Routing = "1" Then %>
					

						<input type="text" size="8" class="style10" id="Custom Fee" name="CustomFee" value="<%=CustomFee%>" readonly>	
						
					<% else %>
					
						<input type="text" size="8" class="style10" id="Custom Fee" name="CustomFee" value="<%=CustomFee%>" onKeyUp="res(this,numb);FlgCF=SetFlag(this);" onBlur="javascript:CalcRate(document.forms[0]);">	
						
					<% end If %>	

				</td>
				<td align="right" class="style4" bgcolor="#999999">
				
					<% If CustomFee_Routing = "99" Then %>
						<input type="hidden" size="5" class="style10" name="TCCF" value="<%=Request.Form("TCCF")%>" readonly id="Tipo de Cobro de Custom Fee">	
						<input type="text" size="5" class="style10" name="TCCF_copy" value="<%=IntLoc(CheckNum(Request.Form("TCCF")))%>"  readonly>				
					<% else %>			
						<select class="style10" name="TCCF" id="Tipo de Cobro de Custom Fee">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>				
					<% end If %>
					
				</td>
				<td align="right" class="style4" bgcolor="#999999">

					<% If CustomFee_Routing = "1" Then %>

						<input type="hidden" size="5" class="style10" name="TPCF" value="<%=Request.Form("TPCF")%>" readonly id="Forma de Pago de Custom Fee">	
						<input type="text" size="5" class="style10" name="TPCF_copy" value="<%=PrepColl(CheckNum(Request.Form("TPCF")))%>"  readonly>
						
					<% else %>
					
						<select class="style10" name="TPCF" id="Forma de Pago de Custom Fee">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
						
					<% end If %>

					<input type="hidden" name="INVCF" value="0">
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DECF" style="VISIBILITY: visible;">
					
					<% If CustomFee_Routing = "1" Then %>
					

					<% else %>
					
						<a href="#" onClick="document.forma.CCF.value=-1;document.forma.CustomFee.value='';document.forma.TCCF.value=-1;document.forma.TPCF.value=-1;return false;" class="menu"><font color="FFFFFF">X</font></a>		
						
					<% end If %>

					</div>
				</td>

			</tr>
			<tr>
				<td align="right" class="style4">
                <%=IIf(replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija","<a href=# onClick=""Javascript:AddCharge2('FuelSurcharge_Routing','FS','CFS',3,12);return false;"">Fuel Surcharge&nbsp","Fuel Surcharge&nbsp")%> 
                <span class=ids>12</span></td>

                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="FS_Tarifa" value="" id="Text35" readonly>
				</td>
                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="FS_TarifaTipo" value="" id="Text53" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="FS_Regimen" value="" id="Text36" readonly>
				</td>

				<td align="right" class="style4"  bgcolor="#999999">

					<input name="FuelSurcharge_Routing" value="<%=FuelSurcharge_Routing%>" type="hidden" size="2" readonly>
				
					<% If FuelSurcharge_Routing = "1" Then %>
					
						
						<input type="text" size="5" class="style10" name="CFS" value="<%=Request.Form("CFS")%>" id="Tipo Moneda de Fuel Surcharge" readonly>
						
					<% else %>

						<select class="style10" name="CFS" id="Tipo Moneda de Fuel Surcharge">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
						
					<% end If %>
					
				</td>
				<td align="center" class="style4" bgcolor="#999999">
					
					<% If FuelSurcharge_Routing = "1" Then %>
					
						
						<input type="text" size="8" class="style10" id="Fuel Surcharge" name="FuelSurcharge" value="<%=FuelSurcharge%>" readonly>
						
					<% else %>

						<input type="text" size="8" class="style10" id="Fuel Surcharge" name="FuelSurcharge" value="<%=FuelSurcharge%>" onKeyUp="res(this,numb);FlgFS=SetFlag(this);" onBlur="javascript:CalcRate(document.forms[0]);">
						
					<% end If %>
					
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					
					<% If FuelSurcharge_Routing = "99" Then %>				
						<input type="hidden" size="5" class="style10" name="TCFS" value="<%=Request.Form("TCFS")%>" id="Tipo de Cobro de Fuel Surcharge" readonly>
						<input type="text" size="5" class="style10" name="TCFS_copy" value="<%=IntLoc(CheckNum(Request.Form("TCFS")))%>"  readonly>
					<% else %>
						<select class="style10" name="TCFS" id="Tipo de Cobro de Fuel Surcharge">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>				
					<% end If %>
					
				</td>
				<td align="right" class="style4" bgcolor="#999999">

					<% If FuelSurcharge_Routing = "1" Then %>
						
						<input type="hidden" size="5" class="style10" name="TPFS" value="<%=Request.Form("TPFS")%>" id="Forma de Pago de Fuel Surcharge" readonly>
						<input type="text" size="5" class="style10" name="TPFS_copy" value="<%=PrepColl(CheckNum(Request.Form("TPFS")))%>"  readonly>

					<% else %>

						<select class="style10" name="TPFS" id="Forma de Pago de Fuel Surcharge">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
						
					<% end If %>

					<input type="hidden" name="INVFS" value="0">
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEFS" style="VISIBILITY: visible;">
					
					<% If FuelSurcharge_Routing = "1" Then %>
					
						
					<% else %>

						<a href="#" onClick="document.forma.CFS.value=-1;document.forma.FuelSurcharge.value='';document.forma.TCFS.value=-1;document.forma.TPFS.value=-1;return false;" class="menu"><font color="FFFFFF">X</font></a>			
						
					<% end If %>

					</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4">
                <%=IIf(replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija","<a href=# onClick=""Javascript:AddCharge2('SecurityFee_Routing','SF','CSF',3,13);return false;"">Security Fee&nbsp","Security Fee&nbsp")%>                 
                <span class=ids>13</span></td>

                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="SF_Tarifa" value="" id="Text37" readonly>
				</td>
                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="SF_TarifaTipo" value="" id="Text54" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="SF_Regimen" value="" id="Text38" readonly>
				</td>

				<td align="right" class="style4" bgcolor="#999999">

					<input name="SecurityFee_Routing" value="<%=SecurityFee_Routing%>" type="hidden" size="2" readonly>
				
					<% If SecurityFee_Routing = "1" Then %>
					
						
						<input type="text" size="5" class="style10" name="CSF" value="<%=Request.Form("CSF")%>" id="Tipo Moneda de Security Fee" readonly>
						
					<% else %>

						<select class="style10" name="CSF" id="Tipo Moneda de Security Fee">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
						
					<% end If %>
							
				</td>
				<td align="center" class="style4" bgcolor="#999999">
				
					<% If SecurityFee_Routing = "1" Then %>
					
						
						<input type="text" size="8" class="style10" id="Security Fee" name="SecurityFee" value="<%=SecurityFee%>" readonly>
						
					<% else %>

						<input type="text" size="8" class="style10" id="Security Fee" name="SecurityFee" value="<%=SecurityFee%>" onKeyUp="res(this,numb);FlgSF=SetFlag(this);" onBlur="javascript:CalcRate(document.forms[0]);">
						
					<% end If %>
				</td>					
					
				<td align="right" class="style4" bgcolor="#999999">
				
					<% If SecurityFee_Routing = "99" Then %>							
						<input type="hidden" size="5" class="style10" name="TCSF" value="<%=Request.Form("TCSF")%>" id="Tipo de Cobro de Security Fee" readonly>
						<input type="text" size="5" class="style10" name="TCSF_copy" value="<%=IntLoc(CheckNum(Request.Form("TCSF")))%>"  readonly>
					<% else %>
						<select class="style10" name="TCSF" id="Tipo de Cobro de Security Fee">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>			
					<% end If %>
				</td>
				<td align="right" class="style4" bgcolor="#999999">

					<% If SecurityFee_Routing = "1" Then %>
									
						<input type="hidden" size="5" class="style10" name="TPSF" value="<%=Request.Form("TPSF")%>" id="Forma de Pago de Security Fee" readonly>
						<input type="text" size="5" class="style10" name="TPSF_copy" value="<%=PrepColl(CheckNum(Request.Form("TPSF")))%>"  readonly>

					<% else %>

						<select class="style10" name="TPSF" id="Forma de Pago de Security Fee">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					
					<% end If %>

				  <input type="hidden" name="INVSF" value="0">
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DESF" style="VISIBILITY: visible;">				
					<% If SecurityFee_Routing = "1" Then %>
					
					
					<% else %>

						<a href="#" onClick="document.forma.CSF.value=-1;document.forma.SecurityFee.value='';document.forma.TCSF.value=-1;document.forma.TPSF.value=-1;return false;" class="menu"><font color="FFFFFF">X</font></a>		
						
					<% end If %>						
					</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4">
                <%=IIf(replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija","<a href=# onClick=""Javascript:AddCharge2('PBA_Routing','PB','CPB',3,116);return false;"">PBA&nbsp","PBA&nbsp")%>                 
                <span class=ids>116</span></td>

                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="PB_Tarifa" value="" id="Text39" readonly>
				</td>
                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="PB_TarifaTipo" value="" id="Text55" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="PB_Regimen" value="" id="Text40" readonly>
				</td>

				<td align="right" class="style4" bgcolor="#999999">

					<input name="PBA_Routing" value="<%=PBA_Routing%>" type="hidden" size="2" readonly>
					
					<% If PBA_Routing = "1" Then %>
						
						<input type="text" size="5" class="style10" name="CPB" value="<%=Request.Form("CPB")%>" id="Tipo Moneda de PBA" readonly>
						
					<% else %>

						<select name="CPB" class="style10" id="Tipo Moneda de PBA">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
					
					<% end If %>
					
				</td>
				<td align="center" class="style4" bgcolor="#999999">

					<% If PBA_Routing = "1" Then %>
					

						<input type="text" size="8" class="style10" id="PBA" name="PBA" value="<%=PBA%>" readonly>		
						
					<% else %>

						<input type="text" size="8" class="style10" id="PBA" name="PBA" value="<%=PBA%>" onBlur="javascript:ValPBA(document.forms[0]);this.value=Round(this.value);" onKeyUp="res(this,numb);">		
					
					<% end If %>

				</td>
				<td align="right" class="style4" bgcolor="#999999">
				
					<% If PBA_Routing = "99" Then %>				
						<input type="hidden" size="5" class="style10" name="TCPB" value="<%=Request.Form("TCPB")%>" id="Tipo de Cobro de PBA" readonly>
						<input type="text" size="5" class="style10" name="TCPB_copy" value="<%=IntLoc(CheckNum(Request.Form("TCPB")))%>"  readonly>
					<% else %>
						<select class="style10" name="TCPB" id="Tipo de Cobro de PBA">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>			
					<% end If %>			

				</td>
				<td align="right" class="style4" bgcolor="#999999">		

					<% If PBA_Routing = "1" Then %>
						
						<input type="hidden" size="5" class="style10" name="TPPB" value="<%=Request.Form("TPPB")%>" id="Forma de Pago de PBA" readonly>
						<input type="text" size="5" class="style10" name="TPPB_copy" value="<%=PrepColl(CheckNum(Request.Form("TPPB")))%>"  readonly>

					<% else %>

						<select class="style10" name="TPPB" id="Forma de Pago de PBA">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>

					<% end If %>			

					<input type="hidden" name="INVPB" value="0">
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEPB" style="VISIBILITY: visible;">
					
					<% If PBA_Routing = "1" Then %>			
						
					<% else %>
					
						<a href="#" onClick="document.forma.CPB.value=-1;document.forma.PBA.value='';document.forma.TCPB.value=-1;document.forma.TPPB.value=-1;return false;" class="menu"><font color="FFFFFF">X</font></a>
						
					<% end If %>

					</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" valign="middle">
					<%if Countries="HN" then%>
						<input type="checkbox" onClick="javascript:CalcTax(document.forms[0]);CalcTot(document.forms[0]);" name="SetTax" <%if TAX > 0 then%> checked <%end if%>>
					<%end if%>
					TAX	 <span class=ids> </span>
				</td>
                <td></td>
                <td></td>
                <td></td>

				<td align="right" class="style4" bgcolor="#999999">
					<select name="CTX" class="style10" id="Tipo Moneda de TAX">
					<option value="-1">---</option>
					<%=Currencies%>
					</select>
				</td>
				<td align="center" class="style4"  bgcolor="#999999">
					<input type="text" size="8" class="style10" name="TAX" value="<%=TAX%>" onKeyUp="res(this,numb);" readonly="readonly">
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="TCTX" id="Tipo de Cobro de TAX">
					<option value="-1">---</option>
					<option value="0">INT</option>
					<option value="1">LOC</option>
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="TPTX" id="Forma de Pago de TAX">
					<option value="-1">---</option>
					<option value="0">PREP</option>
					<option value="1">COLL</option>
					</select>
					<input type="hidden" name="INVTX" value="0">
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DETX" style="VISIBILITY: visible;">
						<a href="#" onClick="document.forma.CTX.value=-1;document.forma.TAX.value='';document.forma.TCTX.value=-1;document.forma.TPTX.value=-1;return false;" class="menu"><font color="FFFFFF">X</font></a>		
					</div>
				</td>
			</tr>

			<tr>

				<td align="right" class="style4">
                <%=IIf(replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija","<a href=# onClick=""Javascript:AddCharge2('PickUp_Routing','PU','CPU',3,31);return false;"">PickUp&nbsp","PickUp&nbsp")%>
                <span class=ids>31</span></td>

                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="PU_Tarifa" value="" id="Text41" readonly>
				</td>
                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="PU_TarifaTipo" value="" id="Text56" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="PU_Regimen" value="" id="Text42" readonly>
				</td>

				<td align="right" class="style4" bgcolor="#999999">

					<input name="PickUp_Routing" value="<%=PickUp_Routing%>" type="hidden" size="2" readonly>
				
					<% If PickUp_Routing = "1" Then %>
						
						<input type="text" size="5" class="style10" name="CPU" value="<%=Request.Form("CPU")%>" id="Tipo Moneda de Pick-Up" readonly>
						
					<% else %>

						<select class="style10" name="CPU" id="Tipo Moneda de Pick-Up">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>		
						
					<% end If %>
							
				</td>
				<td align="center" class="style4" bgcolor="#999999">
					<% If PickUp_Routing = "1" Then %>
					
						
						<input type="text" size="8" class="style10" id="PickUp" name="PickUp" value="<%=PickUp%>" readonly>
						
					<% else %>

						<input type="text" size="8" class="style10" id="PickUp" name="PickUp" value="<%=PickUp%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<% end If %>		

				</td>
				<td align="right" class="style4" bgcolor="#999999">
				
					<% If PickUp_Routing = "99" Then %>				
						<input type="hidden" size="5" class="style10" name="TCPU" value="<%=Request.Form("TCPU")%>" id="Tipo de Cobro de Pick-Up" readonly>
						<input type="text" size="5" class="style10" name="TCPU_copy" value="<%=IntLoc(CheckNum(Request.Form("TCPU")))%>"  readonly>
					<% else %>
						<select class="style10" name="TCPU" id="Tipo de Cobro de Pick-Up">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>								
					<% end If %>		

				</td>
				<td align="right" class="style4" bgcolor="#999999">

					<% If PickUp_Routing = "1" Then %>
						
						<input type="hidden" size="5" class="style10" name="TPPU" value="<%=Request.Form("TPPU")%>" id="Forma de Pago de Pick-Up" readonly>
						<input type="text" size="5" class="style10" name="TPPU_copy" value="<%=PrepColl(CheckNum(Request.Form("TPPU")))%>"  readonly>

					<% else %>

						<select class="style10" name="TPPU" id="Forma de Pago de Pick-Up">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>				

					<% end If %>		

				  <input type="hidden" name="INVPU" value="0">
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEPU" style="VISIBILITY: visible;">
					
					<% If PickUp_Routing = "1" Then %>
					
						
					<% else %>

						<a href="#" onClick="document.forma.CPU.value=-1;document.forma.PickUp.value='';document.forma.TCPU.value=-1;document.forma.TPPU.value=-1;return false;" class="menu"><font color="FFFFFF">X</font></a>		
						
					<% end If %>

					</div>
				</td>
				
			</tr>
			<tr>
				<td align="right" class="style4">
                <%=IIf(replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija","<a href=# onClick=""Javascript:AddCharge2('Intermodal_Routing','IM','CIM',3,115);return false;"">Intermodal&nbsp","Intermodal&nbsp")%> 
                <span class=ids>115</span></td>

                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="IM_Tarifa" value="" id="Text43" readonly>
				</td>
                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="IM_TarifaTipo" value="" id="Text57" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="IM_Regimen" value="" id="Text44" readonly>
				</td>

				<td align="right" class="style4" bgcolor="#999999">
					
					<input name="Intermodal_Routing" value="<%=Intermodal_Routing%>" type="hidden" size="2" readonly>
					
					<% If Intermodal_Routing = "1" Then %>
					
						<input type="text" size="5" class="style10" name="CIM" value="<%=Request.Form("CIM")%>" id="Tipo Moneda de Intermodal" readonly>
					
					<% else %>
									
						<select class="style10" name="CIM" id="Tipo Moneda de Intermodal">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
					
					<% end If %>			
					
				</td>
				<td align="center" class="style4" bgcolor="#999999">
				
					<% If Intermodal_Routing = "1" Then %>
					

						<input type="text" size="8" class="style10" id="Intermodalid" name="Intermodal" value="<%=Intermodal%>" readonly>
					
					<% else %>
								
						<input type="text" size="8" class="style10" id="Intermodal" name="Intermodal" value="<%=Intermodal%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">			
					
					<% end If %>			

				</td>
				<td align="right" class="style4" bgcolor="#999999">
					
					<% If Intermodal_Routing = "99" Then %>			
						<input type="hidden" size="5" class="style10" name="TCIM" value="<%=Request.Form("TCIM")%>"  id="Tipo de Cobro de Intermodal" readonly>
						<input type="text" size="5" class="style10" name="TCIM_copy" value="<%=IntLoc(CheckNum(Request.Form("TCIM")))%>"  readonly>
					<% else %>						
						<select class="style10" name="TCIM" id="Tipo de Cobro de Intermodal">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>			
					<% end If %>			
				</td>
				<td align="right" class="style4" bgcolor="#999999">

					<% If Intermodal_Routing = "1" Then %>			

						<input type="hidden" size="5" class="style10" name="TPIM" value="<%=Request.Form("TPIM")%>" id="Forma de Pago de Intermodal" readonly>
						<input type="text" size="5" class="style10" name="TPIM_copy" value="<%=PrepColl(CheckNum(Request.Form("TPIM")))%>"  readonly>

					<% else %>
								
						<select class="style10" name="TPIM" id="Forma de Pago de Intermodal">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					
					<% end If %>			

					<input type="hidden" name="INVIM" value="0">
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEIM" style="VISIBILITY: visible;">
					
					<% If Intermodal_Routing = "1" Then %>
					
						
					<% else %>
								
						<a href="#" onClick="document.forma.CIM.value=-1;document.forma.Intermodal.value='';document.forma.TCIM.value=-1;document.forma.TPIM.value=-1;return false;" class="menu"><font color="FFFFFF">X</font></a>		
						
					<% end If %>
						
					</div>
				</td>	
			</tr>
			<tr>
				<td align="right" class="style4">
                <%=IIf(replica = "Master-Hija" or replica = "Hija-Directa" or replica = "Master-Master-Hija","<a href=# onClick=""Javascript:AddCharge2('SedFilingFee_Routing','FF','CFF',3,38);return false;"">Sed Filing Fee&nbsp","Sed Filing Fee&nbsp")%> 
                <span class=ids>38</span></td>

                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="FF_Tarifa" value="" id="Text45" readonly>
				</td>
                <td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="FF_TarifaTipo" value="" id="Text58" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="FF_Regimen" value="" id="Text46" readonly>
				</td>

				<td align="right" class="style4" bgcolor="#999999">
				
					<input name="SedFilingFee_Routing" value="<%=SedFilingFee_Routing%>" type="hidden" size="2" readonly>					
				
					<% If SedFilingFee_Routing = "1" Then %>			
						
						<input type="text" size="5" class="style10" name="CFF" value="<%=Request.Form("CFF")%>" id="Tipo Moneda de Sed Filing Fee" readonly>
						
					<% else %>
					
						<select class="style10" name="CFF" id="Tipo Moneda de Sed Filing Fee">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
					
					<% end If %>
					
				</td>
				<td align="center" class="style4" bgcolor="#999999">
				
					<% If SedFilingFee_Routing = "1" Then %>
					
										
						<input type="text" size="8" class="style10" id="Sed Filing Fee" name="SedFilingFee" value="<%=SedFilingFee%>" readonly>		
						
					<% else %>

						<input type="text" size="8" class="style10" id="Sed Filing Fee" name="SedFilingFee" value="<%=SedFilingFee%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">		
					
					<% end If %>
					
				</td>			
				<td align="right" class="style4" bgcolor="#999999">

					<% If SedFilingFee_Routing = "99" Then %>								
						<input type="hidden" size="5" class="style10" name="TCFF" value="<%=Request.Form("TCFF")%>" id="Tipo de Cobro de Sed Filing Fee" readonly>		
						<input type="text" size="5" class="style10" name="TCFF_copy" value="<%=IntLoc(CheckNum(Request.Form("TCFF")))%>"  readonly>
					<% else %>
						<select class="style10" name="TCFF" id="Tipo de Cobro de Sed Filing Fee">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>		
					<% end If %>
				
				</td>
				<td align="right" class="style4" bgcolor="#999999">

					<% If SedFilingFee_Routing = "1" Then %>
										
						<input type="hidden" size="5" class="style10" value="<%=Request.Form("TPFF")%>" name="TPFF" id="Forma de Pago de Sed Filing Fee" readonly>		
						<input type="text" size="5" class="style10" name="TPFF_copy" value="<%=PrepColl(CheckNum(Request.Form("TPFF")))%>"  readonly>

					<% else %>

						<select class="style10" name="TPFF" id="Forma de Pago de Sed Filing Fee">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
				
					<% end If %>

					<input type="hidden" name="INVFF" value="0">
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEFF" style="VISIBILITY: visible;">
					
					<% If SedFilingFee_Routing = "1" Then %>
					
														
					<% else %>

						<a href="#" onClick="document.forma.CFF.value=-1;document.forma.SedFilingFee.value='';document.forma.TCFF.value=-1;document.forma.TPFF.value=-1;return false;" class="menu"><font color="FFFFFF">X</font></a>
						
					<% end If %>
					
					</div>
				</td>	
			</tr>

			<tr>

				<td align="center" class="style4" colspan="6" rowspan="3" valign="middle"><select class="style10" name="WType" onChange="javascript:CalcRate(document.forms[0]);SumOtherCharges(document.forms[0]);AssignChargeWeight(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1" selected>Calculado con Peso Bruto
					  <option value="2">Calculado con Peso a Cobrar
					  </select><br><br>
					  <select class="style10" name="CalcAdminFee">
					<option value="0">NO Calcular Admin Fee</option>
					<option value="1">SI Calcular Admin Fee</option>
				  </select>
				</td>

			</tr>


			</table>


			<table width="100%" border="0">
			<%if ObjectID<>0 then %>
			<tr>
				<td align=center colspan="13">
					<table width="25%" border="0">
					<tr>
						<td align="center" class="style4">
							Cargos Intercompany
						</td>
					</tr>
					<tr>
						<td align="center" bgcolor="#999999"  class="titlelist">
						<a href="#" onClick="window.open('InterCharges.asp?OID=<%=ObjectID%>&AT=<%=AwbType%>','CargosIntercompany','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1100,height=480,top=170,left=170');" class="menu"><font color="FFFFFF"><b>Editar</b></font></a>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<%end if %>


            <tr>
	            <td align="left" class="style4" colspan="13">
		            Otros Cargos		
                </td>
            </tr>
            <tr>
                <td colspan="2"></td>
	            <td align="center" class="style4" colspan="1">
	                <font class="style8">Servicio</font>
	            </td>
	            <td align="center" class="style4" colspan="1">
                    <font class="style8">Rubro</font>
	            </td>
                <td align="center" class="style4" colspan="1">
                    <font class="style8">Tarifa</font>
                </td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Tarifa Tipo</font>
				</td>
                <td align="center" class="style4" colspan="1">
                    <font class="style8">Regimen</font>
                </td>
                <td></td>
	            <td align="center" class="style4">
                    <font class="style8">Moneda</font>
                </td>
	            <td align="center" class="style4">
                    <font class="style8">Monto</font>
                </td>
	            <td align="center" class="style4">
	                <font class="style8">Int/Loc</font>
                </td>
	            <td align="center" class="style4">
	                <font class="style8">CC/PP</font>
                </td>
                <td align="center" class="style4" colspan=2>
	                <font class="style8">Imprimir</font>
                </td>
            </tr>
            <tr>
	            <td align="right" class="style4" colspan="3">
		            <font class="style8">O1&nbsp;</font>
                    <INPUT name="O1" type=text value="0" readonly class=ids size=4>
		            <input type="text" size="10" class="style10" name="SVNO1" value="" id="SVNO1" readonly>
		            <input type="hidden" name="SVIO1" value="" id="SVIO1" readonly>	  
		            <input type="hidden" name="INVO1" value="0">
	            </td>
	            <td align="right" class="style4" colspan="1">
		            <input name="OtherChargeName1" type="text" class="style10" id="Nombre del Rubro para Otros Cargos" value="<%=OtherChargeName1%>" size="15" readonly>
		            <input name="OtherChargeName1_Routing" value="<%=OtherChargeName1_Routing%>" type="hidden" size="2" readonly>			
                </td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O1_Tarifa" value="" id="Text50" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O1_TarifaTipo" value="" id="Text81" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O1_Regimen" value="" id="Text51" readonly>
				</td>

	            <% If OtherChargeName1_Routing = "1" Then %>                        				
		            <td align="right" class="style4">
			            <div id=DRO1 style="VISIBILITY: visible;">					
			            </div>
		            </td>
		            <td align="right" class="style4" bgcolor="#999999">
			            <input type="text" size="5" class="style10" name="CO1" value="<%=Request.Form("CO1")%>" id="Tipo Moneda de Otros Cargos" readonly>
		            </td>
		            <td align="center" class="style4" bgcolor="#999999">
			            <input name="OtherChargeVal1" type="text" class="style10" id="Valor del Rubro de Otros Cargos" value="<%=OtherChargeVal1%>" size="8" readonly>
		            </td>
                <% Else %>				
		            <td align="right" class="style4">
			            <div id=DRO1 style="VISIBILITY: visible;">
				            <a href="#" onClick="Javascript:AddCharge('OtherChargeName1','O1','CO1');return false;" class="menu"><font color="FFFFFF">Buscar</font></a> 
			            </div>
		            </td>
		            <td align="right" class="style4" bgcolor="#999999">
			            <select class="style10" name="CO1" id="Tipo Moneda de Otros Cargos">
			            <option value="-1">---</option>
			            <%=Currencies%>
			            </select>      
		            </td>
		            <td align="center" class="style4" bgcolor="#999999">
			            <input name="OtherChargeVal1" type="text" class="style10" id="Valor del Rubro de Otros Cargos" onBlur="javascript:SumOtherCharges(document.forms[0]);document.forma.VO1.value=this.value;" onKeyUp="res(this,numb);" value="<%=OtherChargeVal1%>" size="8">      
		            </td>
	            <% End If %>			
		
		
	            <% If OtherChargeName1_Routing = "99" Then %>                        				
		            <td align="right" class="style4" bgcolor="#999999">			
			            <input type="hidden" size="5" class="style10" name="TCO1" id="Tipo de Cobro de Otros Cargos" value="<%=Request.Form("TCO1")%>" readonly>
			            <input type="text" size="5" class="style10" name="TCO1_copy" value="<%=IntLoc(CheckNum(Request.Form("TCO1")))%>"  readonly>
		            </td>		
                <% Else %>				
		            <td align="right" class="style4" bgcolor="#999999">
			            <select class="style10" name="TCO1" id="Tipo de Cobro de Otros Cargos">
			            <option value="-1">---</option>
			            <option value="0">INT</option>
			            <option value="1">LOC</option>
			            </select>      
		            </td>				
	            <% End If %>	



	            <% If OtherChargeName1_Routing = "1" Then %>                        				
		            <td align="right" class="style4" bgcolor="#999999">			
			            <input type="hidden" size="5" class="style10" name="TPO1" id="Forma de Pago de Otros Cargos"  value="<%=Request.Form("TPO1")%>"  readonly>
			            <input type="text" size="5" class="style10" name="TPO1_copy" value="<%=PrepColl(CheckNum(Request.Form("TPO1")))%>"  readonly>
		            </td>		
                <% Else %>				
		            <td align="right" class="style4" bgcolor="#999999">
			            <select class="style10" name="TPO1" id="Forma de Pago de Otros Cargos">
			            <option value="-1">---</option>
			            <option value="0">PREP</option>
			            <option value="1">COLL</option>
			            </select>
		            </td>
	            <% End If %>	
		
                <td align="right" class="style4" bgcolor="#999999">
		            <select class="style10" name="CCBLO1" id="Select1" onChange="javascript:SumOtherCharges(document.forms[0]);">
		            <option value="1">SI</option>
		            <option value="0">NO</option>
		            </select>
	            </td>
	            <td align="right" class="style4" bgcolor="#999999">
		            <div id="DEO1" style="VISIBILITY: visible;">			
	            <% If OtherChargeName1_Routing = "1" Then %>                        				

                <% Else %>
		            <a href="#" onClick="DelOtherCharge(1,1);return(false);" class="menu"><font color="FFFFFF">X</font></a> 
                <% End If %>			
		            </div>
	            </td>
            </tr>

            

			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">O2&nbsp;</font>
					<INPUT name="O2" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNO2" value="" id="SVNO2" readonly>
					<input type="hidden" name="SVIO2" value="" id="SVIO2" readonly>	  
					<input type="hidden" name="INVO2" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input name="OtherChargeName2" type="text" class="style10" id="Text52" value="<%=OtherChargeName2%>" size="15" readonly>
					<input name="OtherChargeName2_Routing" value="<%=OtherChargeName2_Routing%>" type="hidden" size="2" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O2_Tarifa" value="" id="Text53" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O2_TarifaTipo" value="" id="Text82" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O2_Regimen" value="" id="Text54" readonly>
				</td>

				<% If OtherChargeName2_Routing = "1" Then %>                        				
					<td align="right" class="style4">
						<div id=DRO2 style="VISIBILITY: visible;">					
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CO2" value="<%=Request.Form("CO2")%>" id="Text55" readonly>
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal2" type="text" class="style10" id="Text56" value="<%=OtherChargeVal2%>" size="8" readonly>
					</td>

				<% Else %>			
					<td align="right" class="style4">
						<div id=DRO2 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('OtherChargeName2','O2','CO2');return false;" class="menu"><font color="FFFFFF">Buscar</font></a> 
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CO2" id="Select2">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>      
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal2" type="text" class="style10" id="Text57" onBlur="javascript:SumOtherCharges(document.forms[0]);document.forma.VO2.value=this.value;" onKeyUp="res(this,numb);" value="<%=OtherChargeVal2%>" size="8">      
					</td>
				<% End If %>		
					
				<% If OtherChargeName2_Routing = "99" Then %>                        				
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TCO2" id="Hidden1" value="<%=Request.Form("TCO2")%>" readonly>
						<input type="text" size="5" class="style10" name="TCO2_copy" value="<%=IntLoc(CheckNum(Request.Form("TCO2")))%>" readonly>
					</td>		
				<% Else %>			
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCO2" id="Select3">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>      
					</td>
				<% End If %>


				<% If OtherChargeName2_Routing = "1" Then %>                        				
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TPO2" id="Hidden2" value="<%=Request.Form("TPO2")%>" readonly>
						<input type="text" size="5" class="style10" name="TPO2_copy" value="<%=PrepColl(CheckNum(Request.Form("TPO2")))%>"  readonly>
					</td>		
				<% Else %>			
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPO2" id="Select4">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>
					
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLO2" id="Select5" onChange="javascript:SumOtherCharges(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEO2" style="VISIBILITY: visible;">			
				<% If OtherChargeName2_Routing = "1" Then %>                        				

				<% Else %>
						<a href="#" onClick="DelOtherCharge(2,2);return(false);" class="menu"><font color="FFFFFF">X</font></a> 
				<% End If %>		
					</div>
				</td>
			</tr>
            
    

			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">O3&nbsp;</font>
					<INPUT name="O3" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNO3" value="" id="SVNO3" readonly>
					<input type="hidden" name="SVIO3" value="" id="SVIO3" readonly>	  
					<input type="hidden" name="INVO3" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input name="OtherChargeName3" type="text" class="style10" id="OtherChargeName3" value="<%=OtherChargeName3%>" size="15" readonly>
					<input name="OtherChargeName3_Routing" value="<%=OtherChargeName3_Routing%>" type="hidden" size="2" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O3_Tarifa" value="" id="Text58" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O3_TarifaTipo" value="" id="Text83" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O3_Regimen" value="" id="Text59" readonly>
				</td>
					
				<% If OtherChargeName3_Routing = "1" Then %>                        				
					<td align="right" class="style4">
						<div id=DRO3 style="VISIBILITY: visible;">					
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CO3" value="<%=Request.Form("CO3")%>" id="Text60" readonly>
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal3" type="text" class="style10" id="Text61" value="<%=OtherChargeVal3%>" size="8" readonly>
					</td>

				<% Else %>			
					<td align="right" class="style4">
						<div id=DRO3 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('OtherChargeName3','O3','CO3');return false;" class="menu"><font color="FFFFFF">
					Buscar</font></a> 
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CO3" id="Select6">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>      
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal3" type="text" class="style10" id="Text62" onBlur="javascript:SumOtherCharges(document.forms[0]);document.forma.VO3.value=this.value;" onKeyUp="res(this,numb);" value="<%=OtherChargeVal3%>" size="8">      
					</td>
				<% End If %>		
					
				<% If OtherChargeName3_Routing = "99" Then %>                        				
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TCO3" id="Hidden5" value="<%=Request.Form("TCO3")%>" readonly>
						<input type="text" size="5" class="style10" name="TCO3_copy" value="<%=PrepColl(CheckNum(Request.Form("TCO3")))%>" readonly>
					</td>		
				<% Else %>			
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCO3" id="Select7">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>      
					</td>		
				<% End If %>	

				<% If OtherChargeName3_Routing = "1" Then %>                        				
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TPO3" id="Hidden6" value="<%=Request.Form("TPO3")%>" readonly>
						<input type="text" size="5" class="style10" name="TPO3_copy" value="<%=PrepColl(CheckNum(Request.Form("TPO3")))%>" readonly>
					</td>		
				<% Else %>			
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPO3" id="Select8">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>		
				<% End If %>	
					
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLO3" id="Select9" onChange="javascript:SumOtherCharges(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>
					</select>
				</td>	  
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEO3" style="VISIBILITY: visible;">
				<% If OtherChargeName3_Routing = "1" Then %>                        				

				<% Else %>
						<a href="#" onClick="DelOtherCharge(3,3);return(false);" class="menu"><font color="FFFFFF">X</font></a> 
				<% End If %>				
					</div>
				</td>
			</tr>



			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">O4&nbsp;</font>
					<INPUT name="O4" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNO4" value="" id="SVNO4" readonly>
					<input type="hidden" name="SVIO4" value="" id="SVIO4" readonly>	  
					<input type="hidden" name="INVO4" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input name="OtherChargeName4" type="text" class="style10" id="Text63" value="<%=OtherChargeName4%>" size="15" readonly>
					<input name="OtherChargeName4_Routing" value="<%=OtherChargeName4_Routing%>" type="hidden" size="2" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O4_Tarifa" value="" id="Text64" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O4_TarifaTipo" value="" id="Text84" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O4_Regimen" value="" id="Text65" readonly>
				</td>

				<% If OtherChargeName4_Routing = "1" Then %>                        				
					<td align="right" class="style4">
						<div id=DRO4 style="VISIBILITY: visible;">					
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CO4" value="<%=Request.Form("CO4")%>" id="Text66" readonly>
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal4" type="text" class="style10" id="Text67" value="<%=OtherChargeVal4%>" size="8" readonly>
					</td>
				<% Else %>				
					<td align="right" class="style4">
						<div id=DRO4 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('OtherChargeName4','O4','CO4');return false;" class="menu"><font color="FFFFFF">Buscar</font></a> 
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CO4" id="Select10">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>      
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal4" type="text" class="style10" id="Text68" onBlur="javascript:SumOtherCharges(document.forms[0]);document.forma.VO4.value=this.value;" onKeyUp="res(this,numb);" value="<%=OtherChargeVal4%>" size="8">      
					</td>
				<% End If %>	
					
				<% If OtherChargeName4_Routing = "99" Then %>
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TCO4" id="Hidden7" value="<%=Request.Form("TPO4")%>" readonly>
						<input type="text" size="5" class="style10" name="TCO4_copy" value="<%=IntLoc(CheckNum(Request.Form("TCO4")))%>"  readonly>
					</td>		
				<% Else %>				
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCO4" id="Select11">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>      
					</td>		
				<% End If %>	

				<% If OtherChargeName4_Routing = "1" Then %>                        				
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TPO4" id="Hidden8" value="<%=Request.Form("TPO4")%>" readonly>
						<input type="text" size="5" class="style10" name="TPO4_copy" value="<%=PrepColl(CheckNum(Request.Form("TPO4")))%>"  readonly>
					</td>		
				<% Else %>				
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPO4" id="Select12">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>		
				<% End If %>	
					
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLO4" id="Select13" onChange="javascript:SumOtherCharges(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>
					</select>
				</td>	
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEO4" style="VISIBILITY: visible;">				
				<% If OtherChargeName4_Routing = "1" Then %>                        				

				<% Else %>
						<a href="#" onClick="DelOtherCharge(4,4);return(false);" class="menu"><font color="FFFFFF">X</font></a> 
				<% End If %>		
					</div>
				</td>
			</tr>



			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">O5&nbsp;</font>
					<INPUT name="O5" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNO5" value="" id="SVNO5" readonly>
					<input type="hidden" name="SVIO5" value="" id="SVIO5" readonly>	  
					<input type="hidden" name="INVO5" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input name="OtherChargeName5" type="text" class="style10" id="Text69" value="<%=OtherChargeName5%>" size="15" readonly>
					<input name="OtherChargeName5_Routing" value="<%=OtherChargeName5_Routing%>" type="hidden" size="2" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O5_Tarifa" value="" id="Text70" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O5_TarifaTipo" value="" id="Text85" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O5_Regimen" value="" id="Text71" readonly>
				</td>

				<% If OtherChargeName5_Routing = "1" Then %>                        				
					<td align="right" class="style4">
						<div id=DRO5 style="VISIBILITY: visible;">					
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CO5" value="<%=Request.Form("CO5")%>" id="Text72" readonly>
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal5" type="text" class="style10" id="Text73" value="<%=OtherChargeVal5%>" size="8" readonly>
					</td>

				<% Else %>			
					<td align="right" class="style4">
						<div id=DRO5 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('OtherChargeName5','O5','CO5');return false;" class="menu"><font color="FFFFFF">Buscar</font></a> 
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CO5" id="Select14">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>      
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal5" type="text" class="style10" id="Text74" onBlur="javascript:SumOtherCharges(document.forms[0]);document.forma.VO5.value=this.value;" onKeyUp="res(this,numb);" value="<%=OtherChargeVal5%>" size="8">      
					</td>
				<% End If %>			
					
					
				<% If OtherChargeName5_Routing = "99" Then %>
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TCO5" id="Hidden9" value="<%=Request.Form("TCO5")%>" readonly>
						<input type="text" size="5" class="style10" name="TCO5_copy" value="<%=IntLoc(CheckNum(Request.Form("TCO5")))%>"  readonly>
					</td>		
				<% Else %>			
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCO5" id="Select15">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>      
					</td>		
				<% End If %>	


				<% If OtherChargeName5_Routing = "1" Then %>                        				
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TPO5" id="Hidden10" value="<%=Request.Form("TPO5")%>" readonly>
						<input type="text" size="5" class="style10" name="TPO5_copy" value="<%=PrepColl(CheckNum(Request.Form("TPO5")))%>"  readonly>
					</td>		
				<% Else %>			
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPO5" id="Select16">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>	
				<% End If %>	
					
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLO5" id="Select17" onChange="javascript:SumOtherCharges(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>
					</select>
				</td>	  
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEO5" style="VISIBILITY: visible;">			
				<% If OtherChargeName5_Routing = "1" Then %>                        				

				<% Else %>
					<a href="#" onClick="DelOtherCharge(5,5);return(false);" class="menu"><font color="FFFFFF">X</font></a> 
				<% End If %>		
					</div>
				</td>
			</tr>


			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">O6&nbsp;</font>
					<INPUT name="O6" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNO6" id="SVNO6" readonly>
					<input type="hidden" name="SVIO6" value="" id="SVIO6" readonly>	  
					<input type="hidden" name="INVO6" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input name="OtherChargeName6" type="text" class="style10" id="Text75" value="<%=OtherChargeName6%>" size="15" readonly>
					<input name="OtherChargeName6_Routing" value="<%=OtherChargeName6_Routing%>" type="hidden" size="2" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O6_Tarifa" value="" id="Text76" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O6_TarifaTipo" value="" id="Text86" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="O6_Regimen" value="" id="Text77" readonly>
				</td>

				<% If OtherChargeName6_Routing = "1" Then %>                        				
					<td align="right" class="style4">
						<div id=DRO6 style="VISIBILITY: visible;">					
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CO6" value="<%=Request.Form("CO6")%>" id="Text78" readonly>
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal6" type="text" class="style10" id="Text79" value="<%=OtherChargeVal6%>" size="8" readonly>
					</td>
				<% Else %>						
					<td align="right" class="style4">
						<div id=DRO6 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('OtherChargeName6','O6','CO6');return false;" class="menu"><font color="FFFFFF">Buscar</font></a> 
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CO6" id="Select18">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>        
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input name="OtherChargeVal6" type="text" class="style10" id="Text80" onBlur="javascript:SumOtherCharges(document.forms[0]);document.forma.VO6.value=this.value;" onKeyUp="res(this,numb);" value="<%=OtherChargeVal6%>" size="8">
					</td>
				<% End If %>	
					
					
				<% If OtherChargeName6_Routing = "99" Then %>
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TCO6" id="Hidden11" value="<%=Request.Form("TCO6")%>" readonly>
						<input type="text" size="5" class="style10" name="TCO6_copy" value="<%=IntLoc(CheckNum(Request.Form("TCO6")))%>" readonly>
					</td>		
				<% Else %>						
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCO6" id="Select19">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>
					</td>
				<% End If %>	


				<% If OtherChargeName6_Routing = "1" Then %>                        				
					<td align="right" class="style4" bgcolor="#999999">			
						<input type="hidden" size="5" class="style10" name="TPO6" id="Hidden12" value="<%=Request.Form("TPO6")%>" readonly>
						<input type="text" size="5" class="style10" name="TPO6_copy" value="<%=PrepColl(CheckNum(Request.Form("TPO6")))%>"  readonly>
					</td>		
				<% Else %>						
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPO6" id="Select20">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>	    
				<% End If %>	
					
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLO6" id="Select21" onChange="javascript:SumOtherCharges(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id="DEO6" style="VISIBILITY: visible;">				
				<% If OtherChargeName6_Routing = "1" Then %>                        				

				<% Else %>
						<a href="#" onClick="DelOtherCharge(6,6);return(false);" class="menu"><font color="FFFFFF">X</font></a> 
				<% End If %>		
					</div>
				</td>
            </tr>

			<tr>
				<td align="left" class="style4" colspan="13">
					Otros Cargos Agente
				</td>
			</tr>
			<tr>
                <td colspan="2"></td>
				<td align="center" class="style4" colspan="1">
				    <font class="style8">Servicio</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Rubro</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Tarifa</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Tarifa Tipo</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Regimen</font>
				</td>
                <td></td>
				<td align="center" class="style4">
				<font class="style8">Moneda</font>
				</td>
				<td align="center" class="style4">
				<font class="style8">Monto</font>
				</td>
				<td align="center" class="style4">
				<font class="style8">Int/Loc</font>
				</td>
				<td align="center" class="style4">
				<font class="style8">CC/PP</font>
				</td>
				<td align="center" class="style4" colspan=2>
				<font class="style8">Imprimir</font>
				</td>


			</tr>


			<tr>
				<td align="right" class="style4" colspan="3">
				<font class="style8">A1&nbsp;</font>
				<INPUT name="A1" type=text value="0" readonly class=ids size=4>
				<input type="text" size="10" class="style10" name="SVNA1" value="" id="SVNA1" readonly>
				<input type="hidden"  class="style10" name="SVIA1" value="" id="SVIA1" readonly>
				<input type="hidden" name="INVA1" value="0">
				</td>

				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName1" value="<%=AdditionalChargeName1%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName1_Routing" value="<%=AdditionalChargeName1_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A1_Tarifa" value="" id="Text1" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A1_TarifaTipo" value="" id="Text59" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A1_Regimen" value="" id="Text2" readonly>
				</td>


				<% If AdditionalChargeName1_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA1 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA1" value="<%=Request.Form("CA1")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal1" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal1%>" readonly>
					</td>

				<% Else %>
					<td align="right" class="style4">
						<div id=DRA1 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName1','A1','CA1');return false;" class="menu"><font color="FFFFFF">Buscar</font></a> 
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA1" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>        
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal1" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal1%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA1.value=this.value;">
					</td>		    

				<% End If %>

				

				<% If AdditionalChargeName1_Routing = "99" Then %>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA1" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA1")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA1_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA1")))%>"  readonly>
					</td>
				<% Else %>		    
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA1" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>
					</td>
				<% End If %>


				<% If AdditionalChargeName1_Routing = "1" Then %>            
							   
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" value="<%=Request.Form("TPA1")%>" name="TPA1" id="Forma de Pago de Cargos de Agente" readonly>
						<input type="text" size="5" class="style10" name="TPA1_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA1")))%>"  readonly>
					</td> 

				<% Else %>
					
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA1" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>



				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA1" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA1 style="VISIBILITY: visible;">
				<% If AdditionalChargeName1_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(1,1);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>		
				</div>
				</td>


			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A2&nbsp;</font>
					<INPUT name="A2" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA2" value="" id="SVNA2" readonly>
					<input type="hidden"  class="style10" name="SVIA2" value="" id="SVIA2" readonly>
					<input type="hidden" name="INVA2" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName2" value="<%=AdditionalChargeName2%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName2_Routing" value="<%=AdditionalChargeName2_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A2_Tarifa" value="" id="Text3" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A2_TarifaTipo" value="" id="Text60" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A2_Regimen" value="" id="Text4" readonly>
				</td>
				<% If AdditionalChargeName2_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA2 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA2" value="<%=Request.Form("CA2")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal2" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal2%>" readonly>
					</td>
					
				<% Else %>		
					<td align="right" class="style4">
						<div id=DRA2 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName2','A2','CA2');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA2" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal2" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal2%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA2.value=this.value;">
					</td>
					
				<% End If %>
				
				
				
				<% If AdditionalChargeName2_Routing = "99" Then %>            
					
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA2" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA2")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA2_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA2")))%>"  readonly>
					</td>

				<% Else %>		
					
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA2" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>        
					</td>

				<% End If %>


				<% If AdditionalChargeName2_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPA2" id="Forma de Pago de Cargos de Agente" value="<%=Request.Form("TPA2")%>" readonly>
						<input type="text" size="5" class="style10" name="TPA2_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA2")))%>"  readonly>
					</td>
				<% Else %>		

					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA2" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>
				

				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA2" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>
					</select>
				</td> 
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA2 style="VISIBILITY: visible;">				
				<% If AdditionalChargeName2_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(2,2);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>
				</div>
				</td>


			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A3&nbsp;</font>
					<INPUT name="A3" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA3" value="" id="SVNA3" readonly>
					<input type="hidden"  class="style10" name="SVIA3" value="" id="SVIA3" readonly>
					<input type="hidden" name="INVA3" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName6" value="<%=AdditionalChargeName6%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName6_Routing" value="<%=AdditionalChargeName6_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A3_Tarifa" value="" id="Text5" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A3_TarifaTipo" value="" id="Text61" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A3_Regimen" value="" id="Text6" readonly>
				</td>
				<% If AdditionalChargeName6_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA3 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA3" value="<%=Request.Form("CA3")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal6" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal6%>" readonly>
					</td>

				<% Else %>		
					<td align="right" class="style4">
						<div id=DRA3 style="VISIBILITY: visible;">
						<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName6','A3','CA3');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA3" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>        
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal6" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal6%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA3.value=this.value;">        
					</td>

				<% End If %>
				
				
				<% If AdditionalChargeName6_Routing = "99" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA3" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA3")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA3_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA3")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA3" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>        
					</td>

				<% End If %>



				<% If AdditionalChargeName6_Routing = "1" Then %>            
					
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPA3" id="Forma de Pago de Cargos de Agente" value="<%=Request.Form("TPA3")%>" readonly>
						<input type="text" size="5" class="style10" name="TPA3_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA3")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA3" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>
				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA3" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>  
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA3 style="VISIBILITY: visible;">
				<% If AdditionalChargeName6_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(3,6);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>
				
				</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A4&nbsp;</font>
					<INPUT name="A4" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA4" value="" id="SVNA4" readonly>
					<input type="hidden"  class="style10" name="SVIA4" value="" id="SVIA4" readonly>
					<input type="hidden" name="INVA4" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName7" value="<%=AdditionalChargeName7%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName7_Routing" value="<%=AdditionalChargeName7_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A4_Tarifa" value="" id="Text7" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A4_TarifaTipo" value="" id="Text62" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A4_Regimen" value="" id="Text8" readonly>
				</td>
				<% If AdditionalChargeName7_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA4 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA4" value="<%=Request.Form("CA4")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal7" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal7%>" readonly>
					</td>

				<% Else %>		
					<td align="right" class="style4">
						<div id=DRA4 style="VISIBILITY: visible;">
						<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName7','A4','CA4');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>		  
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA4" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>		
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal7" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal7%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA4.value=this.value;">		
					</td>

				<% End If %>							
				
				
				
				<% If AdditionalChargeName7_Routing = "99" Then %>            		

					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA4" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA4")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA4_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA4")))%>"  readonly>
					</td>

				<% Else %>		

					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA4" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>		
					</td>

				<% End If %>



				<% If AdditionalChargeName7_Routing = "1" Then %>            
							   
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPA4" id="Forma de Pago de Cargos de Agente" value="<%=Request.Form("TPA4")%>" readonly>
						<input type="text" size="5" class="style10" name="TPA4_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA4")))%>"  readonly>
					</td>
				<% Else %>		
					
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA4" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>
				
				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA4" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA4 style="VISIBILITY: visible;">
				<% If AdditionalChargeName7_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(4,7);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>
				</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A5&nbsp;</font>
					<INPUT name="A5" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA5" value="" id="SVNA5" readonly>
					<input type="hidden"  class="style10" name="SVIA5" value="" id="SVIA5" readonly>
					<input type="hidden" name="INVA5" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName9" value="<%=AdditionalChargeName9%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName9_Routing" value="<%=AdditionalChargeName9_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A5_Tarifa" value="" id="Text9" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A5_TarifaTipo" value="" id="Text63" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A5_Regimen" value="" id="Text10" readonly>
				</td>
				<% If AdditionalChargeName9_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA5 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA5" value="<%=Request.Form("CA5")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal9" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal9%>" readonly>
					</td>

				<% Else %>		
					<td align="right" class="style4">
						<div id=DRA5 style="VISIBILITY: visible;">
						<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName9','A5','CA5');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>		
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA5" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal9" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal9%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA5.value=this.value;">
					</td>

				<% End If %>				
				
				
				<% If AdditionalChargeName9_Routing = "99" Then %>            

					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA5" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA5")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA5_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA5")))%>"  readonly>
					</td>

				<% Else %>		

					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA5" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>		
					</td>

				<% End If %>	


				<% If AdditionalChargeName9_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" value="<%=Request.Form("TPA5")%>" name="TPA5" id="Forma de Pago de Cargos de Agente" readonly>
						<input type="text" size="5" class="style10" name="TPA5_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA5")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA5" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>	

				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA5" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA5 style="VISIBILITY: visible;">
				<% If AdditionalChargeName9_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(5,9);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>
				</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A6&nbsp;</font>
					<INPUT name="A6" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA6" value="" id="SVNA6" readonly>
					<input type="hidden"  class="style10" name="SVIA6" value="" id="SVIA6" readonly>
					<input type="hidden" name="INVA6" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName10" value="<%=AdditionalChargeName10%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName10_Routing" value="<%=AdditionalChargeName10_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A6_Tarifa" value="" id="Text11" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A6_TarifaTipo" value="" id="Text64" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A6_Regimen" value="" id="Text12" readonly>
				</td>
				<% If AdditionalChargeName10_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA6 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA6" value="<%=Request.Form("CA6")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal10" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal10%>" readonly>
					</td>
				<% Else %>				
					<td align="right" class="style4">
						<div id=DRA6 style="VISIBILITY: visible;">
						<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName10','A6','CA6');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>		
						</div>
					</td>
						<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA6" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>		
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal10" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal10%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA6.value=this.value;">
					</td>
				<% End If %>		
				
				
				<% If AdditionalChargeName10_Routing = "99" Then %>            
					
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA6" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA6")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA6_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA6")))%>"  readonly>
					</td>
				<% Else %>				
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA6" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>		
					</td>

				<% End If %>


				<% If AdditionalChargeName10_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPA6" id="Forma de Pago de Cargos de Agente" value="<%=Request.Form("TPA6")%>" readonly>
						<input type="text" size="5" class="style10" name="TPA6_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA6")))%>"  readonly>
					</td>
				<% Else %>				
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA6" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>
				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA6" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA6 style="VISIBILITY: visible;">
				<% If AdditionalChargeName10_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(6,10);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>		
				</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A7&nbsp;</font>
					<INPUT name="A7" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA7" value="" id="SVNA7" readonly>
					<input type="hidden"  class="style10" name="SVIA7" value="" id="SVIA7" readonly>
					<input type="hidden" name="INVA7" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName11" value="<%=AdditionalChargeName11%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName11_Routing" value="<%=AdditionalChargeName11_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A7_Tarifa" value="" id="Text13" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A7_TarifaTipo" value="" id="Text65" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A7_Regimen" value="" id="Text14" readonly>
				</td>
				<% If AdditionalChargeName11_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA7 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA7" value="<%=Request.Form("CA7")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal11" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal11%>" readonly>
					</td>
				<% Else %>			
					<td align="right" class="style4">
						<div id=DRA7 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName11','A7','CA7');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA7" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>		
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal11" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal11%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA7.value=this.value;">		
					</td>
				<% End If %>				
				
				<% If AdditionalChargeName11_Routing = "99" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA7" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA7")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA7_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA7")))%>"  readonly>
					</td>
				<% Else %>			
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA7" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>		
					</td>
				<% End If %>			


				<% If AdditionalChargeName11_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPA7" id="Forma de Pago de Cargos de Agente" value="<%=Request.Form("TPA7")%>" readonly>
						<input type="text" size="5" class="style10" name="TPA7_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA7")))%>"  readonly>
					</td>
				<% Else %>			
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA7" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>			

				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA7" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA7 style="VISIBILITY: visible;">
				<% If AdditionalChargeName11_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(7,11);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>
				</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A8&nbsp;</font>
					<INPUT name="A8" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA8" value="" id="SVNA8" readonly>
					<input type="hidden"  class="style10" name="SVIA8" value="" id="SVIA8" readonly>
					<input type="hidden" name="INVA8" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName12" value="<%=AdditionalChargeName12%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName12_Routing" value="<%=AdditionalChargeName12_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A8_Tarifa" value="" id="Text15" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A8_TarifaTipo" value="" id="Text66" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A8_Regimen" value="" id="Text16" readonly>
				</td>
				<% If AdditionalChargeName12_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA1 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA8" value="<%=Request.Form("CA8")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal12" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal12%>" readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4">
						<div id=DRA8 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName12','A8','CA8');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA8" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>		  
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal12" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal12%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA8.value=this.value;">		
					</td>
				<% End If %>				
				
				
				<% If AdditionalChargeName12_Routing = "99" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA8" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA8")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA8_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA8")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA8" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>		
					</td>		
				<% End If %>		



				<% If AdditionalChargeName12_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPA8" id="Forma de Pago de Cargos de Agente" value="<%=Request.Form("TPA8")%>" readonly>
						<input type="text" size="5" class="style10" name="TPA8_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA8")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA8" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>		
				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA8" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA8 style="VISIBILITY: visible;">
				<% If AdditionalChargeName12_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(8,12);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>
				</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A9&nbsp;</font>
					<INPUT name="A9" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA9" value="" id="SVNA9" readonly>
					<input type="hidden"  class="style10" name="SVIA9" value="" id="SVIA9" readonly>
					<input type="hidden" name="INVA9" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName13" value="<%=AdditionalChargeName13%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName13_Routing" value="<%=AdditionalChargeName13_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A9_Tarifa" value="" id="Text17" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A9_TarifaTipo" value="" id="Text67" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A9_Regimen" value="" id="Text18" readonly>
				</td>
				<% If AdditionalChargeName13_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA9 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA9" value="<%=Request.Form("CA9")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal13" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal13%>" readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4">
						<div id=DRA9 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName13','A9','CA9');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>		  
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA9" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>		  
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal13" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal13%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA9.value=this.value;">		
					</td>
				<% End If %>				
				
				
				<% If AdditionalChargeName13_Routing = "99" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA9" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA9")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA9_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA9")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA9" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>		
					</td>		

				<% End If %>		


				<% If AdditionalChargeName13_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPA9" id="Forma de Pago de Cargos de Agente" value="<%=Request.Form("TPA9")%>" readonly>
						<input type="text" size="5" class="style10" name="TPA9_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA9")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA9" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>		

				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA9" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA9 style="VISIBILITY: visible;">
				<% If AdditionalChargeName13_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(9,13);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>
				</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A10&nbsp;</font>
					<INPUT name="A10" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA10" value="" id="SVNA10" readonly>
					<input type="hidden"  class="style10" name="SVIA10" value="" id="SVIA10" readonly>
					<input type="hidden" name="INVA10" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName14" value="<%=AdditionalChargeName14%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName14_Routing" value="<%=AdditionalChargeName14_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A10_Tarifa" value="" id="Text19" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A10_TarifaTipo" value="" id="Text68" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A10_Regimen" value="" id="Text20" readonly>
				</td>
				<% If AdditionalChargeName14_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA10 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA10" value="<%=Request.Form("CA10")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal14" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal14%>" readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4">
						<div id=DRA10 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName14','A10','CA10');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>		  
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA10" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>		  
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal14" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal14%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA10.value=this.value;">		
					</td>
				<% End If %>				
				
				
				
				<% If AdditionalChargeName14_Routing = "99" Then %>            
					
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA10" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA10")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA10_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA10")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA10" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>		
					</td>		

				<% End If %>



				<% If AdditionalChargeName14_Routing = "1" Then %>            
					
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPA10" id="Forma de Pago de Cargos de Agente" value="<%=Request.Form("TPA10")%>" readonly>
						<input type="text" size="5" class="style10" name="TPA10_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA10")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA10" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>


				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA10" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA10 style="VISIBILITY: visible;">
				<% If AdditionalChargeName14_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(10,14);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>		
				</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">A11&nbsp;</font>
					<INPUT name="A11" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNA11" value="" id="SVNA11" readonly>
					<input type="hidden"  class="style10" name="SVIA11" value="" id="SVIA11" readonly>
					<input type="hidden" name="INVA11" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName15" value="<%=AdditionalChargeName15%>" id="Nombre del Rubro de Agente" readonly>
					<input name="AdditionalChargeName15_Routing" value="<%=AdditionalChargeName15_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A11_Tarifa" value="" id="Text21" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A11_TarifaTipo" value="" id="Text69" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="A11_Regimen" value="" id="Text22" readonly>
				</td>
				<% If AdditionalChargeName15_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRA11 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CA11" value="<%=Request.Form("CA11")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal15" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal15%>" readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4">
					  <div id=DRA11 style="VISIBILITY: visible;">
						<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName15','A11','CA11');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>		  
					  </div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CA11" id="Tipo Moneda de Cargos de Agente">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>		  
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal15" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal15%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VA11.value=this.value;">		
					</td>
				<% End If %>				
				
				<% If AdditionalChargeName15_Routing = "99" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCA11" id="Tipo de Cobro de Cargos de Agente" value="<%=Request.Form("TCA11")%>" readonly>
						<input type="text" size="5" class="style10" name="TCA11_copy" value="<%=IntLoc(CheckNum(Request.Form("TCA11")))%>"  readonly>
					</td>
				<% Else %>		

					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCA11" id="Tipo de Cobro de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>		
					</td>		
				<% End If %>

				<% If AdditionalChargeName15_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPA11" id="Forma de Pago de Cargos de Agente" value="<%=Request.Form("TPA11")%>" readonly>
						<input type="text" size="5" class="style10" name="TPA11_copy" value="<%=PrepColl(CheckNum(Request.Form("TPA11")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPA11" id="Forma de Pago de Cargos de Agente">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>
				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLA11" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEA11 style="VISIBILITY: visible;">
				<% If AdditionalChargeName15_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?')) DelAgentCharge(11,15);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>
				</div>
				</td>
			</tr>

			<tr>
				<td align="left" class="style4" colspan="13">
					Otros Cargos Transportista
				</td>
			</tr>
			<tr>
                <td colspan="2"></td>
				<td align="center" class="style4" colspan="1">
				<font class="style8">Servicio</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Rubro</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Tarifa</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Tarifa Tipo</font>
				</td>
				<td align="center" class="style4" colspan="1">
    				<font class="style8">Regimen</font>
				</td>
                <td></td>
				<td align="center" class="style4">
				<font class="style8">Moneda</font>
				</td>
				<td align="center" class="style4">
				<font class="style8">Monto</font>
				</td>
				<td align="center" class="style4">
				<font class="style8">Int/Loc</font>
				</td>
				<td align="center" class="style4">
				<font class="style8">CC/PP</font>
				</td>
				<td align="center" class="style4" colspan=2>
				<font class="style8">Imprimir</font>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">C1&nbsp;</font>
					<INPUT name="C1" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNC1" value="" id="SVNC1" readonly>
					<input type="hidden"  class="style10" name="SVIC1" value="" id="SVIC1" readonly>
					<input type="hidden" name="INVC1" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName3" id="Nombre del Rubro del Transportista" value="<%=AdditionalChargeName3%>" readonly>
					<input name="AdditionalChargeName3_Routing" value="<%=AdditionalChargeName3_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C1_Tarifa" value="" id="Text23" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C1_TarifaTipo" value="" id="Text70" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C1_Regimen" value="" id="Text24" readonly>
				</td>
				<% If AdditionalChargeName3_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRC1 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CC1" value="<%=Request.Form("CC1")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal3" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal3%>" readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4">
						<div id=DRC1 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName3','C1','CC1');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
						</div>
						</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CC1" id="Tipo Moneda de Cargos de Transportista">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>        
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal3" id="Valor del Rubro del Transportista" value="<%=AdditionalChargeVal3%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VC1.value=this.value;">
					</td>
				<% End If %>				
				
				
				<% If AdditionalChargeName3_Routing = "99" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCC1" id="Tipo de Cobro de Cargos de Transportista" value="<%=Request.Form("TCC1")%>" readonly>
						<input type="text" size="5" class="style10" name="TCC1_copy" value="<%=IntLoc(CheckNum(Request.Form("TCC1")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCC1" id="Tipo de Cobro de Cargos de Transportista">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>
					</td>		
				<% End If %>				



				<% If AdditionalChargeName3_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPC1" id="Forma de Pago de Cargos de Transportista" value="<%=Request.Form("TPC1")%>" readonly>
						<input type="text" size="5" class="style10" name="TPC1_copy" value="<%=PrepColl(CheckNum(Request.Form("TPC1")))%>"  readonly>
					</td>
				<% Else %>		
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPC1" id="Forma de Pago de Cargos de Transportista">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>				

				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLC1" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>  
				<td align="right" class="style4" bgcolor="#999999">
				<div id=DEC1 style="VISIBILITY: visible;">
				<% If AdditionalChargeName3_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?'))  DelCarrierCharge(1,3);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>		
				</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">C2&nbsp;</font>
					<INPUT name="C2" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNC2" value="" id="SVNC2" readonly>
					<input type="hidden"  class="style10" name="SVIC2" value="" id="SVIC2" readonly>
					<input type="hidden" name="INVC2" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					&nbsp;<input type="text" size="15" class="style10" name="AdditionalChargeName4" id="Nombre del Rubro del Transportista" value="<%=AdditionalChargeName4%>" readonly>
					<input name="AdditionalChargeName4_Routing" value="<%=AdditionalChargeName4_Routing%>" type="hidden" size="2">
				</td>		
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C2_Tarifa" value="" id="Text25" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C2_TarifaTipo" value="" id="Text71" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C2_Regimen" value="" id="Text26" readonly>
				</td>
				<% If AdditionalChargeName4_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRC2 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CC2" value="<%=Request.Form("CC2")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal4" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal4%>" readonly>
					</td>

				<% Else %>				
					<td align="right" class="style4">
						<div id=DRC2 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName4','C2','CC2');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CC2" id="Tipo Moneda de Cargos de Transportista">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>        
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal4" id="Valor del Rubro del Transportista" value="<%=AdditionalChargeVal4%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VC2.value=this.value;">        
					</td>

				<% End If %>		
				
				
				<% If AdditionalChargeName4_Routing = "99" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCC2" id="Tipo de Cobro de Cargos de Transportista" value="<%=Request.Form("TCC2")%>" readonly>
						<input type="text" size="5" class="style10" name="TCC2_copy" value="<%=IntLoc(CheckNum(Request.Form("TCC2")))%>"  readonly>
					</td>
				<% Else %>				
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCC2" id="Tipo de Cobro de Cargos de Transportista">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>        
					</td>		

				<% End If %>



				<% If AdditionalChargeName4_Routing = "1" Then %>            		
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPC2" id="Forma de Pago de Cargos de Transportista" value="<%=Request.Form("TPC2")%>" readonly>
						<input type="text" size="5" class="style10" name="TPC2_copy" value="<%=PrepColl(CheckNum(Request.Form("TPC2")))%>"  readonly>
					</td>
				<% Else %>				
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPC2" id="Forma de Pago de Cargos de Transportista">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td> 
				<% End If %>

				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLC2" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td> 
				
				<td align="right" class="style4" bgcolor="#999999">
					<div id=DEC2 style="VISIBILITY: visible;">	
					
				<% If AdditionalChargeName4_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?'))  DelCarrierCharge(2,4);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>					
				
					</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">C3&nbsp;</font>
					<INPUT name="C3" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNC3" value="" id="SVNC3" readonly>
					<input type="hidden"  class="style10" name="SVIC3" value="" id="SVIC3" readonly>
					<input type="hidden" name="INVC3" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName5" id="Nombre del Rubro del Transportista" value="<%=AdditionalChargeName5%>" readonly>
					<input name="AdditionalChargeName5_Routing" value="<%=AdditionalChargeName5_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C3_Tarifa" value="" id="Text27" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C3_TarifaTipo" value="" id="Text72" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C3_Regimen" value="" id="Text28" readonly>
				</td>
				<% If AdditionalChargeName5_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRC3 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CC3" value="<%=Request.Form("CC3")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal5" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal5%>" readonly>
					</td>
					
				<% Else %>						
					<td align="right" class="style4">
						<div id=DRC3 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName5','C3','CC3');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CC3" id="Tipo Moneda de Cargos de Transportista">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal5" id="Valor del Rubro del Transportista" value="<%=AdditionalChargeVal5%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VC3.value=this.value;">        
					</td>

				<% End If %>			
				
				
				<% If AdditionalChargeName5_Routing = "99" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCC3" id="Tipo de Cobro de Cargos de Transportista" value="<%=Request.Form("TCC3")%>" readonly>
						<input type="text" size="5" class="style10" name="TCC3_copy" value="<%=IntLoc(CheckNum(Request.Form("TCC3")))%>"  readonly>
					</td>
				<% Else %>						
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCC3" id="Tipo de Cobro de Cargos de Transportista">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>        
					</td>
				<% End If %>			


				<% If AdditionalChargeName5_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPC3" id="Forma de Pago de Cargos de Transportista" value="<%=Request.Form("TPC3")%>" readonly>
						<input type="text" size="5" class="style10" name="TPC3_copy" value="<%=PrepColl(CheckNum(Request.Form("TPC3")))%>"  readonly>
					</td>
				<% Else %>						
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPC3" id="Forma de Pago de Cargos de Transportista">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>			

				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLC3" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>
				<td align="right" class="style4" bgcolor="#999999">
					<div id=DEC3 style="VISIBILITY: visible;">
				<% If AdditionalChargeName5_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?'))  DelCarrierCharge(3,5);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>	
				
					</div>
				</td>
			</tr>
			<tr>
				<td align="right" class="style4" colspan="3">
					<font class="style8">C4&nbsp;</font>
					<INPUT name="C4" type=text value="0" readonly class=ids size=4>
					<input type="text" size="10" class="style10" name="SVNC4" value="" id="SVNC4" readonly>
					<input type="hidden"  class="style10" name="SVIC4" value="" id="SVIC4" readonly>
					<input type="hidden" name="INVC4" value="0">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="15" class="style10" name="AdditionalChargeName8" id="Nombre del Rubro del Transportista" value="<%=AdditionalChargeName8%>" readonly>
					<input name="AdditionalChargeName8_Routing" value="<%=AdditionalChargeName8_Routing%>" type="hidden" size="2">
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C4_Tarifa" value="" id="Text29" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C4_TarifaTipo" value="" id="Text73" readonly>
				</td>
				<td align="right" class="style4" colspan="1">
					<input type="text" size="6" class="style10" name="C4_Regimen" value="" id="Text30" readonly>
				</td>
				<% If AdditionalChargeName8_Routing = "1" Then %>            
					
					<td align="right" class="style4"><div id=DRC4 style="VISIBILITY: visible;"></div></td>
					<td align="right" class="style4" bgcolor="#999999">
						<input type="text" size="5" class="style10" name="CC4" value="<%=Request.Form("CC4")%>" id="Tipo Moneda de Cargos de Agente" readonly>
					</td>
					<td width="112" align="center" bgcolor="#999999" class="style4">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal8" id="Valor del Rubro de Agente" value="<%=AdditionalChargeVal8%>" readonly>
					</td>
				<% Else %>				
					<td align="right" class="style4">
						<div id=DRC4 style="VISIBILITY: visible;">
							<a href="#" onClick="Javascript:AddCharge('AdditionalChargeName8','C4','CC4');return false;" class="menu"><font color="FFFFFF">Buscar</font></a>
						</div>
					</td>
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="CC4" id="Tipo Moneda de Cargos de Transportista">
						<option value="-1">---</option>
						<%=Currencies%>
						</select>        
					</td>
					<td align="center" class="style4" bgcolor="#999999">
						<input type="text" size="8" class="style10" name="AdditionalChargeVal8" id="Valor del Rubro del Transportista" value="<%=AdditionalChargeVal8%>" onKeyUp="res(this,numb);" onBlur="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);document.forma.VC4.value=this.value;">        
					</td>
				<% End If %>					
				
				
				<% If AdditionalChargeName8_Routing = "99" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TCC4" id="Tipo de Cobro de Cargos de Transportista" value="<%=Request.Form("TCC4")%>" readonly>
						<input type="text" size="5" class="style10" name="TCC4_copy" ivalue="<%=IntLoc(CheckNum(Request.Form("TCC4")))%>"  readonly>
					</td>
				<% Else %>				
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TCC4" id="Tipo de Cobro de Cargos de Transportista">
						<option value="-1">---</option>
						<option value="0">INT</option>
						<option value="1">LOC</option>
						</select>        
					</td>
				<% End If %>					

				<% If AdditionalChargeName8_Routing = "1" Then %>            
					<td align="right" class="style4" bgcolor="#999999">
						<input type="hidden" size="5" class="style10" name="TPC4" id="Forma de Pago de Cargos de Transportista" value="<%=Request.Form("TPC4")%>" readonly>
						<input type="text" size="5" class="style10" name="TPC4_copy" value="<%=PrepColl(CheckNum(Request.Form("TPC4")))%>"  readonly>
					</td>
				<% Else %>				
					<td align="right" class="style4" bgcolor="#999999">
						<select class="style10" name="TPC4" id="Forma de Pago de Cargos de Transportista">
						<option value="-1">---</option>
						<option value="0">PREP</option>
						<option value="1">COLL</option>
						</select>
					</td>
				<% End If %>					

				
				<td align="right" class="style4" bgcolor="#999999">
					<select class="style10" name="CCBLC4" id="Calcular en la Guia" onChange="javascript:SumOtherCharges(document.forms[0]);CalcTax(document.forms[0]);CalcTot(document.forms[0]);">
					<option value="1">SI</option>
					<option value="0">NO</option>				
					</select>
				</td>   
				<td align="right" class="style4" bgcolor="#999999">
					<div id=DEC4 style="VISIBILITY: visible;">								
			   <% If AdditionalChargeName5_Routing = "1" Then %>
					

				<% Else %>
					<a href="#" onClick="if (confirm('¿ Confirme Borrar Este Cargo ?'))  DelCarrierCharge(4,8);return(false);" class="menu"><font color="FFFFFF">X</font></a>
				<% End If %>		
					</div>
				</td>
			</tr>
			</table>
			</td>
		  </tr>
		 <tr>
			<td class="style4" align="center" colspan="2">Cargos por Valor</td>
		 </tr>
		  <tr>
			<td class="style4" align="center" bgcolor="#999999"><input name="TotChargeValuePrepaid" type="text" class="style10" value="<%=TotChargeValuePrepaid%>" size="13" readonly></td>
			<td class="style4" align="center" bgcolor="#999999"><input name="TotChargeValueCollect" type="text" class="style10" value="<%=TotChargeValueCollect%>" size="13" readonly></td>
		  </tr>
		   <tr>
			<td class="style4" align="center" colspan="2">Impuestos</td>
		 </tr>
		  <tr>
			<td class="style4" align="center" bgcolor="#999999"><input name="TotChargeTaxPrepaid" type="text" class="style10" value="<%=TotChargeTaxPrepaid%>" size="13" readonly></td>
			<td class="style4" align="center" bgcolor="#999999"><input name="TotChargeTaxCollect" type="text" class="style10" value="<%=TotChargeTaxCollect%>" size="13" readonly></td>
		  </tr>
		   <tr>
			<td class="style4" align="center" colspan="2">Total Otros Cargos a Pagar<br>Al Agente</td>
		 </tr>
		  <tr>
			<td class="style4" align="center" bgcolor="#999999"><input name="AnotherChargesAgentPrepaid" type="text" class="style10" value="<%=AnotherChargesAgentPrepaid%>" size="13" readonly></td>
			<td class="style4" align="center" bgcolor="#999999"><input name="AnotherChargesAgentCollect" type="text" class="style10" value="<%=AnotherChargesAgentCollect%>" size="13" readonly></td>
		  </tr>
		  <tr>
			<td class="style4" align="center" colspan="2">Total Otros Cargos a Pagar<br>Al Transportista</td>
		 </tr>
		  <tr>
			<td class="style4" align="center" bgcolor="#999999"><input name="AnotherChargesCarrierPrepaid" type="text" class="style10" value="<%=AnotherChargesCarrierPrepaid%>" size="13" readonly></td>
			<td class="style4" align="center" bgcolor="#999999"><input name="AnotherChargesCarrierCollect" type="text" class="style10" value="<%=AnotherChargesCarrierCollect%>" size="13" readonly></td>
		  </tr>
		  <tr>
			<td class="style4" align="center">TOTAL PAGADO</td>
			<td class="style4" align="center">TOTAL DEBIDO</td>
		  </tr>
		  <tr>
			<td class="style4" align="center" bgcolor="#999999"><input name="TotPrepaid" type="text" class="style10" value="<%=TotPrepaid%>" size="13" readonly></td>
			<td class="style4" align="center" bgcolor="#999999"><input name="TotCollect" type="text" class="style10" value="<%=TotCollect%>" size="13" readonly></td>
		  </tr>
		</table>
    </td>
  </tr>
  <tr>
    <td colspan="12" style="border:0px" cellpadding=0 cellspacing=0>

        <table width="100%" border="1" cellpadding="2" cellspacing="0" align="center">
        <td class="style4" align="center">Vendedor</td>
            <td class="style4" align="left" bgcolor="#999999">
	        <select name="SalespersonID" class=label id="Vendedor">
		        <option value="-1">Seleccionar</option>
		        <%
			        For i = 0 To CountList8Values
		        %>
		        <option value="<%=aList8Values(0,i)%>"><%=aList8Values(1,i) & " - " & aList8Values(2,i)%></option>
		        <%
    		        Next
		        %>
	        </select>
	        </td>
		        <td align="right" class="style4">
			        Firma Embarcador o de su Agente
		        </td>
		        <td align="left" bgcolor="#999999" class="style4">
			        <input type="text" class="style10" name="AgentSignature" id="Firma del Embarcador o su Agente" value="<%=AgentSignature%>">
		        </td>
	        </tr>
	        <tr>
		        <td align="right" class="style4">
			        Fecha
		        </td>
		        <td align="left" class="style4" bgcolor="#999999">
			        <input type="text" class="style10" name="AWBDate" value="<%=ConvertDate(AWBDate,5)%>" id="Fecha" readonly><a href="JavaScript:abrir('AWBDate');" class="menu"><font color="#FFFFFF">Seleccionar</font></a>
		        </td>
		        <td align="right" class="style4">
			        Firma Transportista o de su Agente
		        </td>
		        <td align="left" class="style4" bgcolor="#999999">
			        <input type="text" class="style10" name="AgentContactSignature" value="<%=AgentContactSignature%>" id="Firma Transportista Emisor o su Agente">
		        </td>
        </table>
    </td>
  </tr>


  <tr>
    <td colspan="12" style="border:0px" cellpadding=0 cellspacing=0>


        <table width="100%" border="0" cellpadding="2" cellspacing="0" align="center">
			<%if CountTableValues = -1 then%>
			 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class="Boton cBlue"></TD>
			<%else
				'if NoOfPieces <> "" and Weights <> "" and WeightsSymbol <> "" and Commodities <> "" and ChargeableWeights <> "" and CarrierRates <> "" and CarrierSubTot <> "" and NatureQtyGoods <> "" then
			%>	 
			 <TD class=label align=center colspan=2>
			 <INPUT name="Expired" type=hidden value="<%if Expired = 0 then response.Write "on" else response.Write "" %>">

			 <% 'response.write "(" & flg_master & ")(" & flg_totals & ")" %>
			 
			 <% if (flg_master = "0" and flg_totals = "0") or replica = "Master-Master-Hija" then '2017-12-08 %>

			 <span class=erpLab>
			 Leyenda Pie de Pagina
			 <select id=leyenda name=leyenda>
					<option value="">-- Seleccione --</option>
					<option value="1">ORIGINAL #1 (FOR ISSUING CARRIER)</option>
					<option value="2">ORIGINAL #2 (FOR CONSIGNEE)</option>
					<option value="3">ORIGINAL #3 (FOR SHIPPER)</option>
					<option value="4">COPY 4 (DELIVERY RECEIPT)</option>
					<option value="5">COPY 5 (FOR AIRPORT OF DESTINATION)</option>
					<option value="6">COPY 6 (FOR THIRD CARRIER)</option>
					<option value="7">COPY 7 (FOR SECOND CARRIER)</option>
					<option value="8">COPY 8 (FOR FIRST CARRIER)</option>
					<option value="9">COPY 9 (FOR AGENT)</option>
					<option value="10">COPY 10 (EXTRA COPY FOR CARRIER)</option>
					<option value="11">COPY 11 (EXTRA COPY FOR CARRIER)</option>
					<option value="12">COPY 12 (EXTRA COPY FOR CARRIER)</option>
			 </select>
			 </span>

			 <INPUT name=enviar type=button onClick="Javascript:window.open('AWBPrint.asp?Action=4&AWBID=<%=ObjectID%>&CID=<%=CarrierID%>&AT=<%=AwbType%>','AWBPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;" value="&nbsp;&nbsp;Previsualizar Plantilla&nbsp;&nbsp;" class="Boton cBlue">
			 <INPUT name=enviar type=button onClick="if (document.getElementById('leyenda').value != '') { window.open('air-waybill-2.asp?Action=4&AWBID=<%=ObjectID%>&CID=<%=CarrierID%>&AT=<%=AwbType%>&L=' + document.getElementById('leyenda').value,'AWBPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750'); } else { alert('Seleccione Leyenda Pie de Pagina'); } return false;" value="&nbsp;&nbsp;Form New&nbsp;&nbsp;" class="Boton cRed">
			 </TD>
			 <%	else %>
			 <INPUT name=enviar type=button value="&nbsp;&nbsp;Previsualizar Plantilla&nbsp;&nbsp;" class="cGray" disabled></TD>
			 <%	end if %>
			 
			 
			 <%	'end if%>
				<%if Closed=0 then%>
			
					 <% if flg_master = "0" and flg_totals = "0" then '2017-12-08 %>
					<TD class=label align=center><INPUT name=enviar type=button onClick="javascript:if(confirm('Si Actualiza y Cierra ya no podra hacer modificaciones y la informacion continuara su proceso')){document.forma.Closed.value=1;validar(2);};" value="&nbsp;Cerrar&nbsp;" class="Boton cBlue"></TD>
					 <%	else %>
					<TD class=label align=center><INPUT name=enviar type=button  value="&nbsp;Cerrar&nbsp;" class="cGray" disabled></TD>
					 <%	end if %>

					<TD class=label align=center><INPUT name=enviar type=button onClick="Javascript:validar(2);" value="&nbsp;Actualizar&nbsp;" class="Boton cBlue"></TD>
					<% if ObjectID<>9030 and ObjectID<>24788 and ObjectID<>9376 then %>
					
					<%      if iEliminar = 1 then '2018-01-23 %>
					<TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:if(confirm('¿ Esta seguro de Eliminar esta Guia ?')) { validar(3); }" value="&nbsp;Eliminar&nbsp;" class="Boton cBlue">
					<%      else %>
					<TD class=label align=center><INPUT name=enviar type=button disabled value="&nbsp;Eliminar&nbsp;" class="cGray">
					<%      end if '2018-02-01%>
					
					<% end if%>
					<input type=hidden name="eliminar" value=1>
					</TD>
				<%else%>
					<%if Session("OperatorLevel") = 0 then%>
					<TD class=label align=center><INPUT name=enviar type=button onClick="Javascript:document.forma.Closed.value=0;validar(2);" value="&nbsp;&nbsp;Abrir&nbsp;&nbsp;" class="Boton cBlue"></TD>
					<%end if%>
				<%end if%>
			<%end if%>

        </table>
		
    </td>
  </tr>

</table>




</form>



<script>


<%if CountList9Values>=0 then
	j = 1
	k = 1
    l = 1
	for i=0 to CountList9Values
		if aList9Values(6,i) = 0 then
            'No se puede borrar la guia cuando ya tiene facturas relacionadas
            if aList9Values(10,i)<>0 then%>
            
                document.forma.eliminar.value = 0;
            <%end if

			select Case aList9Values(0,i)
			case 11 'Air Freight%>
				document.forma.CAF.value = "<%=aList9Values(2,i)%>";
				document.forma.elements["AF_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["AF_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["AF_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["VAF"].value = "<%=aList9Values(4,i)%>";
				document.forma.TCAF.value = <%=aList9Values(3,i)%>;
				document.forma.TPAF.value = <%=aList9Values(9,i)%>;
				document.forma.INVAF.value = <%=aList9Values(10,i)%>;
				if (document.forma.TCAF_copy) document.forma.TCAF_copy.value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
				if (document.forma.TPAF_copy) document.forma.TPAF_copy.value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.Weights.readOnly = 'true';
					document.forma.ChargeableWeights.readOnly = 'true';
					document.forma.CarrierRates.readOnly = 'true';
					document.forma.CarrierSubTot.readOnly = 'true';
					document.forma.TotCarrierRate.readOnly = 'true';
					document.forma.TotWeight.readOnly = 'true';
					document.forma.CAF.disabled = 'false';
					document.forma.TCAF.disabled = 'false';
					document.forma.TPAF.disabled = 'false';
				<%end if%>
			<%case 12 'Fuel Surcharge%>
				document.forma.CFS.value = "<%=aList9Values(2,i)%>";
				document.forma.elements["FS_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["FS_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["FS_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["VFS"].value = "<%=aList9Values(4,i)%>";
				document.forma.TCFS.value = <%=aList9Values(3,i)%>;
				document.forma.TPFS.value = <%=aList9Values(9,i)%>;
				document.forma.INVFS.value = <%=aList9Values(10,i)%>;
				if (document.forma.TCFS_copy) document.forma.TCFS_copy.value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
				if (document.forma.TPFS_copy) document.forma.TPFS_copy.value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.CFS.disabled = 'false';
					document.forma.FuelSurcharge.readOnly = 'true';
					document.forma.TCFS.disabled = 'false';
					document.forma.TPFS.disabled = 'false';
					document.getElementById("DEFS").style.visibility = "hidden";
				<%end if%>
			<%case 13 'Security Charge%>
				document.forma.CSF.value = "<%=aList9Values(2,i)%>";
				document.forma.elements["SF_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["SF_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["SF_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["VSF"].value = "<%=aList9Values(4,i)%>";
				document.forma.TCSF.value =<%=aList9Values(3,i)%>;
				document.forma.TPSF.value =<%=aList9Values(9,i)%>;
				document.forma.INVSF.value = <%=aList9Values(10,i)%>;
				if (document.forma.TCSF_copy) document.forma.TCSF_copy.value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
				if (document.forma.TPSF_copy) document.forma.TPSF_copy.value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.CSF.disabled = 'false';
					document.forma.SecurityFee.readOnly = 'true';
					document.forma.TCSF.disabled = 'false';
					document.forma.TPSF.disabled = 'false';
					document.getElementById("DESF").style.visibility = "hidden";
				<%end if%>
			<%case 14 'Custom Fee%>
				document.forma.CCF.value = "<%=aList9Values(2,i)%>";
				document.forma.elements["CF_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["CF_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["CF_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["VCF"].value = "<%=aList9Values(4,i)%>";
				document.forma.TCCF.value = <%=aList9Values(3,i)%>;
				document.forma.TPCF.value = <%=aList9Values(9,i)%>;
				document.forma.INVCF.value = <%=aList9Values(10,i)%>;
				if (document.forma.TCCF_copy) document.forma.TCCF_copy.value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
				if (document.forma.TPCF_copy) document.forma.TPCF_copy.value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.CCF.disabled = 'false';
					document.forma.CustomFee.readOnly = 'true';
					document.forma.TCCF.disabled = 'false';
					document.forma.TPCF.disabled = 'false';
					document.getElementById("DECF").style.visibility = "hidden";
				<%end if%>
			<%case 15 'Terminal Fee%>
				document.forma.CTF.value = "<%=aList9Values(2,i)%>";
				document.forma.elements["TF_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["TF_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["TF_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["VTF"].value = "<%=aList9Values(4,i)%>";
				document.forma.TCTF.value = <%=aList9Values(3,i)%>;
				document.forma.TPTF.value = <%=aList9Values(9,i)%>;
				document.forma.INVTF.value = <%=aList9Values(10,i)%>;
				if (document.forma.TCTF_copy) document.forma.TCTF_copy.value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
				if (document.forma.TPTF_copy) document.forma.TPTF_copy.value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.CTF.disabled = 'false';
					document.forma.TerminalFee.readOnly = 'true';
					document.forma.TCTF.disabled = 'false';
					document.forma.TPTF.disabled = 'false';
					document.getElementById("DETF").style.visibility = "hidden";
				<%end if%>
			<%case 31 'Pick Up%>
				document.forma.CPU.value = "<%=aList9Values(2,i)%>";
				document.forma.elements["PU_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["PU_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["PU_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["VPU"].value = "<%=aList9Values(4,i)%>";
				document.forma.TCPU.value = <%=aList9Values(3,i)%>;
				document.forma.TPPU.value = <%=aList9Values(9,i)%>;
				document.forma.INVPU.value = <%=aList9Values(10,i)%>;
				if (document.forma.TCPU_copy) document.forma.TCPU_copy.value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
				if (document.forma.TPPU_copy) document.forma.TPPU_copy.value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.CPU.disabled = 'false';
					document.forma.PickUp.readOnly = 'true';
					document.forma.TCPU.disabled = 'false';
					document.forma.TPPU.disabled = 'false';
					document.getElementById("DEPU").style.visibility = "hidden";
				<%end if%>
			<%case 38 'Sed (Sed Filling Fee)%>
				document.forma.CFF.value = "<%=aList9Values(2,i)%>";
				document.forma.elements["FF_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["FF_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["FF_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["VFF"].value = "<%=aList9Values(4,i)%>";
				document.forma.TCFF.value = <%=aList9Values(3,i)%>;
				document.forma.TPFF.value = <%=aList9Values(9,i)%>;
				document.forma.INVFF.value = <%=aList9Values(10,i)%>;
				if (document.forma.TCFF_copy) document.forma.TCFF_copy.value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
				if (document.forma.TPFF_copy) document.forma.TPFF_copy.value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.CFF.disabled = 'false';
					document.forma.SedFilingFee.readOnly = 'true';
					document.forma.TCFF.disabled = 'false';
					document.forma.TPFF.disabled = 'false';
					document.getElementById("DEFF").style.visibility = "hidden";
				<%end if%>
			<%case 115 'Intermodal%>
				document.forma.CIM.value = "<%=aList9Values(2,i)%>";
				document.forma.elements["IM_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["IM_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["IM_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["VIM"].value = "<%=aList9Values(4,i)%>";
				document.forma.TCIM.value = <%=aList9Values(3,i)%>;
				document.forma.TPIM.value = <%=aList9Values(9,i)%>;
				document.forma.INVIM.value = <%=aList9Values(10,i)%>;
				if (document.forma.TCIM_copy) document.forma.TCIM_copy.value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
				if (document.forma.TPIM_copy) document.forma.TPIM_copy.value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.CIM.disabled = 'false';
					document.forma.Intermodal.readOnly = 'true';
					document.forma.TCIM.disabled = 'false';
					document.forma.TPIM.disabled = 'false';
					document.getElementById("DEIM").style.visibility = "hidden";
				<%end if%>
			<%case 116 'PBA%>
				document.forma.CPB.value = "<%=aList9Values(2,i)%>";
				document.forma.elements["PB_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["PB_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["PB_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["VPB"].value = "<%=aList9Values(4,i)%>";
				document.forma.TCPB.value = <%=aList9Values(3,i)%>;
				document.forma.TPPB.value = <%=aList9Values(9,i)%>;
				document.forma.INVPB.value = <%=aList9Values(10,i)%>;
				if (document.forma.TCPB_copy) document.forma.TCPB_copy.value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
				if (document.forma.TPPB_copy) document.forma.TPPB_copy.value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.CPB.disabled = 'false';
					document.forma.PBA.readOnly = 'true';
					document.forma.TCPB.disabled = 'false';
					document.forma.TPPB.disabled = 'false';
					document.getElementById("DEPB").style.visibility = "hidden";
				<%end if%>
			<%end select
		else 'Cargos de Agente o Transportista

			select case aList9Values(1,i) '0=Carrier, 1=Agente, 2=Otros
			  case 0%>
				document.forma.elements["C<%=aList9Values(6,i)%>"].value = "<%=aList9Values(0,i)%>";
				document.forma.elements["C<%=aList9Values(6,i)%>_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["C<%=aList9Values(6,i)%>_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["C<%=aList9Values(6,i)%>_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["CC<%=aList9Values(6,i)%>"].value = "<%=aList9Values(2,i)%>";
				document.forma.elements["TCC<%=aList9Values(6,i)%>"].value = "<%=aList9Values(3,i)%>";
				document.forma.elements["VC<%=aList9Values(6,i)%>"].value = "<%=aList9Values(4,i)%>";
				document.forma.elements["NC<%=aList9Values(6,i)%>"].value = "<%=aList9Values(5,i)%>";
				document.forma.elements["SVIC<%=aList9Values(6,i)%>"].value = "<%=aList9Values(7,i)%>";
				document.forma.elements["SVNC<%=aList9Values(6,i)%>"].value = "<%=aList9Values(8,i)%>";
				document.forma.elements["TPC<%=aList9Values(6,i)%>"].value = "<%=aList9Values(9,i)%>";
				document.forma.elements["INVC<%=aList9Values(6,i)%>"].value = "<%=aList9Values(10,i)%>";
                document.forma.elements["CCBLC<%=aList9Values(6,i)%>"].value = "<%=aList9Values(11,i)%>";
                if (document.forma.elements["TCC<%=aList9Values(6,i)%>_copy"])
                    document.forma.elements["TCC<%=aList9Values(6,i)%>_copy"].value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
                if (document.forma.elements["TPC<%=aList9Values(6,i)%>_copy"])
                    document.forma.elements["TPC<%=aList9Values(6,i)%>_copy"].value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
				<% if aList9Values(10,i) <> 0 then%>
					document.forma.elements["C<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["CC<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["TCC<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["VC<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["NC<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["SVIC<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["SVNC<%=aList9Values(6,i)%>"].readOnly = 'true';
					document.forma.elements["TPC<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["INVC<%=aList9Values(6,i)%>"].disabled = 'false';
                    document.forma.elements["CCBLC<%=aList9Values(6,i)%>"].disabled = 'false';
					
					document.forma.elements["AdditionalChargeName"+CarriersPos[<%=aList9Values(6,i)%>]].readOnly = 'true';
					document.forma.elements["AdditionalChargeVal"+CarriersPos[<%=aList9Values(6,i)%>]].readOnly = 'true';
					document.getElementById("DEC<%=aList9Values(6,i)%>").style.visibility = "hidden";
					document.getElementById("DRC<%=aList9Values(6,i)%>").style.visibility = "hidden";
				<%end if%>
				<%j=j+1%>
			<%case 1%>
				document.forma.elements["A<%=aList9Values(6,i)%>"].value = "<%=aList9Values(0,i)%>";
				document.forma.elements["A<%=aList9Values(6,i)%>_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["A<%=aList9Values(6,i)%>_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["A<%=aList9Values(6,i)%>_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["CA<%=aList9Values(6,i)%>"].value = "<%=aList9Values(2,i)%>";
				document.forma.elements["TCA<%=aList9Values(6,i)%>"].value = "<%=aList9Values(3,i)%>";
				document.forma.elements["VA<%=aList9Values(6,i)%>"].value = "<%=aList9Values(4,i)%>";
				document.forma.elements["NA<%=aList9Values(6,i)%>"].value = "<%=aList9Values(5,i)%>";
				document.forma.elements["SVIA<%=aList9Values(6,i)%>"].value = "<%=aList9Values(7,i)%>";
				document.forma.elements["SVNA<%=aList9Values(6,i)%>"].value = "<%=aList9Values(8,i)%>";
				document.forma.elements["TPA<%=aList9Values(6,i)%>"].value = "<%=aList9Values(9,i)%>";
				document.forma.elements["INVA<%=aList9Values(6,i)%>"].value = "<%=aList9Values(10,i)%>";
				document.forma.elements["CCBLA<%=aList9Values(6,i)%>"].value = "<%=aList9Values(11,i)%>";
                if (document.forma.elements["TCA<%=aList9Values(6,i)%>_copy"])
                    document.forma.elements["TCA<%=aList9Values(6,i)%>_copy"].value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
                if (document.forma.elements["TPA<%=aList9Values(6,i)%>_copy"])
                    document.forma.elements["TPA<%=aList9Values(6,i)%>_copy"].value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';
                <% if aList9Values(10,i) <> 0 then%>
					document.forma.elements["A<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["CA<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["TCA<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["VA<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["NA<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["SVIA<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["SVNA<%=aList9Values(6,i)%>"].readOnly = 'true';
					document.forma.elements["TPA<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["INVA<%=aList9Values(6,i)%>"].disabled = 'false';
                    document.forma.elements["CCBLA<%=aList9Values(6,i)%>"].disabled = 'false';

					document.forma.elements["AdditionalChargeName"+AgentsPos[<%=aList9Values(6,i)%>]].readOnly = 'true';
					document.forma.elements["AdditionalChargeVal"+AgentsPos[<%=aList9Values(6,i)%>]].readOnly = 'true';
					document.getElementById("DEA<%=aList9Values(6,i)%>").style.visibility = "hidden";
					document.getElementById("DRA<%=aList9Values(6,i)%>").style.visibility = "hidden";
				<%end if%>
				<%k=k+1
			case 2%>
                document.forma.elements["O<%=aList9Values(6,i)%>"].value = "<%=aList9Values(0,i)%>";
				document.forma.elements["O<%=aList9Values(6,i)%>_Tarifa"].value = "<%=aList9Values(14,i)%>";
				document.forma.elements["O<%=aList9Values(6,i)%>_TarifaTipo"].value = "<%=aList9Values(16,i)%>";
				document.forma.elements["O<%=aList9Values(6,i)%>_Regimen"].value = "<%=aList9Values(15,i)%>";
				document.forma.elements["CO<%=aList9Values(6,i)%>"].value = "<%=aList9Values(2,i)%>";
				document.forma.elements["TCO<%=aList9Values(6,i)%>"].value = "<%=aList9Values(3,i)%>";
				document.forma.elements["VO<%=aList9Values(6,i)%>"].value = "<%=aList9Values(4,i)%>";
				document.forma.elements["NO<%=aList9Values(6,i)%>"].value = "<%=aList9Values(5,i)%>";
				document.forma.elements["SVIO<%=aList9Values(6,i)%>"].value = "<%=aList9Values(7,i)%>";
				document.forma.elements["SVNO<%=aList9Values(6,i)%>"].value = "<%=aList9Values(8,i)%>";
				document.forma.elements["TPO<%=aList9Values(6,i)%>"].value = "<%=aList9Values(9,i)%>";
				document.forma.elements["INVO<%=aList9Values(6,i)%>"].value = "<%=aList9Values(10,i)%>";
                document.forma.elements["CCBLO<%=aList9Values(6,i)%>"].value = "<%=aList9Values(11,i)%>";

                if (document.forma.elements["TCO<%=aList9Values(6,i)%>_copy"])
                    document.forma.elements["TCO<%=aList9Values(6,i)%>_copy"].value = '<%=IntLoc(CheckNum(aList9Values(3,i)))%>';
                if (document.forma.elements["TPO<%=aList9Values(6,i)%>_copy"])
                    document.forma.elements["TPO<%=aList9Values(6,i)%>_copy"].value = '<%=PrepColl(CheckNum(aList9Values(9,i)))%>';

				<% if aList9Values(10,i) <> 0 then%>
					document.forma.elements["O<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["CO<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["TCO<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["VO<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["NO<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["SVIO<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["SVNO<%=aList9Values(6,i)%>"].readOnly = 'true';
					document.forma.elements["TPO<%=aList9Values(6,i)%>"].disabled = 'false';
					document.forma.elements["INVO<%=aList9Values(6,i)%>"].disabled = 'false';
                    document.forma.elements["CCBLO<%=aList9Values(6,i)%>"].disabled = 'false';

					document.forma.elements["OtherChargeName"+<%=aList9Values(6,i)%>].readOnly = 'true';
					document.forma.elements["OtherChargeVal"+<%=aList9Values(6,i)%>].readOnly = 'true';
					document.getElementById("DEO<%=aList9Values(6,i)%>").style.visibility = "hidden";
					document.getElementById("DRO<%=aList9Values(6,i)%>").style.visibility = "hidden";
				<%end if%>
                <%l=l+1
			end select
		end if
	next
else%>
	<%if Request.Form("CAF") <> "" then%>
	document.forma.CAF.value = "<%=Request.Form("CAF")%>";
	<%end if%>
	<%if Request.Form("TCAF") <> "" then%>
	document.forma.TCAF.value = "<%=Request.Form("TCAF")%>";
	<%end if%>
	<%if Request.Form("TPAF") <> "" then%>
	document.forma.TPAF.value = "<%=Request.Form("TPAF")%>";
	<%end if%>
	<%if Request.Form("INVAF") <> "" then%>
	document.forma.INVAF.value = "<%=Request.Form("INVAF")%>";
	<%end if%>

	<%if Request.Form("CFS") <> "" then%>
	document.forma.CFS.value = "<%=Request.Form("CFS")%>";
	<%end if%>
	<%if Request.Form("TCFS") <> "" then%>
	document.forma.TCFS.value = "<%=Request.Form("TCFS")%>";
	<%end if%>
	<%if Request.Form("TPFS") <> "" then%>
	document.forma.TPFS.value = "<%=Request.Form("TPFS")%>";
	<%end if%>
	<%if Request.Form("INVFS") <> "" then%>
	document.forma.INVFS.value = "<%=Request.Form("INVFS")%>";
	<%end if%>

	<%if Request.Form("CSF") <> "" then%>
	document.forma.CSF.value = "<%=Request.Form("CSF")%>";
	<%end if%>
	<%if Request.Form("TCSF") <> "" then%>
	document.forma.TCSF.value = "<%=Request.Form("TCSF")%>";
	<%end if%>
	<%if Request.Form("TPSF") <> "" then%>
	document.forma.TPSF.value = "<%=Request.Form("TPSF")%>";
	<%end if%>
	<%if Request.Form("INVSF") <> "" then%>
	document.forma.INVSF.value = "<%=Request.Form("INVSF")%>";
	<%end if%>

	<%if Request.Form("CCF") <> "" then%>
	document.forma.CCF.value = "<%=Request.Form("CCF")%>";
	<%end if%>
	<%if Request.Form("TCCF") <> "" then%>
	document.forma.TCCF.value = "<%=Request.Form("TCCF")%>";
	<%end if%>
	<%if Request.Form("TPCF") <> "" then%>
	document.forma.TPCF.value = "<%=Request.Form("TPCF")%>";
	<%end if%>
	<%if Request.Form("INVCF") <> "" then%>
	document.forma.INVCF.value = "<%=Request.Form("INVCF")%>";
	<%end if%>

	<%if Request.Form("CTF") <> "" then%>
	document.forma.CTF.value = "<%=Request.Form("CTF")%>";
	<%end if%>
	<%if Request.Form("TCTF") <> "" then%>
	document.forma.TCTF.value = "<%=Request.Form("TCTF")%>";
	<%end if%>
	<%if Request.Form("TPTF") <> "" then%>
	document.forma.TPTF.value = "<%=Request.Form("TPTF")%>";
	<%end if%>
	<%if Request.Form("INVTF") <> "" then%>
	document.forma.INVTF.value = "<%=Request.Form("INVTF")%>";
	<%end if%>

	<%if Request.Form("CPU") <> "" then%>
	document.forma.CPU.value = "<%=Request.Form("CPU")%>";
	<%end if%>
	<%if Request.Form("TCPU") <> "" then%>
	document.forma.TCPU.value = "<%=Request.Form("TCPU")%>";
	<%end if%>
	<%if Request.Form("TPPU") <> "" then%>
	document.forma.TPPU.value = "<%=Request.Form("TPPU")%>";
	<%end if%>
	<%if Request.Form("INVPU") <> "" then%>
	document.forma.INVPU.value = "<%=Request.Form("INVPU")%>";
	<%end if%>

	<%if Request.Form("CFF") <> "" then%>
	document.forma.CFF.value = "<%=Request.Form("CFF")%>";
	<%end if%>
	<%if Request.Form("TCFF") <> "" then%>
	document.forma.TCFF.value = "<%=Request.Form("TCFF")%>";
	<%end if%>
	<%if Request.Form("TPFF") <> "" then%>
	document.forma.TPFF.value = "<%=Request.Form("TPFF")%>";
	<%end if%>
	<%if Request.Form("INVFF") <> "" then%>
	document.forma.INVFF.value = "<%=Request.Form("INVFF")%>";
	<%end if%>

	<%if Request.Form("CIM") <> "" then%>
	document.forma.CIM.value = "<%=Request.Form("CIM")%>";
	<%end if%>
	<%if Request.Form("TCIM") <> "" then%>
	document.forma.TCIM.value = "<%=Request.Form("TCIM")%>";
	<%end if%>
	<%if Request.Form("TPIM") <> "" then%>
	document.forma.TPIM.value = "<%=Request.Form("TPIM")%>";
	<%end if%>
	<%if Request.Form("INVIM") <> "" then%>
	document.forma.INVIM.value = "<%=Request.Form("INVIM")%>";
	<%end if%>

	<%if Request.Form("CPB") <> "" then%>
	document.forma.CPB.value = "<%=Request.Form("CPB")%>";
	<%end if%>
	<%if Request.Form("TCPB") <> "" then%>
	document.forma.TCPB.value = "<%=Request.Form("TCPB")%>";
	<%end if%>
	<%if Request.Form("TPPB") <> "" then%>
	document.forma.TPPB.value = "<%=Request.Form("TPPB")%>";
	<%end if%>
	<%if Request.Form("INVPB") <> "" then%>
	document.forma.INVPB.value = "<%=Request.Form("INVPB")%>";
	<%end if%>

	<%if Request.Form("CTX") <> "" then%>
	document.forma.CTX.value = "<%=Request.Form("CTX")%>";
	<%end if%>
	<%if Request.Form("TCTX") <> "" then%>
	document.forma.TCTX.value = "<%=Request.Form("TCTX")%>";
	<%end if%>
	<%if Request.Form("TPTX") <> "" then%>
	document.forma.TPTX.value = "<%=Request.Form("TPTX")%>";
	<%end if%>
	<%if Request.Form("INVTX") <> "" then%>
	document.forma.INVTX.value = "<%=Request.Form("INVTX")%>";
	<%end if%>

	<%for i=0 to 3%>
		document.forma.elements["C<%=i+1%>"].value = "<%=Request.Form("C"&(i+1))%>";

		<%if Request.Form("CC"&(i+1)) <> "" then%>
		document.forma.elements["CC<%=i+1%>"].value = "<%=Request.Form("CC"&(i+1))%>";
		<%end if%>
		<%if Request.Form("TCC"&(i+1)) <> "" then%>
		document.forma.elements["TCC<%=i+1%>"].value = "<%=Request.Form("TCC"&(i+1))%>";
		<%end if%>
		<%if Request.Form("TPC"&(i+1)) <> "" then%>
		document.forma.elements["TPC<%=i+1%>"].value = "<%=Request.Form("TPC"&(i+1))%>";
		<%end if%>
		
		document.forma.elements["VC<%=i+1%>"].value = "<%=Request.Form("VC"&(i+1))%>";
		document.forma.elements["NC<%=i+1%>"].value = "<%=Request.Form("NC"&(i+1))%>";
		document.forma.elements["SVNC<%=i+1%>"].value = "<%=Request.Form("SVNC"&(i+1))%>";
		document.forma.elements["SVIC<%=i+1%>"].value = "<%=Request.Form("SVIC"&(i+1))%>";
		document.forma.elements["INVC<%=i+1%>"].value = "<%=Request.Form("INVC"&(i+1))%>";
        document.forma.elements["CCBLC<%=i+1%>"].value = "<%=Request.Form("CCBLC"&(i+1))%>";
	<%next
    for i=0 to 5%>
		document.forma.elements["O<%=i+1%>"].value = "<%=Request.Form("O"&(i+1))%>";
		
		<%if Request.Form("CO"&(i+1)) <> "" then%>
		document.forma.elements["CO<%=i+1%>"].value = "<%=Request.Form("CO"&(i+1))%>";
		<%end if%>		
		<%if Request.Form("TCO"&(i+1)) <> "" then%>
		document.forma.elements["TCO<%=i+1%>"].value = "<%=Request.Form("TCO"&(i+1))%>";
		<%end if%>
		<%if Request.Form("TPO"&(i+1)) <> "" then%>
		document.forma.elements["TPO<%=i+1%>"].value = "<%=Request.Form("TPO"&(i+1))%>";
		<%end if%>
		
		document.forma.elements["VO<%=i+1%>"].value = "<%=Request.Form("VO"&(i+1))%>";
		document.forma.elements["NO<%=i+1%>"].value = "<%=Request.Form("NO"&(i+1))%>";
		document.forma.elements["SVNO<%=i+1%>"].value = "<%=Request.Form("SVNO"&(i+1))%>";
		document.forma.elements["SVIO<%=i+1%>"].value = "<%=Request.Form("SVIO"&(i+1))%>";
		document.forma.elements["INVO<%=i+1%>"].value = "<%=Request.Form("INVO"&(i+1))%>";
        document.forma.elements["CCBLO<%=i+1%>"].value = "<%=Request.Form("CCBLO"&(i+1))%>";
	<%next
	  for i=0 to 10%>
		document.forma.elements["A<%=i+1%>"].value = "<%=Request.Form("A"&(i+1))%>";

		<%if Request.Form("CA"&(i+1)) <> "" then%>
		document.forma.elements["CA<%=i+1%>"].value = "<%=Request.Form("CA"&(i+1))%>";
		<%end if%>
		<%if Request.Form("TCA"&(i+1)) <> "" then%>
		document.forma.elements["TCA<%=i+1%>"].value = "<%=Request.Form("TCA"&(i+1))%>";
		<%end if%>
		<%if Request.Form("TPA"&(i+1)) <> "" then%>
		document.forma.elements["TPA<%=i+1%>"].value = "<%=Request.Form("TPA"&(i+1))%>";
		<%end if%>

		document.forma.elements["VA<%=i+1%>"].value = "<%=Request.Form("VA"&(i+1))%>";
		document.forma.elements["NA<%=i+1%>"].value = "<%=Request.Form("NA"&(i+1))%>";
		document.forma.elements["SVNA<%=i+1%>"].value = "<%=Request.Form("SVNA"&(i+1))%>";
		document.forma.elements["SVIA<%=i+1%>"].value = "<%=Request.Form("SVIA"&(i+1))%>";
		document.forma.elements["INVA<%=i+1%>"].value = "<%=Request.Form("INVA"&(i+1))%>";
        document.forma.elements["CCBLA<%=i+1%>"].value = "<%=Request.Form("CCBLA"&(i+1))%>";
	<%next%>
<%end if%>
	selecciona('forma.CarrierID','<%=CarrierID%>');
	selecciona('forma.AirportDepID','<%=AirportDepID%>');
	selecciona('forma.AirportDesID','<%=AirportDesID%>');
	selecciona('forma.ChargeType','<%=ChargeType%>');
	selecciona('forma.ValChargeType','<%=ValChargeType%>');
	selecciona('forma.OtherChargeType','<%=OtherChargeType%>');
	selecciona('forma.CurrencyID','<%=CurrencyID%>');
	selecciona('forma.CalcAdminFee','<%=CalcAdminFee%>');
	selecciona('forma.SalespersonID','<%=SalespersonID%>');
	selecciona('forma.CTX','<%=CTX%>');
	selecciona('forma.TCTX','<%=TCTX%>');
	selecciona('forma.TPTX','<%=TPTX%>');
    selecciona('forma.replica','<%=replica%>');
	selecciona('forma.TipoCarga2','<%=TipoCarga2%>');

<% if CheckNum(Request("CarrierID2")) <> 0 then %>
    onChangeTransportista();
<% end if %>

<%if CarrierID <> 0 then%>
    FlgTF = SetFlag(document.forma.TerminalFee);
	FlgCF = SetFlag(document.forma.CustomFee);
	FlgFS = SetFlag(document.forma.FuelSurcharge);
	FlgSF = SetFlag(document.forma.SecurityFee);

	<%if Request("RAirportDepID") <> "" and Request("RAirportDesID") <> "" then%>
		selecciona('forma.AirportDepID','<%=Request("RAirportDepID")%>');
		selecciona('forma.AirportDesID','<%=Request("RAirportDesID")%>');
	<%end if%>	
	CheckRates(document.forms[0]);
<%end if
	Set aList1Values = Nothing
	Set aList2Values = Nothing
	Set aList3Values = Nothing
	Set aList4Values = Nothing
	Set aList5Values = Nothing
	Set aList6Values = Nothing
	Set aList7Values = Nothing
	Set aList8Values = Nothing
	Set aList9Values = Nothing

%>
var tio = 0;
tio = '<%=IIf(CheckNum(No) = 1 and IFNULL(ConsignerData) = "" and replica = "Master-Master-Hija",1,0)%>';
//console.log('********************** (' + tio + ')');
if (tio == 1) {
    RecalculoPeso(document.forms[0]);
}
</script>

<%
'response.write "(" & CheckNum(No) & ")(" & IFNULL(ConsignerData) & ")(" & IIf(CheckNum(No) = 1 and IFNULL(ConsignerData) = "",1,0) & ")<br>" 
%>

</body>
</html>