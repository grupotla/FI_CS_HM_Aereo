<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim Conn, rs, Action, AWBID, aTableValues, aTableValues2, QuerySelect, QuerySelect2, CreatedDate, CreatedTime, i, ntr
Dim CarrierID, CustomFee, FuelSurcharge, SecurityFee, CountTableValues, CountTableValues2, PickUp, Intermodal, SedFilingFee
Dim CountPositionValues, aPositionValues, Val, IATANo, Expired, Name, CarrierCode
Dim AWBNumber, AccountShipperNo, ShipperData, AccountConsignerNo, AirportCode
Dim ConsignerData, AgentData, AccountInformation, AccountAgentNo, AirportDepID, RequestedRouting
Dim AirportToCode1, AirportToCode2, AirportToCode3, CarrierCode2, CarrierCode3, CurrencyID
Dim ChargeType, ValChargeType, OtherChargeType, DeclaredValue, AduanaValue, AirportDesID, FlightDate1
Dim FlightDate2, SecuredValue, HandlingInformation, Observations, NoOfPieces, Weights, WeightsSymbol, Instructions
Dim Commodities, ChargeableWeights, CarrierRates, CarrierSubTot, NatureQtyGoods, TotNoOfPieces, TotWeight
Dim TotCarrierRate, TotChargeWeightPrepaid, TotChargeWeightCollect, TotChargeValuePrepaid, TotChargeValueCollect
Dim TotChargeTaxPrepaid, TotChargeTaxCollect, AnotherChargesAgentPrepaid, AnotherChargesAgentCollect
Dim AnotherChargesCarrierPrepaid, AnotherChargesCarrierCollect, TotPrepaid, TotCollect, TerminalFee
Dim PBA, TAX, AdditionalChargeName1, AdditionalChargeVal1, AdditionalChargeName2, CarrierName
Dim AdditionalChargeVal2, Invoice, ExportLic, AgentContactSignature, AWBDate, AgentSignature, CommoditiesTypes, TotWeightChargeable
Dim AdditionalChargeName3, AdditionalChargeVal3, AdditionalChargeName4, AdditionalChargeVal4, HAWBNumber
Dim AdditionalChargeName5, AdditionalChargeVal5, AdditionalChargeName6, AdditionalChargeVal6, AwbType
Dim DisplayNumber, AdditionalChargeName7, AdditionalChargeVal7, AdditionalChargeName8, AdditionalChargeVal8, AWBNo
Dim OtherChargeName1, OtherChargeName2, OtherChargeName3, OtherChargeName4, OtherChargeName5, OtherChargeName6
Dim OtherChargeVal1, OtherChargeVal2, OtherChargeVal3, OtherChargeVal4, OtherChargeVal5, OtherChargeVal6
Dim AdditionalChargeName9, AdditionalChargeVal9, AdditionalChargeName10, AdditionalChargeVal10, TDs, OtherChargesPrintType
Dim AdditionalChargeName11, AdditionalChargeVal11, AdditionalChargeName12, AdditionalChargeVal12, AdditionalChargeName13, AdditionalChargeVal13
Dim AdditionalChargeName14, AdditionalChargeVal14, AdditionalChargeName15, AdditionalChargeVal15, Countries
Dim ChargesDisplayLimit, ChargesDisplayLimit2, ChargesDisplayLimit3
Dim LimitHorizontal1, LimitHorizontal2, LimitHorizontal3, id_cliente_order, id_cliente_orderData, iMinimo



Dim lineData, fso, fs
Set fso = Server.CreateObject("Scripting.FileSystemObject") 
set fs = fso.OpenTextFile(Server.MapPath("air-waybill-2.htm"), 1, true) 
Do Until fs.AtEndOfStream 
    lineData = lineData & fs.ReadLine
    'do some parsing on lineData to get image data
    'output parsed data to screen
    'Response.Write lineData
Loop 

fs.close 

    LimitHorizontal1 = 4
    LimitHorizontal2 = 9
    LimitHorizontal3 = 14
    ChargesDisplayLimit=-1
    ChargesDisplayLimit2=-1
    ChargesDisplayLimit3=-1
	AWBID = CheckNum(Request("AWBID"))
	CountTableValues = -1
	CountPositionValues = -1
	Action = CheckNum(Request("Action"))
	CarrierID = CheckNum(Request("CID"))
	AwbType = CheckNum(Request("AT"))

	select case Action
	case 4
		QuerySelect = "select a.AWBID, a.CreatedDate, a.CreatedTime, a.Expired, " & _
					  "a.AWBNumber, a.AccountShipperNo, a.ShipperData, a.AccountConsignerNo, " & _
					  "a.ConsignerData, a.AgentData, a.AccountInformation, a.IATANo, a.AccountAgentNo, c.Name, a.RequestedRouting, " & _
					  "a.AirportToCode1, b.Name, a.AirportToCode2, a.AirportToCode3, a.CarrierCode2, a.CarrierCode3, e.CurrencyCode, " & _
					  "a.ChargeType, a.ValChargeType, a.OtherChargeType, a.DeclaredValue, a.AduanaValue, d.Name, a.FlightDate1, " & _
					  "a.FlightDate2, a.SecuredValue, a.HandlingInformation, a.Observations, a.NoOfPieces, a.Weights, a.WeightsSymbol, " & _
					  "a.Commodities, a.ChargeableWeights, a.CarrierRates, a.CarrierSubTot, a.NatureQtyGoods, a.TotNoOfPieces, a.TotWeight, " & _
					  "a.TotCarrierRate, a.TotChargeWeightPrepaid, a.TotChargeWeightCollect, a.TotChargeValuePrepaid, a.TotChargeValueCollect, " & _
					  "a.TotChargeTaxPrepaid, a.TotChargeTaxCollect, a.AnotherChargesAgentPrepaid, a.AnotherChargesAgentCollect, " & _
					  "a.AnotherChargesCarrierPrepaid, a.AnotherChargesCarrierCollect, a.TotPrepaid, a.TotCollect, a.TerminalFee, a.CustomFee, " & _
					  "a.FuelSurcharge, a.SecurityFee, a.PBA, a.TAX, a.AdditionalChargeName1, a.AdditionalChargeVal1, a.AdditionalChargeName2, " & _
					  "a.AdditionalChargeVal2, a.Invoice, a.ExportLic, a.AgentContactSignature, " & _
					  "a.Instructions, a.AgentSignature, a.AWBDate, c.AirportCode, " & _
					  "a.AdditionalChargeName3, a.AdditionalChargeVal3, a.AdditionalChargeName4, a.AdditionalChargeVal4, " & _					  
					  "a.HAWBNumber, a.AdditionalChargeName5, a.AdditionalChargeVal5, a.AdditionalChargeName6, a.AdditionalChargeVal6, " & _
					  "a.DisplayNumber, a.AdditionalChargeName7, a.AdditionalChargeVal7, a.AdditionalChargeName8, a.AdditionalChargeVal8, " & _
					  "a.AdditionalChargeName9, a.AdditionalChargeVal9, a.AdditionalChargeName10, a.AdditionalChargeVal10, " & _
					  "a.AdditionalChargeName11, a.AdditionalChargeVal11, a.AdditionalChargeName12, a.AdditionalChargeVal12, a.AdditionalChargeName13, a.AdditionalChargeVal13, " & _
					  "a.AdditionalChargeName14, a.AdditionalChargeVal14, a.AdditionalChargeName15, a.AdditionalChargeVal15, PickUp, Intermodal, SedFilingFee, a.Countries "
		if AwbType = 1 then '1.Export
			QuerySelect = QuerySelect & ", id_cliente_order, id_cliente_orderData from Awb a, Carriers b, Airports c, Airports d, Currencies e " & _
					  "where a.CarrierID = b.CarrierID and a.AirportDepID=c.AirportID and a.AirportDesID=d.AirportID and a.CurrencyID=e.CurrencyID " & _
				      "and a.AWBID=" & AWBID
            'En el query se excluye Air Freight (RubID=11) porque vienen el a variable totCarrierRate
            QuerySelect2 = "select ItemName, Value from ChargeItems where Expired=0 and CalcInBL=1 and ItemID<>11 and DocTyp=0 and AWBID=" & AWBID

		else '2.Import
			QuerySelect = QuerySelect & ", a.OtherChargeName1, a.OtherChargeName2, a.OtherChargeName3, a.OtherChargeName4, a.OtherChargeName5, a.OtherChargeName6, " & _
				" a.OtherChargeVal1, a.OtherChargeVal2, a.OtherChargeVal3, a.OtherChargeVal4, a.OtherChargeVal5, a.OtherChargeVal6, id_cliente_order, id_cliente_orderData from " & _
					  "Awbi a, Carriers b, Airports c, Airports d, Currencies e " & _
					  "where a.CarrierID = b.CarrierID and a.AirportDepID=c.AirportID and a.AirportDesID=d.AirportID and a.CurrencyID=e.CurrencyID " & _
				      "and a.AWBID=" & AWBID
            'En el query se excluye Air Freight (RubID=11) porque vienen el a variable totCarrierRate
            QuerySelect2 = "select ItemName, Value from ChargeItems where Expired=0 and CalcInBL=1 and ItemID<>11 and DocTyp=1 and AWBID=" & AWBID
		end if
		
        iMinimo = ""

        OpenConn Conn
		Set rs = Conn.Execute(QuerySelect)
        'response.write QuerySelect & "<br>"
		If Not rs.EOF Then
    		aTableValues = rs.GetRows
    		CountTableValues = rs.RecordCount


            QuerySelect = "select tarifa_minimo from Awb_Columns where AwbId=" & aTableValues(0, 0) & " and DocTyp = '" & IIf(AWBType = 1,"0","1") & "'"
            'response.write QuerySelect
            Set rs = Conn.Execute(QuerySelect)
            if Not rs.EOF then
                iMinimo = rs(0)
            end if
            CloseOBJ rs

	    End If
    	closeOBJs rs, Conn

		if CountTableValues <> -1 then
			ntr = chr(13) & chr(10)
			CreatedDate = aTableValues(1, 0)
			CreatedTime = aTableValues(2, 0)
			Expired = aTableValues(3, 0)
			AWBNumber = aTableValues(4, 0)
			AccountShipperNo = aTableValues(5, 0)
			ShipperData = FRegExp(ntr, aTableValues(6, 0), "<br>", 4)
			AccountConsignerNo = aTableValues(7, 0)
			ConsignerData = FRegExp(ntr, aTableValues(8, 0), "<br>", 4)
			AgentData = FRegExp(ntr, aTableValues(9, 0), "<br>", 4)
			AccountInformation = FRegExp(ntr, aTableValues(10, 0), "<br>", 4)
			IATANo = aTableValues(11, 0)
			AccountAgentNo = aTableValues(12, 0)
			AirportDepID = aTableValues(13, 0)
			RequestedRouting = aTableValues(14, 0)
			AirportToCode1 = aTableValues(15, 0)
			CarrierName = aTableValues(16, 0)
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
			HandlingInformation = FRegExp(ntr, aTableValues(31, 0), "<br>", 4)
			Observations = FRegExp(ntr, aTableValues(32, 0), "<br>", 4)
			'NoOfPieces = FRegExp(ntr, aTableValues(33, 0), "<br>", 4)
			NoOfPieces = aTableValues(33, 0)
			'Weights = FRegExp(ntr, aTableValues(34, 0), "<br>", 4)
            Weights = aTableValues(34, 0)
			WeightsSymbol = FRegExp(ntr, aTableValues(35, 0), "<br>", 4)
			Commodities = FRegExp(ntr, aTableValues(36, 0), "<br>", 4)
			'ChargeableWeights = FRegExp(ntr, aTableValues(37, 0), "<br>", 4)
            ChargeableWeights = aTableValues(37, 0)
			'CarrierRates = FRegExp(ntr, aTableValues(38, 0), "<br>", 4)
            CarrierRates = aTableValues(38, 0)
			CarrierSubTot = FRegExp(ntr, aTableValues(39, 0), "<br>", 4)
            CarrierSubTot = aTableValues(39, 0)
			NatureQtyGoods = FRegExp(ntr, aTableValues(40, 0), "<br>", 4)
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
			TerminalFee = CheckNum(aTableValues(56, 0))
			CustomFee = CheckNum(aTableValues(57, 0))
			FuelSurcharge = CheckNum(aTableValues(58, 0))
			SecurityFee = CheckNum(aTableValues(59, 0))
			PBA = CheckNum(aTableValues(60, 0))
			TAX = CheckNum(aTableValues(61, 0))
			AdditionalChargeName1 = aTableValues(62, 0)
			AdditionalChargeVal1 = CheckNum(aTableValues(63, 0))
			AdditionalChargeName2 = aTableValues(64, 0)
			AdditionalChargeVal2 = CheckNum(aTableValues(65, 0))
			Invoice = aTableValues(66, 0)
			ExportLic = aTableValues(67, 0)
			AgentContactSignature = aTableValues(68, 0)
			Instructions = FRegExp(ntr, aTableValues(69, 0), "<br>", 4)
			AgentSignature = aTableValues(70, 0)
			AWBDate = aTableValues(71, 0)
			AirportCode = aTableValues(72, 0)
			AdditionalChargeName3 = aTableValues(73, 0)
			AdditionalChargeVal3 = CheckNum(aTableValues(74, 0))
			AdditionalChargeName4 = aTableValues(75, 0)
			AdditionalChargeVal4 = CheckNum(aTableValues(76, 0))
			HAWBNumber = aTableValues(77, 0)
			AdditionalChargeName5 = aTableValues(78, 0)
			AdditionalChargeVal5 = CheckNum(aTableValues(79, 0))
			AdditionalChargeName6 = aTableValues(80, 0)
			AdditionalChargeVal6 = CheckNum(aTableValues(81, 0))
			DisplayNumber = aTableValues(82, 0)
			AdditionalChargeName7 = aTableValues(83, 0)
			AdditionalChargeVal7 = CheckNum(aTableValues(84, 0))
			AdditionalChargeName8 = aTableValues(85, 0)
			AdditionalChargeVal8 = CheckNum(aTableValues(86, 0))
			AdditionalChargeName9 = aTableValues(87, 0)
			AdditionalChargeVal9 = CheckNum(aTableValues(88, 0))
			AdditionalChargeName10 = aTableValues(89, 0)
			AdditionalChargeVal10 = CheckNum(aTableValues(90, 0))
			AdditionalChargeName11 = aTableValues(91, 0)
			AdditionalChargeVal11 = CheckNum(aTableValues(92, 0))
			AdditionalChargeName12 = aTableValues(93, 0)
			AdditionalChargeVal12 = CheckNum(aTableValues(94, 0))
			AdditionalChargeName13 = aTableValues(95, 0)
			AdditionalChargeVal13 = CheckNum(aTableValues(96, 0))
			AdditionalChargeName14 = aTableValues(97, 0)
			AdditionalChargeVal14 = CheckNum(aTableValues(98, 0))
			AdditionalChargeName15 = aTableValues(99, 0)
			AdditionalChargeVal15 = CheckNum(aTableValues(100, 0))
			PickUp = CheckNum(aTableValues(101, 0))
			Intermodal = CheckNum(aTableValues(102, 0))
			SedFilingFee = CheckNum(aTableValues(103, 0))	
            
			if AwbType=1 then 'export 2016-10-11
                id_cliente_order = CheckNum(aTableValues(105, 0))	                
                id_cliente_orderData = FRegExp(ntr, aTableValues(106, 0), "<br>", 4)
            else
                id_cliente_order = CheckNum(aTableValues(117, 0))	
                id_cliente_orderData = FRegExp(ntr, aTableValues(118, 0), "<br>", 4)
            end if
                
			if ChargeType=1 and AwbType=1 and aTableValues(104, 0)="GT" then 'Si es Prepaid-Export en Guatemala los valores del Agente no se imprimen
				AdditionalChargeVal1 = 0
				AdditionalChargeVal2 = 0
				AdditionalChargeVal6 = 0
				AdditionalChargeVal7 = 0
				AdditionalChargeVal9 = 0
				AdditionalChargeVal10 = 0
				AdditionalChargeVal11 = 0
				AdditionalChargeVal12 = 0
				AdditionalChargeVal13 = 0
				AdditionalChargeVal14 = 0
				AdditionalChargeVal15 = 0
                'Para GT solo se traen los rubros PickUp, SedFilingFee, Intermodal  PBA
                'QuerySelect2 = QuerySelect2 & " and ItemID in (31,38,115,116)"
			end if
            
            'Obteniendo los Rubros a Desplegar en la Guia
            CountTableValues2=-1
            'response.write QuerySelect2 & "<br>"
            OpenConn Conn
		    Set rs = Conn.Execute(QuerySelect2)
		    If Not rs.EOF Then
    		    aTableValues2 = rs.GetRows
    		    CountTableValues2 = rs.RecordCount - 1
	        End If
    	    closeOBJs rs, Conn
            
            'Esto funcionaba cuando se mostraban los rubros de la tabla Awb/Awbi, pero se comento porque ahora muestra rubros de la tabla ChargeItems
            'if AwbType = 1 then
			'	OtherChargeVal1 = 0
			'	OtherChargeVal2 = 0
			'	OtherChargeVal3 = 0
			'	OtherChargeVal4 = 0
			'	OtherChargeVal5 = 0
			'	OtherChargeVal6 = 0
			'else
			'	OtherChargeName1 = aTableValues(105, 0)
			'	OtherChargeName2 = aTableValues(106, 0)
			'	OtherChargeName3 = aTableValues(107, 0)
			'	OtherChargeName4 = aTableValues(108, 0)
			'	OtherChargeName5 = aTableValues(109, 0)
			'	OtherChargeName6 = aTableValues(110, 0)
			'	OtherChargeVal1 = CheckNum(aTableValues(111, 0))
			'	OtherChargeVal2 = CheckNum(aTableValues(112, 0))
			'	OtherChargeVal3 = CheckNum(aTableValues(113, 0))
			'	OtherChargeVal4 = CheckNum(aTableValues(114, 0))
			'	OtherChargeVal5 = CheckNum(aTableValues(115, 0))
			'	OtherChargeVal6 = CheckNum(aTableValues(116, 0))
			'end if

			if DisplayNumber = 1 then
				AWBNo = AWBNumber
			end if
			if HAWBNumber <> "" then
				if AWBNo <> "" then
					AWBNo = "<br>" & AWBNo 'Si el AWBNo ya trae el AWBNumber
				end if
				AWBNo = HAWBNumber & AWBNo
			end if
		end if
	case 5
		HAWBNumber = "No.AWB"
		AccountShipperNo = "Cuenta Embarcador"
		ShipperData = "Datos Embarcador"
		AccountConsignerNo = "Cuenta Destinatario"
		ConsignerData = "Datos Destinatario"
		AgentData = "Datos Agente"
		AccountInformation = "Informacion Contable"
		IATANo = "Cod.IATA Agente"
		AccountAgentNo = "Cuenta Agente"
		AirportDepID = "Aerop.Salida"
		RequestedRouting = "Ruta Solicitada"
		AirportToCode1 = "Cod.<br>Arp.1"
		CarrierName = "Trns<br>1"
		AirportToCode2 = "Cod.<br>Arp.2"
		AirportToCode3 = "Cod.<br>Arp.3"
		CarrierCode2 = "Trns<br>2"
		CarrierCode3 = "Trns<br>3"
		CurrencyID = "Mnd."
		ChargeType = 3
		ValChargeType = 3
		OtherChargeType = 3
		DeclaredValue = "Val.Decl.Trans."
		AduanaValue = "Val.Decl.Aduana"
		AirportDesID = "Aerop.Destino"
		FlightDate1 = "Fecha V.1"
		FlightDate2 = "Fecha V.2"
		SecuredValue = "Valor Asegurado"
		HandlingInformation = "Informacion de Manejo"
		Observations = "Observaciones"
		NoOfPieces = "No.<br>Bultos"
		Weights = "Peso<br>Bruto"
		WeightsSymbol = "Smb."
		Commodities = "Cod.Commod."
		ChargeableWeights = "Peso Cobrar"
		CarrierRates = "Tarifa Cargo"
		CarrierSubTot = "Tot."
		NatureQtyGoods = "Naturaleza Mercancia"
		TotNoOfPieces = "No.<br>Bultos"
		TotWeight = "Peso<br>Bruto"
		TotCarrierRate = "Tarifas"
		TotChargeWeightPrepaid = "C.Peso PP"
		TotChargeWeightCollect = "C.Peso CC"
		TotChargeValuePrepaid = "C.Valor PP"
		TotChargeValueCollect = "C.Valor CC"
		TotChargeTaxPrepaid = "Imp.PP"
		TotChargeTaxCollect = "Imp.CC"
		AnotherChargesAgentPrepaid = "C.Agente PP"
		AnotherChargesAgentCollect = "C.Agente CC"
		AnotherChargesCarrierPrepaid = "C.Trans.PP"
		AnotherChargesCarrierCollect = "C.Trans.CC"
		TotPrepaid = "Total PP"
		TotCollect = "Total CC"
		TerminalFee = "Ter.Fee"
		CustomFee = "Cus.Fee"
		FuelSurcharge = "Fue.Sur"
		SecurityFee = "Sec.Fee"
		PBA = "PBA"
		TAX = "TAX"
		AdditionalChargeName1 = "NCAA1"
		AdditionalChargeVal1 = "VCAA1"
		AdditionalChargeName2 = "NCAA2"
		AdditionalChargeVal2 = "VCAA2"
		Invoice = "Factura"
		ExportLic = "Lic.Exportacion"
		AgentContactSignature = "Firma Contacto Agente"
		Instructions = "Instrucciones"
		AgentSignature = "Firma Agente"
		AWBDate = "Fecha AWB"
		AirportCode = "C.Arpt.Sal."
		AdditionalChargeName3 = "NCAC1"
		AdditionalChargeVal3 = "VCAC1"
		AdditionalChargeName4 = "NCAC2"
		AdditionalChargeVal4 = "VCAC2"
		AdditionalChargeName5 = "NCAC3"
		AdditionalChargeVal5 = "VCAC3"
		AdditionalChargeName6 = "NCAA3"
		AdditionalChargeVal6 = "VCAA3"
		AdditionalChargeName7 = "NCAA4"
		AdditionalChargeVal7 = "VCAA4"
		AdditionalChargeName8 = "NCAC4"
		AdditionalChargeVal8 = "VCAC4"
		AdditionalChargeName9 = "NCAA5"
		AdditionalChargeVal9 = "VCAA5"
		AdditionalChargeName10 = "NCAA6"
		AdditionalChargeVal10 = "VCAA6"
		AdditionalChargeName11 = "NCAA7"
		AdditionalChargeVal11 = "VCAA7"
		AdditionalChargeName12 = "NCAA8"
		AdditionalChargeVal12 = "VCAA8"
		AdditionalChargeName13 = "NCAA9"
		AdditionalChargeVal13 = "VCAA9"
		AdditionalChargeName14 = "NCAA10"
		AdditionalChargeVal14 = "VCAA10"
		AdditionalChargeName15 = "NCAA11"
		AdditionalChargeVal15 = "VCAA11"
		OtherChargeName1 = "OCN1"
		OtherChargeVal1 = "OCV1"
		OtherChargeName2 = "OCN2"
		OtherChargeVal2 = "OCV2"
		OtherChargeName3 = "OCN3"
		OtherChargeVal3 = "OCV3"
		OtherChargeName4 = "OCN4"
		OtherChargeVal4 = "OCV4"
		OtherChargeName5 = "OCN5"
		OtherChargeVal5 = "OCV5"
		OtherChargeName6 = "OCN6"
		OtherChargeVal6 = "OCV6"
	End Select
	
	QuerySelect ="select ver_AWBNumber, hor_AWBNumber, ver_AccountShipperNo, hor_AccountShipperNo, ver_ShipperData, hor_ShipperData, " & _
				  "ver_AccountConsignerNo, hor_AccountConsignerNo, ver_ConsignerData, hor_ConsignerData, ver_AgentData, hor_AgentData, " & _
				  "ver_AccountInformation, hor_AccountInformation, ver_IATANo, hor_IATANo, ver_AccountAgentNo, hor_AccountAgentNo, ver_AirportDepID, " & _
				  "hor_AirportDepID, ver_RequestedRouting, hor_RequestedRouting, ver_AirportToCode1, hor_AirportToCode1, ver_CarrierID, hor_CarrierID, " & _
				  "ver_AirportToCode2, hor_AirportToCode2, ver_AirportToCode3, hor_AirportToCode3, ver_CarrierCode2, hor_CarrierCode2, ver_CarrierCode3, " & _
				  "hor_CarrierCode3, ver_CurrencyID, hor_CurrencyID, ver_ChargeType, hor_ChargeType, ver_ValChargeType, hor_ValChargeType, ver_OtherChargeType, " & _
				  "hor_OtherChargeType, ver_DeclaredValue, hor_DeclaredValue, ver_AduanaValue, hor_AduanaValue, ver_AirportDesID, hor_AirportDesID, ver_FlightDate1, " & _
				  "hor_FlightDate1, ver_FlightDate2, hor_FlightDate2, ver_SecuredValue, hor_SecuredValue, ver_HandlingInformation, hor_HandlingInformation, " & _
				  "ver_Observations, hor_Observations, ver_NoOfPieces, hor_NoOfPieces, ver_Weights, hor_Weights, ver_WeightsSymbol, hor_WeightsSymbol, " & _
				  "ver_Commodities, hor_Commodities, ver_ChargeableWeights, hor_ChargeableWeights, ver_CarrierRates, hor_CarrierRates, ver_CarrierSubTot, " & _
				  "hor_CarrierSubTot, ver_NatureQtyGoods, hor_NatureQtyGoods, ver_TotNoOfPieces, hor_TotNoOfPieces, ver_TotWeight, hor_TotWeight, ver_TotCarrierRate, " & _
				  "hor_TotCarrierRate, ver_TotChargeWeightPrepaid, hor_TotChargeWeightPrepaid, ver_TotChargeWeightCollect, hor_TotChargeWeightCollect, " & _
				  "ver_TotChargeValuePrepaid, hor_TotChargeValuePrepaid, ver_TotChargeValueCollect, hor_TotChargeValueCollect, ver_TotChargeTaxPrepaid, " & _
				  "hor_TotChargeTaxPrepaid, ver_TotChargeTaxCollect, hor_TotChargeTaxCollect, ver_AnotherChargesAgentPrepaid, hor_AnotherChargesAgentPrepaid, " & _ 
				  "ver_AnotherChargesAgentCollect, hor_AnotherChargesAgentCollect, ver_AnotherChargesCarrierPrepaid, hor_AnotherChargesCarrierPrepaid, " & _
				  "ver_AnotherChargesCarrierCollect, hor_AnotherChargesCarrierCollect, ver_TotPrepaid, hor_TotPrepaid, ver_TotCollect, hor_TotCollect, " & _
				  "ver_TerminalFee, hor_TerminalFee, ver_CustomFee, hor_CustomFee, ver_FuelSurcharge, hor_FuelSurcharge, ver_SecurityFee, hor_SecurityFee, " & _
				  "ver_PBA, hor_PBA, ver_TAX, hor_TAX, ver_AdditionalChargeName1, hor_AdditionalChargeName1, ver_AdditionalChargeVal1, hor_AdditionalChargeVal1, " & _ 
				  "ver_AdditionalChargeName2, hor_AdditionalChargeName2, ver_AdditionalChargeVal2, hor_AdditionalChargeVal2, ver_Invoice, hor_Invoice, ver_ExportLic, " & _ 
				  "hor_ExportLic, ver_AgentContactSignature, hor_AgentContactSignature, ver_Instructions, hor_Instructions, ver_AgentSignature, hor_AgentSignature, " & _ 
				  "ver_AWBDate, hor_AWBDate, ver_AirportCode, hor_AirportCode, " & _ 
				  "ver_AdditionalChargeName3, hor_AdditionalChargeName3, ver_AdditionalChargeVal3, hor_AdditionalChargeVal3, " & _ 
				  "ver_AdditionalChargeName4, hor_AdditionalChargeName4, ver_AdditionalChargeVal4, hor_AdditionalChargeVal4, " & _
				  "ver_ChargeType2, hor_ChargeType2, ver_ValChargeType2, hor_ValChargeType2, ver_OtherChargeType2, hor_OtherChargeType2, OtherChargesPrintType, " & _
                  "ver_AWBNumber2, hor_AWBNumber2, ver_AWBNumber3, hor_AWBNumber3, ver_id_cliente_order, hor_id_cliente_order from Carriers where CarrierID=" & CarrierID
	
        OpenConn Conn
		Set rs = Conn.Execute(QuerySelect)
		If Not rs.EOF Then
    		aPositionValues = rs.GetRows
    		CountPositionValues = rs.RecordCount
	    End If
        closeOBJs rs, Conn

        'Dividiendo los grupos para imprimir los rubros segun la cantidad de rubros que si deben verse en la guia
        if CountTableValues2>=0 and CountPositionValues<>-1 then
            if aPositionValues(152,0) = 0 then '0=Vertical, 1=Horizontal
                'Si el despliegue de Rubros es Vertical los rubros se presentan en 2 grupos de 15 lineas
                if CountTableValues2>15 then
                    ChargesDisplayLimit = 15
                else 
                    ChargesDisplayLimit = CountTableValues2
                end if
            else
                if CountTableValues2>LimitHorizontal1 then
                    ChargesDisplayLimit = LimitHorizontal1
                else
                    ChargesDisplayLimit = CountTableValues2
                end if
                    
                if CountTableValues2>LimitHorizontal2 then
                    ChargesDisplayLimit2 = LimitHorizontal2
                else
                    ChargesDisplayLimit2 = CountTableValues2
                end if

                if CountTableValues2>LimitHorizontal3 then
                    ChargesDisplayLimit3 = LimitHorizontal3
                else
                    ChargesDisplayLimit3 = CountTableValues2
                end if
            end if
        end if     
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        dim variable, x1, x2, x3, x4, iCarrierRates
        variable = ""

lineData = Replace(lineData,"#*no*#",variable)
lineData = Replace(lineData,"#*dep_port_code*#",AirportCode)
lineData = Replace(lineData,"#*hwb*#",AWBNo)
lineData = Replace(lineData,"#*carrier_name*#",CarrierName)
lineData = Replace(lineData,"#*shipper_name*#",ShipperData)
lineData = Replace(lineData,"#*shipper_account*#",AccountShipperNo)
lineData = Replace(lineData,"#*shipper_address*#",variable)
lineData = Replace(lineData,"#*carrier_address*#",variable)
lineData = Replace(lineData,"#*consignee_name*#",ConsignerData)
lineData = Replace(lineData,"#*consignee_account*#",AccountConsignerNo)
lineData = Replace(lineData,"#*consignee_address*#",variable)
lineData = Replace(lineData,"#*agent_name*#",AgentData)
lineData = Replace(lineData,"#*accounting_info*#",AccountInformation)
lineData = Replace(lineData,"#*agent_address*#",variable)
lineData = Replace(lineData,"#*iata_code*#",IATANo)
lineData = Replace(lineData,"#*agent_account*#",AccountAgentNo)
lineData = Replace(lineData,"#*dep_port*#",AirportDepID & " " & RequestedRouting)
lineData = Replace(lineData,"#*des_port_code*#",AirportToCode1)
lineData = Replace(lineData,"#*route*#",CarrierName & " " & RequestedRouting                            )
lineData = Replace(lineData,"#*to1*#",AirportToCode1)
lineData = Replace(lineData,"#*by1*#",CarrierCode2)
lineData = Replace(lineData,"#*to2*#",AirportToCode2)
lineData = Replace(lineData,"#*by2*#",CarrierCode3)
lineData = Replace(lineData,"#*currency*#",CurrencyID)
lineData = Replace(lineData,"#*nvd*#",DeclaredValue)
lineData = Replace(lineData,"#*ncv*#",AduanaValue)


    Select Case ChargeType 'PP
	Case 1 'PP
        x1 = "X"
	Case 2 'CC
        x1 = "X"
	Case Else
        x1 = "X"
        x2 = "X"
	end select

	Select Case ValChargeType 'PP
	Case 1 'PP
        x1 = "X"
	Case 2 'CC
        x2 = "X"
	Case Else
        x1 = "X"
        x2 = "X"
	end select

	Select Case OtherChargeType
	Case 1 'PP
        x3 = "X"
	Case 2 'CC
        x4 = "X"
	Case Else 
        x3 = "X"
        x4 = "X"
	end select


lineData = Replace(lineData,"#*x1*#",x1)
lineData = Replace(lineData,"#*x2*#",x2)
lineData = Replace(lineData,"#*x3*#",x3)
lineData = Replace(lineData,"#*x4*#",x3)
lineData = Replace(lineData,"#*des_port*#",AirportDesID)
lineData = Replace(lineData,"#*date_fly1*#",FlightDate1)
lineData = Replace(lineData,"#*date_fly2*#",FlightDate2)
lineData = Replace(lineData,"#*insurance_amount*#",SecuredValue)
lineData = Replace(lineData,"#*handling_info*#",HandlingInformation)
lineData = Replace(lineData,"#*country*#",Observations)
lineData = Replace(lineData,"#*pieces*#", FormatNumber(IIF(IsNumeric(NoOfPieces) , NoOfPieces, 0),2) )
lineData = Replace(lineData,"#*weight*#", FormatNumber(IIF(IsNumeric(Weights) , Weights, 0),2) )
lineData = Replace(lineData,"#*kl*#", WeightsSymbol)
lineData = Replace(lineData,"#*cod_commodity*#",Commodities)
lineData = Replace(lineData,"#*weight_charg*#", FormatNumber(IIF(IsNumeric(ChargeableWeights) , ChargeableWeights, 0),2) )
lineData = Replace(lineData,"#*rate_Charg*#",  IIf(iMinimo <> "", iMinimo, CarrierRates))
lineData = Replace(lineData,"#*tot_comm*#", FormatNumber(IIF(IsNumeric(CarrierSubTot) , CarrierSubTot, 0),2) )
lineData = Replace(lineData,"#*commodities*#",NatureQtyGoods)
lineData = Replace(lineData,"#*tot_pieces*#", FormatNumber(IIF(IsNumeric(TotNoOfPieces) , TotNoOfPieces, 0),2) )
lineData = Replace(lineData,"#*tot_weight*#", FormatNumber(IIF(IsNumeric(TotWeight) , TotWeight, 0),2) )
lineData = Replace(lineData,"#*total*#", FormatNumber(IIF(IsNumeric(TotCarrierRate) , TotCarrierRate, 0),2) )
lineData = Replace(lineData,"#*pp*#", FormatNumber(IIF(IsNumeric(TotChargeWeightPrepaid) , TotChargeWeightPrepaid, 0),2) )
lineData = Replace(lineData,"#*pp_wc*#",variable)
lineData = Replace(lineData,"#*coll_wc*#",variable)
lineData = Replace(lineData,"#*coll*#", FormatNumber(IIF(IsNumeric(TotChargeWeightCollect) , TotChargeWeightCollect, 0),2) )

iCarrierRates = ""

if CarrierRates = "AS AGREED" then 

    iCarrierRates = iCarrierRates & "AS AGREED"

else
	for i=0 to CountTableValues2  
                
        iCarrierRates = iCarrierRates & "<div style='display:inline-block;width:32%;background-color:#eee;margin:1px'>" & _
"<div style='padding-left:3px;float:left;border:0px solid green;font-size:10px;width:92px;display:block;;vertical-alignment:top;'>" & aTableValues2(0,i) & "</div>" & _
"<div style='padding-right:3px;float:right;border:0px solid orange;font-size:10px;width:52px;display:block;vertical-alignment:bottom;' align=right>" & FormatNumber(aTableValues2(1,i),2) & "</div>" & _
"</div>"
'"<div style='padding:0px;float:right;border:1px solid orange;width:5px;display:block;vertical-alignment:top'></div>" & _

    next 


    'if aPositionValues(152,0) = 0 then '0=Vertical, 1=Horizontal

		'iCarrierRates = iCarrierRates & "<table cellpadding=0 cellspacing=0>"
		'	for i=0 to CountTableValues2 'ChargesDisplayLimit 
        '        iCarrierRates = iCarrierRates & "<tr><td style='font-size:10px;'>" & aTableValues2(0,i) & "&nbsp;</td><td align=right valign=bottom  style='font-size:10px;'>" & aTableValues2(1,i) & "</td></tr>"
        '    next 

        '    if TAX <> 0 then iCarrierRates = iCarrierRates & "<tr><td>TAX&nbsp;</td><td align=right valign=bottom>" & TAX & "</td></tr>" end if
		
        'iCarrierRates = iCarrierRates & "</table>"

		'iCarrierRates = iCarrierRates & "<table cellpadding=0 cellspacing=0>"
		'	for i=16 to CountTableValues2 
        '        iCarrierRates = iCarrierRates & "<tr><td>" & TableValues2(0,i) & "&nbsp;</td><td align=right valign=bottom>" & aTableValues2(1,i) & "</td></tr>"
        '    next         
        'iCarrierRates = iCarrierRates & "</table>"
end if

 'response.write aTableValues(104, 0)

                dim firma : firma = "AIMAR"

                if InStr(1,aTableValues(104, 0),"TLA",1) > 0 then
                    firma = "GRUPO TLA"
                end if

                 if InStr(1,aTableValues(104, 0),"LTF",1) > 0 then
                    firma = "LATIN FREIGHT"
                end if

                 if InStr(1,aTableValues(104, 0),"N1",1) > 0 then
                    firma = "GRH"
                end if

lineData = Replace(lineData,"#*rates*#",iCarrierRates)
lineData = Replace(lineData,"#*pp_val_charg*#", FormatNumber(IIF(IsNumeric(TotChargeValuePrepaid) , TotChargeValuePrepaid, 0),2) )
lineData = Replace(lineData,"#*coll_val_charg*#", FormatNumber(IIF(IsNumeric(TotChargeValueCollect) , TotChargeValueCollect, 0),2) )
lineData = Replace(lineData,"#*pp_tax*#", FormatNumber(IIF(IsNumeric(TotChargeTaxPrepaid) , TotChargeTaxPrepaid, 0),2) )
lineData = Replace(lineData,"#*coll_tax*#", FormatNumber(IIF(IsNumeric(TotChargeTaxCollect) , TotChargeTaxCollect, 0),2) )
lineData = Replace(lineData,"#*other_charg_agen1*#", FormatNumber(IIF(IsNumeric(AnotherChargesAgentPrepaid) , AnotherChargesAgentPrepaid, 0),2) )
lineData = Replace(lineData,"#*other_charg_agen2*#", FormatNumber(IIF(IsNumeric(AnotherChargesAgentCollect) , AnotherChargesAgentCollect, 0),2) )
lineData = Replace(lineData,"#*other_charg_carr1*#", FormatNumber(IIF(IsNumeric(AnotherChargesCarrierPrepaid) , AnotherChargesCarrierPrepaid, 0),2) )
lineData = Replace(lineData,"#*other_charg_carr2*#", FormatNumber(IIF(IsNumeric(AnotherChargesCarrierCollect) , AnotherChargesCarrierCollect, 0),2) )
lineData = Replace(lineData,"#*other_charg_carr3*#",variable)
lineData = Replace(lineData,"#*other_charg_carr4*#",variable)
lineData = Replace(lineData,"#*tla*#",AgentContactSignature)
lineData = Replace(lineData,"#*agente*#","")
lineData = Replace(lineData,"#*pp_tot*#", FormatNumber(IIF(IsNumeric(TotPrepaid) , TotPrepaid, 0),2) ) 
lineData = Replace(lineData,"#*coll_tot*#", FormatNumber(IIF(IsNumeric(TotCollect) , TotCollect, 0),2) ) 
lineData = Replace(lineData,"#*date*#",ConvertDate(AWBDate,5))
lineData = Replace(lineData,"#*place*#",ExportLic)
lineData = Replace(lineData,"#*signature*#",AgentSignature)
lineData = Replace(lineData,"#*cc_rate*#",Invoice)
lineData = Replace(lineData,"#*cc_charg_des*#",variable)
lineData = Replace(lineData,"#*charg_des*#",variable)
lineData = Replace(lineData,"#*tot_coll_charg*#",variable)
lineData = Replace(lineData,"#*sci*#",variable)
lineData = Replace(lineData,"#*awb*#",variable)
lineData = Replace(lineData,"#*firma*#",firma & " " & Left(ConvertDate(AWBDate,4),4))
lineData = Replace(lineData,"background:","backgroundX:")

lineData = Replace(lineData,"<link rel=File-List href=""air-waybill-2_archivos/filelist.xml"">","")

lineData = Replace(lineData,"'>***",";text-align: justify; text-justify: inter-word;'>")

lineData = Replace(lineData,">DL<","><span class='dleft'></span><")
lineData = Replace(lineData,">DR<","><span class='dright'></span><")

dim dleft : dleft = Mid(lineData, InStr(1,lineData,"#*dl1*#",1)+7, InStr(1,lineData,"#*dl2*#",1) - InStr(1,lineData,"#*dl1*#",1)-7)
dim dright : dright = Mid(lineData, InStr(1,lineData,"#*dr1*#",1)+7, InStr(1,lineData,"#*dr2*#",1) - InStr(1,lineData,"#*dr1*#",1)-7)

'lineData = Replace(lineData,"font-style:italic;","font-style:italic;" & dleft )
'lineData = Replace(lineData,"text-decoration:underline;","text-decoration:underline;" & dright )
'lineData = Replace(lineData,"<style id=""air-waybill-2_13432_Styles"">","<style id=""air-waybill-2_13432_Styles"">  .dleft {position:relative;left:-3;top:-0.51;} .dright {position:relative;left:-1;top:-0.5;}")

lineData = Replace(lineData,"<style id=""air-waybill-2_13432_Styles"">","<style id=""air-waybill-2_13432_Styles"">  .dleft {" & dleft & "} .dright {" & dright & "}")

lineData = Replace(lineData,"#*dl1*#" & dleft & "#*dl2*#","")
lineData = Replace(lineData,"#*dr1*#" & dright & "#*dr2*#","")


lineData = Replace(lineData,"<meta name=ProgId content=Excel.Sheet>","")
'lineData = Replace(lineData,"<meta name=Generator content=""Microsoft Excel 14"">","")
lineData = Replace(lineData,"Microsoft Excel 14","")


lineData = Replace(lineData,"<!--[if !excel]>&nbsp;&nbsp;<![endif]-->","")
lineData = Replace(lineData,"<!--La siguiente informaci?n se gener? mediante el Asistente para publicar como","")
lineData = Replace(lineData,"p?gina web de Microsoft Excel.-->","")
lineData = Replace(lineData,"<!--Si se vuelve a publicar el mismo elemento desde Excel, se reemplazar? toda","")
lineData = Replace(lineData,"la informaci?n comprendida entre las etiquetas DIV.-->","")
lineData = Replace(lineData,"<!----------------------------->","")
lineData = Replace(lineData,"<!--INICIO DE LOS RESULTADOS DEL ASISTENTE PARA PUBLICAR COMO P?GINA WEB DE","")
lineData = Replace(lineData,"EXCEL -->","")
'<!----------------------------->


'<!----------------------------->
lineData = Replace(lineData,"<!--FINAL DE LOS RESULTADOS DEL ASISTENTE PARA PUBLICAR COMO P?GINA WEB DE","")'
lineData = Replace(lineData,"EXCEL-->","")
'<!----------------------------->



'response.Clear
'Response.Buffer = False
'This is download
'Response.ContentType = "application/pdf"
'Set file name
'Response.AddHeader "Content-Disposition", "inline; filename=myfile.pdf"

        Response.Write lineData   
        Response.End
%>
<html>
<!--
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
-->
<style type="text/css">
<!--
body {
	margin: 0px;
}
.style6 {font-size: 10}
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: bold;
}
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif;}
.style11 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
}
.style2 {color: #FFFFFF}
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #000000;
}
.style12 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; color: #FFFFFF; }
.style13 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; color: #FFFFFF; }
-->
</style>
<body onLoad="JavaScript:self.focus();">
<%
if CountPositionValues <> -1 then
%>

<% if aPositionValues(153,0) <> 0 OR aPositionValues(154,0) <> 0 then %>
<DIV style="LEFT: <%=aPositionValues(153,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(154,0)%>px;" class="style11"><%=AWBNo%></DIV>
<% end if %>
<% if aPositionValues(155,0) <> 0 OR aPositionValues(156,0) <> 0 then %>
<DIV style="LEFT: <%=aPositionValues(155,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(156,0)%>px;" class="style11"><%=AWBNo%></DIV>
<% end if %>

<% if aPositionValues(157,0) <> 0 OR aPositionValues(158,0) <> 0 then %>
<DIV style="LEFT: <%=aPositionValues(157,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(158,0)%>px;" class="style10"><table width="330" border="0" class="style10"><tr><td><%=id_cliente_orderData%></td></tr></table></DIV>
<% end if %>


<DIV style="LEFT: <%=aPositionValues(0,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(1,0)%>px;" class="style11"><%=AWBNo%></DIV>
<DIV style="LEFT: <%=aPositionValues(2,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(3,0)%>px;" class="style10"><%=AccountShipperNo%></DIV>
<DIV style="LEFT: <%=aPositionValues(4,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(5,0)%>px;" class="style10"><table width="330" border="0" class="style10"><tr><td><%=ShipperData%></td></tr></table></DIV>
<DIV style="LEFT: <%=aPositionValues(6,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(7,0)%>px;" class="style10"><%=AccountConsignerNo%></DIV>
<DIV style="LEFT: <%=aPositionValues(8,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(9,0)%>px;" class="style10"><table width="330" border="0" class="style10"><tr><td><%=ConsignerData%></td></tr></table></DIV>
<DIV style="LEFT: <%=aPositionValues(10,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(11,0)%>px;" class="style10"><table width="330" border="0" class="style10"><tr><td><%=AgentData%></td></tr></table></DIV>
<DIV style="LEFT: <%=aPositionValues(12,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(13,0)%>px;" class="style10"><%=AccountInformation%></DIV>
<DIV style="LEFT: <%=aPositionValues(14,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(15,0)%>px;" class="style10"><%=IATANo%></DIV>
<DIV style="LEFT: <%=aPositionValues(16,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(17,0)%>px;" class="style10"><%=AccountAgentNo%></DIV>
<DIV style="LEFT: <%=aPositionValues(18,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(19,0)%>px;" class="style10"><%=AirportDepID%></DIV>
<DIV style="LEFT: <%=aPositionValues(20,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(21,0)%>px;" class="style10"><%=RequestedRouting%></DIV>
<DIV style="LEFT: <%=aPositionValues(22,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(23,0)%>px;" class="style10"><%=AirportToCode1%></DIV>
<DIV style="LEFT: <%=aPositionValues(24,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(25,0)%>px;" class="style10"><%=CarrierName%></DIV>
<DIV style="LEFT: <%=aPositionValues(26,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(27,0)%>px;" class="style10"><%=AirportToCode2%></DIV>
<DIV style="LEFT: <%=aPositionValues(28,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(29,0)%>px;" class="style10"><%=AirportToCode3%></DIV>
<DIV style="LEFT: <%=aPositionValues(30,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(31,0)%>px;" class="style10"><%=CarrierCode2%></DIV>
<DIV style="LEFT: <%=aPositionValues(32,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(33,0)%>px;" class="style10"><%=CarrierCode3%></DIV>
<DIV style="LEFT: <%=aPositionValues(34,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(35,0)%>px;" class="style10"><%=CurrencyID%></DIV>
	<%Select Case ChargeType 'PP
	Case 1 'PP%>
<DIV style="LEFT: <%=aPositionValues(36,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(37,0)%>px;" class="style10">PP</DIV>
	<%Case 2 'CC%>
<DIV style="LEFT: <%=aPositionValues(146,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(147,0)%>px;" class="style10">CC</DIV>
	<%Case Else%>
<DIV style="LEFT: <%=aPositionValues(36,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(37,0)%>px;" class="style10">cg<br>PP</DIV>
<DIV style="LEFT: <%=aPositionValues(146,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(147,0)%>px;" class="style10">cg<br>CC</DIV>
	<%end select%>
	<%Select Case ValChargeType 'PP
	Case 1 'PP%>
<DIV style="LEFT: <%=aPositionValues(38,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(39,0)%>px;" class="style10">PP</DIV>
	<%Case 2 'CC%>
<DIV style="LEFT: <%=aPositionValues(148,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(149,0)%>px;" class="style10">CC</DIV>
	<%Case Else%>
<DIV style="LEFT: <%=aPositionValues(38,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(39,0)%>px;" class="style10">vc<br>PP</DIV>
<DIV style="LEFT: <%=aPositionValues(148,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(149,0)%>px;" class="style10">vc<br>CC</DIV>
	<%end select%>
	<%Select Case OtherChargeType
	Case 1 'PP%>
<DIV style="LEFT: <%=aPositionValues(40,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(41,0)%>px;" class="style10">PP</DIV>
	<%Case 2 'CC %>
<DIV style="LEFT: <%=aPositionValues(150,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(151,0)%>px;" class="style10">CC</DIV>
	<%Case Else %>
<DIV style="LEFT: <%=aPositionValues(40,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(41,0)%>px;" class="style10">oc<br>PP</DIV>
<DIV style="LEFT: <%=aPositionValues(150,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(151,0)%>px;" class="style10">oc<br>CC</DIV>
	<%end select%>
<DIV style="LEFT: <%=aPositionValues(42,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(43,0)%>px;" class="style10"><%=DeclaredValue%></DIV>
<DIV style="LEFT: <%=aPositionValues(44,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(45,0)%>px;" class="style10"><%=AduanaValue%></DIV>
<DIV style="LEFT: <%=aPositionValues(46,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(47,0)%>px;" class="style10"><%=AirportDesID%></DIV>
<DIV style="LEFT: <%=aPositionValues(48,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(49,0)%>px;" class="style10"><%=FlightDate1%></DIV>
<DIV style="LEFT: <%=aPositionValues(50,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(51,0)%>px;" class="style10"><%=FlightDate2%></DIV>
<DIV style="LEFT: <%=aPositionValues(52,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(53,0)%>px;" class="style10"><%=SecuredValue%></DIV>
<DIV style="LEFT: <%=aPositionValues(54,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(55,0)%>px;" class="style10"><%=HandlingInformation%></DIV>
<DIV style="LEFT: <%=aPositionValues(56,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(57,0)%>px;" class="style10"><%=Observations%></DIV>
<DIV style="LEFT: <%=aPositionValues(58,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(59,0)%>px;" class="style10"><%=NoOfPieces%></DIV>
<DIV style="LEFT: <%=aPositionValues(60,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(61,0)%>px;" class="style10"><%=Weights%></DIV>
<DIV style="LEFT: <%=aPositionValues(62,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(63,0)%>px;" class="style10"><%=WeightsSymbol%></DIV>
<DIV style="LEFT: <%=aPositionValues(64,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(65,0)%>px;" class="style10"><%=Commodities%></DIV>
<DIV style="LEFT: <%=aPositionValues(66,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(67,0)%>px;" class="style10"><%=ChargeableWeights%></DIV>


<DIV style="LEFT: <%=aPositionValues(68,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(69,0)%>px;" class="style10"><% if iMinimo <> "" then response.write(iMinimo) else response.write(CarrierRates) end if %></DIV>

<DIV style="LEFT: <%=aPositionValues(70,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(71,0)%>px;" class="style10"><%=CarrierSubTot%></DIV>
<DIV style="LEFT: <%=aPositionValues(72,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(73,0)%>px;" class="style10"><%=NatureQtyGoods%></DIV>
<DIV style="LEFT: <%=aPositionValues(74,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(75,0)%>px;" class="style10"><%=TotNoOfPieces%></DIV>
<DIV style="LEFT: <%=aPositionValues(76,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(77,0)%>px;" class="style10"><%=TotWeight%></DIV>
<DIV style="LEFT: <%=aPositionValues(78,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(79,0)%>px;" class="style10"><%=TotCarrierRate%></DIV>
<DIV style="LEFT: <%=aPositionValues(80,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(81,0)%>px;" class="style10"><%=TotChargeWeightPrepaid%></DIV>
<DIV style="LEFT: <%=aPositionValues(82,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(83,0)%>px;" class="style10"><%=TotChargeWeightCollect%></DIV>
<DIV style="LEFT: <%=aPositionValues(84,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(85,0)%>px;" class="style10"><%=TotChargeValuePrepaid%></DIV>
<DIV style="LEFT: <%=aPositionValues(86,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(87,0)%>px;" class="style10"><%=TotChargeValueCollect%></DIV>
<DIV style="LEFT: <%=aPositionValues(88,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(89,0)%>px;" class="style10"><%=TotChargeTaxPrepaid%></DIV>
<DIV style="LEFT: <%=aPositionValues(90,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(91,0)%>px;" class="style10"><%=TotChargeTaxCollect%></DIV>
<DIV style="LEFT: <%=aPositionValues(92,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(93,0)%>px;" class="style10"><%=AnotherChargesAgentPrepaid%></DIV>
<DIV style="LEFT: <%=aPositionValues(94,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(95,0)%>px;" class="style10"><%=AnotherChargesAgentCollect%></DIV>
<DIV style="LEFT: <%=aPositionValues(96,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(97,0)%>px;" class="style10"><%=AnotherChargesCarrierPrepaid%></DIV>
<DIV style="LEFT: <%=aPositionValues(98,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(99,0)%>px;" class="style10"><%=AnotherChargesCarrierCollect%></DIV>
<DIV style="LEFT: <%=aPositionValues(100,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(101,0)%>px;" class="style10"><%=TotPrepaid%></DIV>
<DIV style="LEFT: <%=aPositionValues(102,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(103,0)%>px;" class="style10"><%=TotCollect%></DIV>
<%if CarrierRates = "AS AGREED" then%>
<DIV style="LEFT: <%=aPositionValues(104,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(105,0)%>px;" class="style10">AS AGREED</DIV>
<%else%>
<DIV style="LEFT: <%=aPositionValues(104,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(105,0)%>px;" class="style10">
	<%  'response.write "LIMIT" & ChargesDisplayLimit & "<br>"
        'response.write "TIPO" & aPositionValues(152,0) & "<br>"
        'response.write "REGISTROS" & CountTableValues2 & "<br>"    

    if aPositionValues(152,0) = 0 then '0=Vertical, 1=Horizontal%>
	<table class="style10" cellpadding="0" cellspacing="0">
	<tr>
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%for i=0 to ChargesDisplayLimit %>
                <tr><td><%=aTableValues2(0,i) %>&nbsp;</td><td align="right" valign="bottom"><%=aTableValues2(1,i)%></td></tr>
            <%next %>
            <%if TAX <> 0 then %><tr><td>TAX&nbsp;</td><td align="right" valign="bottom"><%=TAX%></td></tr><%end if%>
		</table>
	</td>
	<td>&nbsp;&nbsp;&nbsp;</td>	
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%for i=16 to CountTableValues2 %>
                <tr><td><%=aTableValues2(0,i) %>&nbsp;</td><td align="right" valign="bottom"><%=aTableValues2(1,i)%></td></tr>
            <%next %>        
        </table>
	</td>	
	</tr>
	</table>
	<%else
			if AwbType = 1 then
				OtherChargeName1 = "TF"
				OtherChargeName2 = "CF"
				OtherChargeName3 = "FSS"
				OtherChargeName4 = "SSC"
				OtherChargeName5 = "PBA"
				OtherChargeName6 = "TAX"
			end if
	%>
	<table class="style10" cellpadding="0" cellspacing="0">
	<tr>
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%for i=0 to ChargesDisplayLimit %>
                <tr><td><%=aTableValues2(0,i) %>&nbsp;</td><td align="right" valign="bottom"><%=aTableValues2(1,i)%></td></tr>
            <%next %>
            <%if CountTableValues2<=LimitHorizontal1 then%>
            <%if TAX <> 0 then %><tr><td>TAX&nbsp;</td><td align="right" valign="bottom"><%=TAX%></td></tr><%end if%>
            <%end if %>
		</table>	
	</td>
	<td>&nbsp;&nbsp;</td>
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%for i=5 to ChargesDisplayLimit2 %>
                <tr><td><%=aTableValues2(0,i) %>&nbsp;</td><td align="right" valign="bottom"><%=aTableValues2(1,i)%></td></tr>
            <%next %>
            <%if CountTableValues2>LimitHorizontal1 and CountTableValues2<=LimitHorizontal2 then%>
            <%if TAX <> 0 then %><tr><td>TAX&nbsp;</td><td align="right" valign="bottom"><%=TAX%></td></tr><%end if%>
            <%end if %>
		</table>	
	</td>
	<td>&nbsp;&nbsp;</td>
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%for i=10 to ChargesDisplayLimit3 %>
                <tr><td><%=aTableValues2(0,i) %>&nbsp;</td><td align="right" valign="bottom"><%=aTableValues2(1,i)%></td></tr>
            <%next %>
            <%if CountTableValues2>LimitHorizontal2 and CountTableValues2<=LimitHorizontal3 then%>
            <%if TAX <> 0 then %><tr><td>TAX&nbsp;</td><td align="right" valign="bottom"><%=TAX%></td></tr><%end if%>
            <%end if %>
		</table>	
	</td>
	<td>&nbsp;&nbsp;</td>
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%for i=15 to CountTableValues2 %>
                <tr><td><%=aTableValues2(0,i) %>&nbsp;</td><td align="right" valign="bottom"><%=aTableValues2(1,i)%></td></tr>
            <%next %>
            <%if CountTableValues2>LimitHorizontal3 then%>
            <%if TAX <> 0 then %><tr><td>TAX&nbsp;</td><td align="right" valign="bottom"><%=TAX%></td></tr><%end if%>
            <%end if %>
		</table>	
	</td>
	</tr>
	</table>	
	<%end if%>
</DIV>
<%end if%>
<DIV style="LEFT: <%=aPositionValues(124,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(125,0)%>px;" class="style10"><%=Invoice%></DIV>
<DIV style="LEFT: <%=aPositionValues(126,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(127,0)%>px;" class="style10"><%=ExportLic%></DIV>
<DIV style="LEFT: <%=aPositionValues(128,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(129,0)%>px;" class="style10"><%=AgentContactSignature%></DIV>
<DIV style="LEFT: <%=aPositionValues(130,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(131,0)%>px;" class="style10"><%=Instructions%></DIV>
<DIV style="LEFT: <%=aPositionValues(132,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(133,0)%>px;" class="style10"><%=AgentSignature%></DIV>
<DIV style="LEFT: <%=aPositionValues(134,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(135,0)%>px;" class="style10"><%=ConvertDate(AWBDate,5)%></DIV>
<DIV style="LEFT: <%=aPositionValues(136,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(137,0)%>px;" class="style10"><%=AirportCode%></DIV>
<%
Else
%>
<script>    alert("No Existe plantilla");</script>
<%
End if
%>
</body>
</html>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>