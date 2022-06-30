<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim Conn, rs, Action, AWBID, aTableValues, QuerySelect, CreatedDate, CreatedTime, i, ntr
Dim CarrierID, CustomFee, FuelSurcharge, SecurityFee, CountTableValues, TableName, PickUp, Intermodal, SedFilingFee
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
			TableName = "Awb"
			QuerySelect = QuerySelect & " from " & TableName & " a, Carriers b, Airports c, Airports d, Currencies e " & _
					  "where a.CarrierID = b.CarrierID and a.AirportDepID=c.AirportID and a.AirportDesID=d.AirportID and a.CurrencyID=e.CurrencyID " & _
				      "and a.AWBID=" & AWBID
		else '2.Import
			TableName = "Awbi"
			QuerySelect = QuerySelect & ", a.OtherChargeName1, a.OtherChargeName2, a.OtherChargeName3, a.OtherChargeName4, a.OtherChargeName5, a.OtherChargeName6, " & _
				" a.OtherChargeVal1, a.OtherChargeVal2, a.OtherChargeVal3, a.OtherChargeVal4, a.OtherChargeVal5, a.OtherChargeVal6 from " & _
					  TableName & " a, Carriers b, Airports c, Airports d, Currencies e " & _
					  "where a.CarrierID = b.CarrierID and a.AirportDepID=c.AirportID and a.AirportDesID=d.AirportID and a.CurrencyID=e.CurrencyID " & _
				      "and a.AWBID=" & AWBID
		end if
					  
		OpenConn Conn
        Set rs = Conn.Execute(QuerySelect)
		If Not rs.EOF Then
    		aTableValues = rs.GetRows
    		CountTableValues = rs.RecordCount
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
			NoOfPieces = FRegExp(ntr, aTableValues(33, 0), "<br>", 4)
			Weights = FRegExp(ntr, aTableValues(34, 0), "<br>", 4)
			WeightsSymbol = FRegExp(ntr, aTableValues(35, 0), "<br>", 4)
			Commodities = FRegExp(ntr, aTableValues(36, 0), "<br>", 4)
			ChargeableWeights = FRegExp(ntr, aTableValues(37, 0), "<br>", 4)
			CarrierRates = FRegExp(ntr, aTableValues(38, 0), "<br>", 4)
			CarrierSubTot = FRegExp(ntr, aTableValues(39, 0), "<br>", 4)
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
			end if
			
			if AwbType = 1 then
				OtherChargeVal1 = 0
				OtherChargeVal2 = 0
				OtherChargeVal3 = 0
				OtherChargeVal4 = 0
				OtherChargeVal5 = 0
				OtherChargeVal6 = 0
			else
				OtherChargeName1 = aTableValues(105, 0)
				OtherChargeName2 = aTableValues(106, 0)
				OtherChargeName3 = aTableValues(107, 0)
				OtherChargeName4 = aTableValues(108, 0)
				OtherChargeName5 = aTableValues(109, 0)
				OtherChargeName6 = aTableValues(110, 0)
				OtherChargeVal1 = CheckNum(aTableValues(111, 0))
				OtherChargeVal2 = CheckNum(aTableValues(112, 0))
				OtherChargeVal3 = CheckNum(aTableValues(113, 0))
				OtherChargeVal4 = CheckNum(aTableValues(114, 0))
				OtherChargeVal5 = CheckNum(aTableValues(115, 0))
				OtherChargeVal6 = CheckNum(aTableValues(116, 0))
			end if
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
				  "ver_ChargeType2, hor_ChargeType2, ver_ValChargeType2, hor_ValChargeType2, ver_OtherChargeType2, hor_OtherChargeType2, OtherChargesPrintType from Carriers where CarrierID=" & CarrierID
	
		OpenConn Conn
		Set rs = Conn.Execute(QuerySelect)
		If Not rs.EOF Then
    		aPositionValues = rs.GetRows
    		CountPositionValues = rs.RecordCount
	    End If
    	closeOBJs rs, Conn
%>
<html>
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
<DIV style="LEFT: <%=aPositionValues(68,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(69,0)%>px;" class="style10"><%=CarrierRates%></DIV>
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
	<% if aPositionValues(152,0) = 0 then%>
	<table class="style10" cellpadding="0" cellspacing="0">
	<tr>
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%if TerminalFee <> 0 then %><tr><td>TERMINAL FEE&nbsp;</td><td align="right"><%=TerminalFee%></td></tr><%end if%>
			<%if CustomFee <> 0 then %><tr><td>CUSTOM FEE&nbsp;</td><td align="right"><%=CustomFee%></td></tr><%end if%>
			<%if FuelSurcharge <> 0 then %><tr><td>FUEL SURCHARGE&nbsp;</td><td align="right"><%=FuelSurcharge%></td></tr><%end if%>
			<%if SecurityFee <> 0 then %><tr><td>SECURITY FEE&nbsp;</td><td align="right"><%=SecurityFee%></td></tr><%end if%>
			<%if PBA <> 0 then %><tr><td>PBA&nbsp;</td><td align="right"><%=PBA%></td></tr><%end if%>
			<%if TAX <> 0 then %><tr><td>TAX&nbsp;</td><td align="right"><%=TAX%></td></tr><%end if%>
			<%if OtherChargeVal1 <> 0 then %><tr><td><%=OtherChargeName1%>&nbsp;</td><td align="right"><%=OtherChargeVal1%></td></tr><%end if%>
			<%if OtherChargeVal2 <> 0 then %><tr><td><%=OtherChargeName2%>&nbsp;</td><td align="right"><%=OtherChargeVal2%></td></tr><%end if%>
			<%if OtherChargeVal3 <> 0 then %><tr><td><%=OtherChargeName3%>&nbsp;</td><td align="right"><%=OtherChargeVal3%></td></tr><%end if%>
			<%if OtherChargeVal4 <> 0 then %><tr><td><%=OtherChargeName4%>&nbsp;</td><td align="right"><%=OtherChargeVal4%></td></tr><%end if%>
			<%if OtherChargeVal5 <> 0 then %><tr><td><%=OtherChargeName5%>&nbsp;</td><td align="right"><%=OtherChargeVal5%></td></tr><%end if%>
			<%if OtherChargeVal6 <> 0 then %><tr><td><%=OtherChargeName6%>&nbsp;</td><td align="right"><%=OtherChargeVal6%></td></tr><%end if%>
			<%if AdditionalChargeVal3 <> 0 then %><tr><td><%=AdditionalChargeName3%>&nbsp;</td><td align="right"><%=AdditionalChargeVal3%></td></tr><%end if%>
			<%if AdditionalChargeVal4 <> 0 then %><tr><td><%=AdditionalChargeName4%>&nbsp;</td><td align="right"><%=AdditionalChargeVal4%></td></tr><%end if%>
			<%if AdditionalChargeVal5 <> 0 then %><tr><td><%=AdditionalChargeName5%>&nbsp;</td><td align="right"><%=AdditionalChargeVal5%></td></tr><%end if%>
			<%if AdditionalChargeVal8 <> 0 then %><tr><td><%=AdditionalChargeName8%>&nbsp;</td><td align="right"><%=AdditionalChargeVal8%></td></tr><%end if%>
		</table>
	</td>
	<td>&nbsp;&nbsp;&nbsp;</td>	
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%if PickUp <> 0 then %><tr><td>PICK-UP&nbsp;</td><td align="right"><%=PickUp%></td></tr><%end if%>
			<%if Intermodal <> 0 then %><tr><td>INTERMODAL&nbsp;</td><td align="right"><%=Intermodal%></td></tr><%end if%>
			<%if SedFilingFee <> 0 then %><tr><td>SED FILING FEE&nbsp;</td><td align="right"><%=SedFilingFee%></td></tr><%end if%>
			<%if AdditionalChargeVal1 <> 0 then %><tr><td><%=AdditionalChargeName1%>&nbsp;</td><td align="right"><%=AdditionalChargeVal1%></td></tr><%end if%>
			<%if AdditionalChargeVal2 <> 0 then %><tr><td><%=AdditionalChargeName2%>&nbsp;</td><td align="right"><%=AdditionalChargeVal2%></td></tr><%end if%>
			<%if AdditionalChargeVal6 <> 0 then %><tr><td><%=AdditionalChargeName6%>&nbsp;</td><td align="right"><%=AdditionalChargeVal6%></td></tr><%end if%>
			<%if AdditionalChargeVal7 <> 0 then %><tr><td><%=AdditionalChargeName7%>&nbsp;</td><td align="right"><%=AdditionalChargeVal7%></td></tr><%end if%>
			<%if AdditionalChargeVal9 <> 0 then %><tr><td><%=AdditionalChargeName9%>&nbsp;</td><td align="right"><%=AdditionalChargeVal9%></td></tr><%end if%>
			<%if AdditionalChargeVal10 <> 0 then %><tr><td><%=AdditionalChargeName10%>&nbsp;</td><td align="right"><%=AdditionalChargeVal10%></td></tr><%end if%>
			<%if AdditionalChargeVal11 <> 0 then %><tr><td><%=AdditionalChargeName11%>&nbsp;</td><td align="right"><%=AdditionalChargeVal11%></td></tr><%end if%>
			<%if AdditionalChargeVal12 <> 0 then %><tr><td><%=AdditionalChargeName12%>&nbsp;</td><td align="right"><%=AdditionalChargeVal12%></td></tr><%end if%>
			<%if AdditionalChargeVal13 <> 0 then %><tr><td><%=AdditionalChargeName13%>&nbsp;</td><td align="right"><%=AdditionalChargeVal13%></td></tr><%end if%>
			<%if AdditionalChargeVal14 <> 0 then %><tr><td><%=AdditionalChargeName14%>&nbsp;</td><td align="right"><%=AdditionalChargeVal14%></td></tr><%end if%>
			<%if AdditionalChargeVal15 <> 0 then %><tr><td><%=AdditionalChargeName15%>&nbsp;</td><td align="right"><%=AdditionalChargeVal15%></td></tr><%end if%>
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
			<%if TerminalFee <> 0 then %><tr><td>TERMINAL FEE&nbsp;</td><td align="right"><%=TerminalFee%></td></tr><%end if%>
			<%if CustomFee <> 0 then %><tr><td>CUSTOM FEE&nbsp;</td><td align="right"><%=CustomFee%></td></tr><%end if%>
			<%if FuelSurcharge <> 0 then %><tr><td>FUEL SURCHARGE&nbsp;</td><td align="right"><%=FuelSurcharge%></td></tr><%end if%>
			<%if SecurityFee <> 0 then %><tr><td>SECURITY FEE&nbsp;</td><td align="right"><%=SecurityFee%></td></tr><%end if%>
			<%if PBA <> 0 then %><tr><td>PBA&nbsp;</td><td align="right"><%=PBA%></td></tr><%end if%>
			<%if TAX <> 0 then %><tr><td>TAX&nbsp;</td><td align="right"><%=TAX%></td></tr><%end if%>
			<%if OtherChargeVal1 <> 0 then %><tr><td><%=OtherChargeName1%>&nbsp;</td><td align="right"><%=OtherChargeVal1%></td></tr><%end if%>
			<%if OtherChargeVal2 <> 0 then %><tr><td><%=OtherChargeName2%>&nbsp;</td><td align="right"><%=OtherChargeVal2%></td></tr><%end if%>
		</table>	
	</td>
	<td>&nbsp;&nbsp;</td>
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%if OtherChargeVal3 <> 0 then %><tr><td><%=OtherChargeName3%>&nbsp;</td><td align="right"><%=OtherChargeVal3%></td></tr><%end if%>
			<%if OtherChargeVal4 <> 0 then %><tr><td><%=OtherChargeName4%>&nbsp;</td><td align="right"><%=OtherChargeVal4%></td></tr><%end if%>
			<%if OtherChargeVal5 <> 0 then %><tr><td><%=OtherChargeName5%>&nbsp;</td><td align="right"><%=OtherChargeVal5%></td></tr><%end if%>
			<%if OtherChargeVal6 <> 0 then %><tr><td><%=OtherChargeName6%>&nbsp;</td><td align="right"><%=OtherChargeVal6%></td></tr><%end if%>
			<%if AdditionalChargeVal3 <> 0 then %><tr><td><%=AdditionalChargeName3%>&nbsp;</td><td align="right"><%=AdditionalChargeVal3%></td></tr><%end if%>
			<%if AdditionalChargeVal4 <> 0 then %><tr><td><%=AdditionalChargeName4%>&nbsp;</td><td align="right"><%=AdditionalChargeVal4%></td></tr><%end if%>
			<%if AdditionalChargeVal5 <> 0 then %><tr><td><%=AdditionalChargeName5%>&nbsp;</td><td align="right"><%=AdditionalChargeVal5%></td></tr><%end if%>
			<%if AdditionalChargeVal8 <> 0 then %><tr><td><%=AdditionalChargeName8%>&nbsp;</td><td align="right"><%=AdditionalChargeVal8%></td></tr><%end if%>
		</table>	
	</td>
	<td>&nbsp;&nbsp;</td>
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%if PickUp <> 0 then %><tr><td>PICK-UP&nbsp;</td><td align="right"><%=PickUp%></td></tr><%end if%>
			<%if Intermodal <> 0 then %><tr><td>INTERMODAL&nbsp;</td><td align="right"><%=Intermodal%></td></tr><%end if%>
			<%if SedFilingFee <> 0 then %><tr><td>SED FILING FEE&nbsp;</td><td align="right"><%=SedFilingFee%></td></tr><%end if%>
			<%if AdditionalChargeVal1 <> 0 then %><tr><td><%=AdditionalChargeName1%>&nbsp;</td><td align="right"><%=AdditionalChargeVal1%></td></tr><%end if%>
			<%if AdditionalChargeVal2 <> 0 then %><tr><td><%=AdditionalChargeName2%>&nbsp;</td><td align="right"><%=AdditionalChargeVal2%></td></tr><%end if%>
			<%if AdditionalChargeVal6 <> 0 then %><tr><td><%=AdditionalChargeName6%>&nbsp;</td><td align="right"><%=AdditionalChargeVal6%></td></tr><%end if%>
			<%if AdditionalChargeVal7 <> 0 then %><tr><td><%=AdditionalChargeName7%>&nbsp;</td><td align="right"><%=AdditionalChargeVal7%></td></tr><%end if%>
		</table>	
	</td>
	<td>&nbsp;&nbsp;</td>
	<td valign="top">
		<table class="style10" cellpadding="0" cellspacing="0">
			<%if AdditionalChargeVal9 <> 0 then %><tr><td><%=AdditionalChargeName9%>&nbsp;</td><td align="right"><%=AdditionalChargeVal9%></td></tr><%end if%>
			<%if AdditionalChargeVal10 <> 0 then %><tr><td><%=AdditionalChargeName10%>&nbsp;</td><td align="right"><%=AdditionalChargeVal10%></td></tr><%end if%>
			<%if AdditionalChargeVal11 <> 0 then %><tr><td><%=AdditionalChargeName11%>&nbsp;</td><td align="right"><%=AdditionalChargeVal11%></td></tr><%end if%>
			<%if AdditionalChargeVal12 <> 0 then %><tr><td><%=AdditionalChargeName12%>&nbsp;</td><td align="right"><%=AdditionalChargeVal12%></td></tr><%end if%>
			<%if AdditionalChargeVal13 <> 0 then %><tr><td><%=AdditionalChargeName13%>&nbsp;</td><td align="right"><%=AdditionalChargeVal13%></td></tr><%end if%>
			<%if AdditionalChargeVal14 <> 0 then %><tr><td><%=AdditionalChargeName14%>&nbsp;</td><td align="right"><%=AdditionalChargeVal14%></td></tr><%end if%>
			<%if AdditionalChargeVal15 <> 0 then %><tr><td><%=AdditionalChargeName15%>&nbsp;</td><td align="right"><%=AdditionalChargeVal15%></td></tr><%end if%>
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