<script LANGUAGE="VBScript" RUNAT="Server">
Const rsOpenStatic = 3
Const rsLockReadOnly = 1
Const rsClientSide = 3
Const spacer = "&nbsp;&nbsp;"

'PRODUCCION - PRUEBAS_
'Const Connection = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=DbAereo;PWD=aereoaimar;DATABASE=pruebas_db_aereo" 
'Const Connection2 = "Driver={PostgreSQL Unicode(x64)};SERVER=10.10.1.20;UID=dbmaster;PWD=aimargt;DATABASE=pruebas_master-aimar;Fetch=50000;"'PORT=5432
'Const ConnectionBAW = "Driver={PostgreSQL Unicode(x64)};SERVER=10.10.1.18;UID=dbmaster;PWD=aimargt;DATABASE=pruebas_aimar_baw;Fetch=50000;"'PORT=5432
'Const Connection3 = "Driver={PostgreSQL Unicode(x64)};SERVER=10.10.1.20;UID=dbmaster;PWD=aimargt;DATABASE=pruebas_ti_pricing;Fetch=50000;"'PORT=5432

'LOCALHOST
Const Connection = "Driver={MySQL ODBC 3.51 Driver};SERVER=10.10.1.18;UID=DbAereo;PWD=aereoaimar;DATABASE=pruebas_db_aereo" 
'Const Connection = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;UID=root;PWD=123456;DATABASE=pruebas_db_aereo" 
Const Connection2 = "Driver={PostgreSQL Unicode};SERVER=10.10.1.20;UID=dbmaster;PWD=aimargt;DATABASE=pruebas_master-aimar;Fetch=50000;"'PORT=5432
'Const Connection2 = "Driver={PostgreSQL Unicode};SERVER=192.168.56.1;UID=postgres;PWD=adminpass;DATABASE=master-aimar;Fetch=50000;"'PORT=5432
Const ConnectionBAW = "Driver={PostgreSQL Unicode};SERVER=10.10.1.18;UID=dbmaster;PWD=aimargt;DATABASE=pruebas_aimar_baw;Fetch=50000;"'PORT=5432
'Const ConnectionBAW = "Driver={PostgreSQL Unicode};SERVER=192.168.56.1;UID=postgres;PWD=adminpass;DATABASE=aimar_baw;Fetch=50000;"'PORT=5432
Const Connection3 = "Driver={PostgreSQL Unicode};SERVER=10.10.1.20;UID=dbmaster;PWD=aimargt;DATABASE=pruebas_ti_pricing2;Fetch=50000;"'PORT=5432
'Const Connection3 = "Driver={PostgreSQL Unicode};SERVER=192.168.56.1;UID=postgres;PWD=adminpass;DATABASE=ti_pricing;Fetch=50000;"'PORT=5432

Const PtrnCountries = "'GT'|'SV'|'HN'|'NI'|'CR'|'PA'|'BZ'|'N1'|'GT2'|'SV2'|'HN2'|'NI2'|'CR2'|'PA2'|'BZ2'|'GTLTF'|'SVLTF'|'HNLTF'|'NILTF'|'CRLTF'|'PALTF'|'CN'|'BR'|'GTRMR'|'BE'|'ES'"
Const PtrnEconoCodes = "100|1721"'Codigos de Agente Econocaribe
Const FilterAimarLatin = 0

'Constantes para Conectar al BAW para realizar Provisiones
Const soapServer = "10.10.1.7:8181"
Const soapPath = "/WebServices/BAW_Provisionar_BL.asmx"
Const soapPathIntercompany = "/WebServices/BAW_Intercompany_Operativo.asmx"

Const soapServerCombex = "200.35.177.74"
Const soapPathCombex = "/combex_ws/Combex_WS.asmx"
Const strWebMethodCombex = "Guia"
Const iIps = "[::1,127.0.0.1,localhost]"

Function GetHTMLSource(Page)
Dim http
    'Funcion para obtener codigo HTML de la pagina indicada en parametro Page
    Set http = Server.CreateObject("Msxml2.ServerXMLHTTP")
    http.Open "GET", Page, False
    http.setRequestHeader "Content-Type", "text/html; charset=utf-8"
    http.Send
    GetHTMLSource = http.ResponseText
    Set http = nothing
End Function

Dim iArr2 'para los perfiles
Dim iPerfilOpcion

Sub openConn(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (Connection)		
End Sub

Sub openConn2(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (Connection2)
End Sub

Sub openConn3(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (Connection3)
End Sub

Sub openConnBAW(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (ConnectionBAW)
End Sub

Sub closeOBJ(theOBJ)
    On Error Resume Next
    theOBJ.Close
    Set theOBJ = Nothing
End Sub

Sub openTable(Conn, szTable, rs)
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open szTable, Conn, 2, 3, 2
End Sub

Sub SaveData(rs, Action, DataToInsert)
Dim CountElements, i
		CountElements = UBound(DataToInsert, 1) - 1
		if Action =1 then
    	 rs.AddNew
		end if
    For i = 0 To CountElements
        rs(DataToInsert(i)) = DataToInsert(i + 1)
				i = i + 1
    Next
    rs.Update
End Sub

Sub closeOBJs(theOBJ1, theOBJ2)
    On Error Resume Next
    theOBJ1.Close
    theOBJ2.Close
    Set theOBJ1 = Nothing
    Set theOBJ2 = Nothing
End Sub

Function FRegExp(patrn, string1, string2, Action)
Dim regEx ' Create variable.
   If IsNull(string1) Then
    string1 = ""
   End If
   
   If IsNull(string2) Then
    string2 = ""
   End If

   Set regEx = New RegExp ' Create a regular expression.
        regEx.Pattern = patrn   ' Set pattern.
        regEx.IgnoreCase = True   ' Set case insensitivity.
        regEx.Global = True   ' Set global applicability.
   Select Case Action
   Case 1
        Set FRegExp = regEx.Execute(string1)
   Case 2
        FRegExp = regEx.test(string1)
   Case 3
        regEx.Global = False 'Not Set global applicability.
        FRegExp = regEx.Replace(string1, string2)   ' Make replacement.
   Case 4
        regEx.Global = True ' Set global applicability.
        FRegExp = regEx.Replace(string1, string2)   ' Make replacement.
   End Select	 
End Function

Function CheckNum(Data)
CheckNum = 0
If InStr(1, Data, " ") = 0 Then
    If FRegExp("[0-9]*", Data, "", 2) Then
         On Error Resume Next
		CheckNum = CDbl(Data)
    End If
End If
End Function




if Request("OperAjax") <> "" then

    On Error Resume Next

        dim iStatus, iValue, iError, iErrorStr        
        iStatus = "0"
        iError = "0"

	    select case Request("OperAjax")
	    case "GetRoutingErr"        
            OpenConn2 Conn
            SQLQuery = "select id_error from routings_errores where id_routing = " & Request("id_routing") & " AND id_usuario = " & Request("id_usuario") & " AND id_trafico = " & Request("id_trafico") & " LIMIT 1"
            Set rs = Conn.Execute(SQLQuery)
            if Not rs.EOF then             
                iValue = rs(0)               
                if iValue <> "" then
                    iStatus = "1"
                end if
            end if
            CloseOBJs rs, Conn
    
        case "GetTarifa" 'esto funciona para Ajax desde Awb Awbi
            
            iValue = TarifarioPricing (Request("AwbType"), Request("Countries"), Request("ObjectID"), Request("ServiceID"), Request("ItemID"), Request("No"), Request("ItemTarifa"), Request("ItemTarifaHidden"), Request("ItemMonto"), Request("peso"))

            iStatus = "2"
        
        end select    

    If Err.Number <> 0 Then
      iError = Err.Number
      iErrorStr = Err.Description
    End If


	select case Request("OperAjax")
	case "GetRoutingErr"            
        Response.Write iError & "," & iStatus & "," & SQLQuery & "," & iErrorStr & "," & iValue
    
    case "GetTarifa"
        Response.Write iValue 'retorna el resultado a Awb / Awbi : TarifaRubro
    
    end select    

    Response.End

end if


'Function CheckNum2(Data)
'    CheckNum2 = 0
'    If InStr(1, Data, " ") = 0 Then
'        If FRegExp("[0-9]*", Data, "", 2) Then
'             On Error Resume Next
'		    CheckNum2 = CDbl(Data)
'        End If
'    End If    
'    If InStr(1, Data, ".") = 0 Then    '
'	    CheckNum2 = CheckNum2 * 100   
'    End If
'    If CheckNum2 > 0 Then
'        CheckNum2 = CheckNum2 / 100 
'    End If
'    CheckNum2 = Replace(Cstr(CheckNum2),",",".")
'End Function




Function CheckTxt(Data)
CheckTxt = ""
If InStr(1, Data, " ") = 0 Then
    If FRegExp("[a-zA-Z0-9]", Data, "", 2) Then
         CheckTxt = CStr(Data)
    End If
End If
End Function

Sub InsertData(rs, Elements)
Dim CountElements, CantElements, Val, Val2
CountElements = UBound(Elements, 1) - 1
'        rs.AddNew
        For i = 0 To CountElements
        Val = Elements(i)
        i = i + 1
        Val2 = Elements(i)
        Next
 '       rs.Update
End Sub

Sub GetTableData(GroupID, ByRef TableName, ByRef ObjectName, ByRef QuerySelect, ByRef AwbType)
    Select Case GroupID 'Tipo de Grupo = 1 Categoria, 2 Producto, 3 Mensaje, 4 User, 5 Noticia
        Case 1
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
                          "CalcAdminFee, Routing, RoutingID, CTX, TCTX, TPTX, Closed, ConsignerColoader, ShipperColoader, AgentNeutral, ManifestNo "

            if AwbType = 1 then
				QuerySelect = QuerySelect & ", id_coloader" & _
                          ", TotCarrierRate_Routing, FuelSurcharge_Routing, SecurityFee_Routing, CustomFee_Routing, TerminalFee_Routing, PickUp_Routing, SedFilingFee_Routing, Intermodal_Routing, PBA_Routing, TAX_Routing, AdditionalChargeName1_Routing, AdditionalChargeName2_Routing, AdditionalChargeName3_Routing, AdditionalChargeName4_Routing, AdditionalChargeName5_Routing, AdditionalChargeName6_Routing, AdditionalChargeName7_Routing, AdditionalChargeName8_Routing, AdditionalChargeName9_Routing, AdditionalChargeName10_Routing, AdditionalChargeName11_Routing, AdditionalChargeName12_Routing, AdditionalChargeName13_Routing, AdditionalChargeName14_Routing, AdditionalChargeName15_Routing, COALESCE(id_cliente_order,0), id_cliente_orderData, replica, flg_master, flg_totals, file " & _ 
                          " FROM "
				TableName = "Awb"
			else
				QuerySelect = QuerySelect & ", OtherChargeName1, OtherChargeName2, OtherChargeName3, OtherChargeName4, OtherChargeName5, OtherChargeName6, " & _
						  "OtherChargeVal1, OtherChargeVal2, OtherChargeVal3, OtherChargeVal4, OtherChargeVal5, OtherChargeVal6, id_coloader" & _ 
                          ", TotCarrierRate_Routing, FuelSurcharge_Routing, SecurityFee_Routing, PickUp_Routing, SedFilingFee_Routing, Intermodal_Routing, PBA_Routing, AdditionalChargeName1_Routing, AdditionalChargeName2_Routing, AdditionalChargeName3_Routing, AdditionalChargeName4_Routing, AdditionalChargeName5_Routing, AdditionalChargeName6_Routing, AdditionalChargeName7_Routing, AdditionalChargeName8_Routing, AdditionalChargeName9_Routing, AdditionalChargeName10_Routing, AdditionalChargeName11_Routing, AdditionalChargeName12_Routing, AdditionalChargeName13_Routing, AdditionalChargeName14_Routing, AdditionalChargeName15_Routing, OtherChargeName1_Routing, OtherChargeName2_Routing, OtherChargeName3_Routing, OtherChargeName4_Routing, OtherChargeName5_Routing, OtherChargeName6_Routing, COALESCE(id_cliente_order,0), id_cliente_orderData, ReplicaAwbID, flg_master, flg_totals, file " & _ 
                          " FROM "
				TableName = "Awbi"
			end if	

            ObjectName = "AWBID"
        Case 2
            QuerySelect = "select CarrierID, CreatedDate, CreatedTime, Expired, Name, CarrierCode, " & _
						  "ver_AWBNumber, hor_AWBNumber, ver_AccountShipperNo, hor_AccountShipperNo, ver_ShipperData, hor_ShipperData, " & _
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
						  "ver_AdditionalChargeName4, hor_AdditionalChargeName4, ver_AdditionalChargeVal4, hor_AdditionalChargeVal4, Countries, " & _
						  "ver_ChargeType2, hor_ChargeType2, ver_ValChargeType2, hor_ValChargeType2, ver_OtherChargeType2, hor_OtherChargeType2, OtherChargesPrintType, ComisionRate, " & _
                          "ver_AWBNumber2, hor_AWBNumber2, ver_AWBNumber3, hor_AWBNumber3, ver_id_cliente_order, hor_id_cliente_order from "
            TableName = "Carriers"
            ObjectName = "CarrierID"
        Case 3
            QuerySelect = "select CarrierID, CreatedDate, CreatedTime, Expired, AirportID, TerminalFeePD, TerminalFeeCS, CustomFee, FuelSurcharge, SecurityFee from "
            TableName = "CarrierDepartures"
            ObjectName = "CarrierDepartureID"
        Case 4
            QuerySelect = "select AWBID, CreatedDate, CreatedTime, Expired, AWBNumber, ReservationDate, DeliveryDate, DepartureDate, Comment, ManifestNo, flg_master, flg_totals from "
            if AwbType = 1 then
				TableName = "Awb"
			else
				TableName = "Awbi"
			end if			
            ObjectName = "AWBID"
        Case 5
            QuerySelect = "select CarrierID, CreatedDate, CreatedTime, Expired, RangeID from "
            TableName = "CarrierRanges"
            ObjectName = "CarrierRangeID"
        Case 6
            QuerySelect = "select AWBID, CreatedDate, CreatedTime, Expired, AWBNumber, Comment2, ManifestNo, flg_master, flg_totals from "
            if AwbType = 1 then
				TableName = "Awb"
			else
				TableName = "Awbi"
			end if			
            ObjectName = "AWBID"
		Case 7
            'QuerySelect = "select ConsignerID, CreatedDate, CreatedTime, Expired, Name, Address, Phone1, Phone2, AccountNo, Attn, Countries, Address2 from "
            'TableName = "Consigners"
            'ObjectName = "ConsignerID"
			QuerySelect = "select c.id_cliente, c.fecha_creacion, c.hora_creacion, c.id_estatus, c.nombre_cliente, c.nombre_facturar, " & _
							"d.direccion_completa, d.id_direccion, d.""phone number"", p.codigo, n.id_nivel, c.codigo_tributario, " & _
							"c.es_consigneer, c.es_shipper, c.id_grupo, c.id_usuario_creacion, c.id_usuario_modificacion, c.id_pais " & _
							"from clientes c, direcciones d, niveles_geograficos n, paises p " & _
							"where c.id_cliente = d.id_cliente " & _
							"and d.id_nivel_geografico = n.id_nivel " & _
							"and n.id_pais = p.codigo " '& _
							'"and c.es_consigneer = true "
			'if SearchOption<>0 then
			'	QuerySelect = QuerySelect & " and c.es_consigneer = true "
			'end if						
            TableName = "clientes"
            ObjectName = "id_cliente"
		Case 8
            QuerySelect = "select agente_id, fecha_creacion, hora_creacion, activo, agente, direccion, telefono, fax, correo, contacto, id_grupo, id_usuario_creacion, id_usuario_modificacion,	accountno, iatano, defaultval from "
            TableName = "agentes"
            ObjectName = "agente_id"
            'QuerySelect = "select AgentID, CreatedDate, CreatedTime, Expired, Name, Address, Phone1, Phone2, AccountNo, IATANo, DefaultVal, Countries, Address2 from "
            'TableName = "Agents"
            'ObjectName = "AgentID"
		Case 9
            QuerySelect = "select AirportID, CreatedDate, CreatedTime, Expired, Name, AirportCode, Country from "
            TableName = "Airports"
            ObjectName = "AirportID"
		Case 10
            'QuerySelect = "select ShipperID, CreatedDate, CreatedTime, Expired, Name, Address, Phone1, Phone2, AccountNo, Countries, Address2 from "
            'TableName = "Shippers"
            'ObjectName = "ShipperID"
			QuerySelect = "select c.id_cliente, c.fecha_creacion, c.hora_creacion, c.id_estatus, c.nombre_cliente, c.nombre_facturar, " & _
							"d.direccion_completa, d.id_direccion, d.""phone number"", p.codigo, n.id_nivel, c.codigo_tributario, " & _
							"c.es_consigneer, c.es_shipper, c.id_grupo, c.id_usuario_creacion, c.id_usuario_modificacion, c.id_pais " & _
							"from clientes c, direcciones d, niveles_geograficos n, paises p " & _
							"where c.id_cliente = d.id_cliente " & _
							"and d.id_nivel_geografico = n.id_nivel " & _
							"and n.id_pais = p.codigo " '& _
							'"and c.es_shipper = true "
			'if SearchOption<>0 then
			'	QuerySelect = QuerySelect & " and c.es_shipper = true "
			'end if						
            TableName = "clientes"
            ObjectName = "id_cliente"
		Case 11
            QuerySelect = "select commodityid, createddate, createdtime, expired, namees, nameen, typeval, commoditycode, arancel_gt, arancel_sv, arancel_hn, arancel_ni, arancel_cr, arancel_pa, arancel_bz from "
            TableName = "commodities"
            ObjectName = "commodityid"
		Case 12
            QuerySelect = "select CurrencyID, CreatedDate, CreatedTime, Expired, Name, CurrencyCode, Xchange, Countries, Symbol from "
            TableName = "Currencies"
            ObjectName = "CurrencyID"
		Case 13
            QuerySelect = "select RangeID, CreatedDate, CreatedTime, Expired, Val from "
            TableName = "Ranges"
            ObjectName = "RangeID"
		Case 14
            QuerySelect = "select TaxID, CreatedDate, CreatedTime, Expired, Tax, Countries from "
            TableName = "Taxes"
            ObjectName = "TaxID"
		Case 15
            QuerySelect = "select AWBID, CreatedDate, CreatedTime, Expired, HAWBNumber, ArrivalDate, HDepartureDate, Cont, Destinity, TotalToPay, Concept, FiscalFactory, ArrivalAttn, ArrivalFlight, Comment3, ManifestNumber, flg_master, flg_totals from "
            if AwbType = 1 then
				TableName = "Awb"
			else
				TableName = "Awbi"
			end if			
            ObjectName = "AWBID"
        Case 18
            QuerySelect = "select TrackingID, CreatedDate, CreatedTime, AWBID, ClientID, Comment, BLStatus, DocTyp from "
            TableName = "Tracking"
            ObjectName = "TrackingID"		
        Case 21              
            QuerySelect = "SELECT GuideID, GuideNumber, GuideStatus, CreatedDate, CreatedUser, UpdatedUser, GuideActive, GuideType, GuideCarrierID, Comentarios, CreatedTime, UpdatedDate, UpdatedTime FROM "
            TableName = "Guides"
            ObjectName = "GuideID"		

        Case 22
            '                       0           1           2           3       4           5           6       7       8           9       10      11          12      13          14          15          16              17              18      19          20          21              22          23         24       25,     26                     
            QuerySelect = "SELECT MedicionID, CreatedDate, CreatedTime, AwbID, AwbNumber, HAwbNumber, AwbType, DateUno, TimeUno, DateDos, TimeDos, DateTres, TimeTres, DateCuatro, TimeCuatro, MedicionUno, MedicionDos, TotNoOfPieces, TotWeight, Destinity, ShipperData, UserInsert, UserUpdate, DateUpdate, Status,  DateCinco, TimeCinco FROM "
            TableName = "mediciones"
            'ObjectName = "MedicionID"
            ObjectName = "AwbID"

    End Select
End Sub

Function SetOn(Val)
	if Val = "on" then
		SetOn = 0
	else
		SetOn = 1
	end if	
End Function

Function SetActive(Val)
	if Val = "on" then
		SetActive = 1
	else
		SetActive = 0
	end if	
End Function

Sub SaveInfo(Conn, rs, Action, GroupID, CreatedDate, CreatedTime, AwbType)
If Action = 1 Then
    rs.AddNew
	select case GroupID
	case 7, 8, 10
		rs("fecha_creacion") = CreatedDate 
	case else
    	rs("CreatedDate") = CreatedDate 
	end select
End If

Select Case GroupID
Case 1 'AWB
	rs("CreatedTime") = CreatedTime 
    rs("Expired") = SetOn(Request.Form("Expired"))
	'rs("AWBNumber") = Request.Form("AWBNumber")

	if AwbType <> 1 then '1 export 2 import  

		rs("OtherChargeName1") = Request.Form("OtherChargeName1")
		rs("OtherChargeName2") = Request.Form("OtherChargeName2")
		rs("OtherChargeName3") = Request.Form("OtherChargeName3")
		rs("OtherChargeName4") = Request.Form("OtherChargeName4")
		rs("OtherChargeName5") = Request.Form("OtherChargeName5")
		rs("OtherChargeName6") = Request.Form("OtherChargeName6")
		rs("OtherChargeVal1") = Request.Form("OtherChargeVal1")
		rs("OtherChargeVal2") = Request.Form("OtherChargeVal2")
		rs("OtherChargeVal3") = Request.Form("OtherChargeVal3")
		rs("OtherChargeVal4") = Request.Form("OtherChargeVal4")
		rs("OtherChargeVal5") = Request.Form("OtherChargeVal5")
		rs("OtherChargeVal6") = Request.Form("OtherChargeVal6")


        rs("TotCarrierRate_Routing") = CheckNum(Request.Form("TotCarrierRate_Routing"))
        rs("FuelSurcharge_Routing") = CheckNum(Request.Form("FuelSurcharge_Routing"))
        rs("SecurityFee_Routing") = CheckNum(Request.Form("SecurityFee_Routing"))
        rs("PickUp_Routing") = CheckNum(Request.Form("PickUp_Routing"))
        rs("SedFilingFee_Routing") = CheckNum(Request.Form("SedFilingFee_Routing"))
        rs("Intermodal_Routing") = CheckNum(Request.Form("Intermodal_Routing"))
        'rs("PBA_Routing") = CheckNum(Request.Form("PBA_Routing")
        rs("AdditionalChargeName1_Routing") = CheckNum(Request.Form("AdditionalChargeName1_Routing"))
        rs("AdditionalChargeName2_Routing") = CheckNum(Request.Form("AdditionalChargeName2_Routing"))
        rs("AdditionalChargeName3_Routing") = CheckNum(Request.Form("AdditionalChargeName3_Routing"))
        rs("AdditionalChargeName4_Routing") = CheckNum(Request.Form("AdditionalChargeName4_Routing"))
        rs("AdditionalChargeName5_Routing") = CheckNum(Request.Form("AdditionalChargeName5_Routing"))
        rs("AdditionalChargeName6_Routing") = CheckNum(Request.Form("AdditionalChargeName6_Routing"))
        rs("AdditionalChargeName7_Routing") = CheckNum(Request.Form("AdditionalChargeName7_Routing"))
        rs("AdditionalChargeName8_Routing") = CheckNum(Request.Form("AdditionalChargeName8_Routing"))
        rs("AdditionalChargeName9_Routing") = CheckNum(Request.Form("AdditionalChargeName9_Routing"))
        rs("AdditionalChargeName10_Routing") = CheckNum(Request.Form("AdditionalChargeName10_Routing"))
        rs("AdditionalChargeName11_Routing") = CheckNum(Request.Form("AdditionalChargeName11_Routing"))
        rs("AdditionalChargeName12_Routing") = CheckNum(Request.Form("AdditionalChargeName12_Routing"))
        rs("AdditionalChargeName13_Routing") = CheckNum(Request.Form("AdditionalChargeName13_Routing"))
        rs("AdditionalChargeName14_Routing") = CheckNum(Request.Form("AdditionalChargeName14_Routing"))
        rs("AdditionalChargeName15_Routing") = CheckNum(Request.Form("AdditionalChargeName15_Routing"))
        rs("OtherChargeName1_Routing") = CheckNum(Request.Form("OtherChargeName1_Routing"))
        rs("OtherChargeName2_Routing") = CheckNum(Request.Form("OtherChargeName2_Routing"))
        rs("OtherChargeName3_Routing") = CheckNum(Request.Form("OtherChargeName3_Routing"))
        rs("OtherChargeName4_Routing") = CheckNum(Request.Form("OtherChargeName4_Routing"))
        rs("OtherChargeName5_Routing") = CheckNum(Request.Form("OtherChargeName5_Routing"))
        rs("OtherChargeName6_Routing") = CheckNum(Request.Form("OtherChargeName6_Routing"))        

    else

        rs("TotCarrierRate_Routing") = CheckNum(Request.Form("TotCarrierRate_Routing"))
        rs("FuelSurcharge_Routing") = CheckNum(Request.Form("FuelSurcharge_Routing"))
        rs("SecurityFee_Routing") = CheckNum(Request.Form("SecurityFee_Routing"))
        rs("CustomFee_Routing") = CheckNum(Request.Form("CustomFee_Routing"))
        rs("TerminalFee_Routing") = CheckNum(Request.Form("TerminalFee_Routing"))
        rs("PickUp_Routing") = CheckNum(Request.Form("PickUp_Routing"))
        rs("SedFilingFee_Routing") = CheckNum(Request.Form("SedFilingFee_Routing"))
        rs("Intermodal_Routing") = CheckNum(Request.Form("Intermodal_Routing"))
        rs("PBA_Routing") = CheckNum(Request.Form("PBA_Routing"))
        rs("TAX_Routing") = CheckNum(Request.Form("TAX_Routing"))
        rs("AdditionalChargeName1_Routing") = CheckNum(Request.Form("AdditionalChargeName1_Routing"))
        rs("AdditionalChargeName2_Routing") = CheckNum(Request.Form("AdditionalChargeName2_Routing"))
        rs("AdditionalChargeName3_Routing") = CheckNum(Request.Form("AdditionalChargeName3_Routing"))
        rs("AdditionalChargeName4_Routing") = CheckNum(Request.Form("AdditionalChargeName4_Routing"))
        rs("AdditionalChargeName5_Routing") = CheckNum(Request.Form("AdditionalChargeName5_Routing"))
        rs("AdditionalChargeName6_Routing") = CheckNum(Request.Form("AdditionalChargeName6_Routing"))
        rs("AdditionalChargeName7_Routing") = CheckNum(Request.Form("AdditionalChargeName7_Routing"))
        rs("AdditionalChargeName8_Routing") = CheckNum(Request.Form("AdditionalChargeName8_Routing"))
        rs("AdditionalChargeName9_Routing") = CheckNum(Request.Form("AdditionalChargeName9_Routing"))
        rs("AdditionalChargeName10_Routing") = CheckNum(Request.Form("AdditionalChargeName10_Routing"))
        rs("AdditionalChargeName11_Routing") = CheckNum(Request.Form("AdditionalChargeName11_Routing"))
        rs("AdditionalChargeName12_Routing") = CheckNum(Request.Form("AdditionalChargeName12_Routing"))
        rs("AdditionalChargeName13_Routing") = CheckNum(Request.Form("AdditionalChargeName13_Routing"))
        rs("AdditionalChargeName14_Routing") = CheckNum(Request.Form("AdditionalChargeName14_Routing"))
        rs("AdditionalChargeName15_Routing") = CheckNum(Request.Form("AdditionalChargeName15_Routing"))        

	end if

	rs("AccountShipperNo") = Request.Form("AccountShipperNo")
	rs("ShipperData") = Request.Form("ShipperData")
	rs("AccountConsignerNo") = Request.Form("AccountConsignerNo")
	rs("ConsignerData") = Request("ConsignerData")
	rs("AgentData") = Request.Form("AgentData")
	rs("AccountInformation") = Request.Form("AccountInformation")
	rs("IATANo") = Request.Form("IATANo")
	rs("AccountAgentNo") = Request.Form("AccountAgentNo")
	rs("AirportDepID") = Request.Form("AirportDepID")
	rs("RequestedRouting") = Request.Form("RequestedRouting")
	rs("AirportToCode1") = Request.Form("AirportToCode1")
	rs("CarrierID") = Request.Form("CarrierID")
	rs("AirportToCode2") = Request.Form("AirportToCode2")
	rs("AirportToCode3") = Request.Form("AirportToCode3")
	rs("CarrierCode2") = Request.Form("CarrierCode2")
	rs("CarrierCode3") = Request.Form("CarrierCode3")
	rs("CurrencyID") = Request.Form("CurrencyID")
	rs("ChargeType") = Request.Form("ChargeType")
	rs("ValChargeType") = Request.Form("ValChargeType")
	rs("OtherChargeType") = Request.Form("OtherChargeType")
	rs("DeclaredValue") = Request.Form("DeclaredValue")
	rs("AduanaValue") = Request.Form("AduanaValue")
	rs("AirportDesID") = Request.Form("AirportDesID")
	rs("FlightDate1") = Request.Form("FlightDate1")
	rs("FlightDate2") = Request.Form("FlightDate2")
	rs("SecuredValue") = Request.Form("SecuredValue")
	rs("HandlingInformation") = Request.Form("HandlingInformation")
	rs("Observations") = Request.Form("Observations")
	rs("NoOfPieces") = Request.Form("NoOfPieces")
	rs("Weights") = Request.Form("Weights")
	rs("WeightsSymbol") = Request.Form("WeightsSymbol")
	rs("Commodities") = Request.Form("Commodities")
	rs("ChargeableWeights") = Request.Form("ChargeableWeights")
	rs("CarrierRates") = Request.Form("CarrierRates")
	rs("CarrierSubTot") = Request.Form("CarrierSubTot")
	rs("NatureQtyGoods") = Request.Form("NatureQtyGoods")
	rs("TotNoOfPieces") = Request.Form("TotNoOfPieces")
	rs("TotWeight") = Request.Form("TotWeight")
	rs("TotCarrierRate") = Request.Form("TotCarrierRate")
	rs("TotChargeWeightPrepaid") = Request.Form("TotChargeWeightPrepaid")
	rs("TotChargeWeightCollect") = Request.Form("TotChargeWeightCollect")
	rs("TotChargeValuePrepaid") = Request.Form("TotChargeValuePrepaid")
	rs("TotChargeValueCollect") = Request.Form("TotChargeValueCollect")
	rs("TotChargeTaxPrepaid") = Request.Form("TotChargeTaxPrepaid")
	rs("TotChargeTaxCollect") = Request.Form("TotChargeTaxCollect")
	rs("AnotherChargesAgentPrepaid") = Request.Form("AnotherChargesAgentPrepaid")
	rs("AnotherChargesAgentCollect") = Request.Form("AnotherChargesAgentCollect")
	rs("AnotherChargesCarrierPrepaid") = Request.Form("AnotherChargesCarrierPrepaid")
	rs("AnotherChargesCarrierCollect") = Request.Form("AnotherChargesCarrierCollect")
	rs("TotPrepaid") = Request.Form("TotPrepaid")
	rs("TotCollect") = Request.Form("TotCollect")
	rs("TerminalFee") = Request.Form("TerminalFee")
	rs("CustomFee") = Request.Form("CustomFee")
	rs("FuelSurcharge") = Request.Form("FuelSurcharge")
	rs("SecurityFee") = Request.Form("SecurityFee")
	rs("PBA") = Request.Form("PBA")
	rs("TAX") = Request.Form("TAX")
	rs("AdditionalChargeName1") = Replace(Request.Form("AdditionalChargeName1"),chr(13)&chr(10),"",1,-1)
	rs("AdditionalChargeVal1") = Request.Form("AdditionalChargeVal1")
	rs("AdditionalChargeName2") = Request.Form("AdditionalChargeName2")
	rs("AdditionalChargeVal2") = Request.Form("AdditionalChargeVal2")
	rs("Invoice") = Request.Form("Invoice")
	rs("ExportLic") = Request.Form("ExportLic")
	rs("AgentContactSignature") = Request.Form("AgentContactSignature")
	rs("CommoditiesTypes") = Request.Form("CommoditiesTypes")
	rs("TotWeightChargeable") = Request.Form("TotWeightChargeable")
	rs("Instructions") = Request.Form("Instructions")
	rs("AgentSignature") = Request.Form("AgentSignature")
	dim AWBDate 
	AWBDate = CreatedDate        
	if Request.Form("AWBDate") <> "" then
		AWBDate = Mid(Request.Form("AWBDate"),7,4) &"-"& Mid(Request.Form("AWBDate"),4,2)  &"-"& Mid(Request.Form("AWBDate"),1,2)
	end if
	rs("AWBDate") = AWBDate
	rs("AdditionalChargeName3") = Request.Form("AdditionalChargeName3")
	rs("AdditionalChargeVal3") = Request.Form("AdditionalChargeVal3")
	rs("AdditionalChargeName4") = Request.Form("AdditionalChargeName4")
	rs("AdditionalChargeVal4") = Request.Form("AdditionalChargeVal4")
	rs("Countries") = Request.Form("Countries")

    if rs("Countries") = "GT" and AwbType = 1 then   '1 export 2 import  
		rs("HAWBNumber") = HAWBNumber
		rs("AWBNumber") = AWBNumber
	else
		if CheckHAWBNumber(Conn, AwbType, Trim(Request.Form("HAWBNumber"))) = 0 then
			rs("HAWBNumber") = Trim(Request.Form("HAWBNumber"))
		end if		
		rs("AWBNumber") = Request.Form("AWBNumber")
	end if

	rs("AdditionalChargeName5") = Request.Form("AdditionalChargeName5")
	rs("AdditionalChargeVal5") = Request.Form("AdditionalChargeVal5")
	rs("AdditionalChargeName6") = Request.Form("AdditionalChargeName6")
	rs("AdditionalChargeVal6") = Request.Form("AdditionalChargeVal6")
	If Request.Form("DisplayNumber") = "on" Then
		rs("DisplayNumber") = 1
	Else
		rs("DisplayNumber") = 0
	End If
	rs("AdditionalChargeName7") = Request.Form("AdditionalChargeName7")
	rs("AdditionalChargeVal7") = Request.Form("AdditionalChargeVal7")
	rs("AdditionalChargeName8") = Request.Form("AdditionalChargeName8")
	rs("AdditionalChargeVal8") = Request.Form("AdditionalChargeVal8")
	rs("WType") = Request.Form("WType")
	rs("AdditionalChargeName9") = Request.Form("AdditionalChargeName9")
	rs("AdditionalChargeVal9") = Request.Form("AdditionalChargeVal9")
	rs("AdditionalChargeName10") = Request.Form("AdditionalChargeName10")
	rs("AdditionalChargeVal10") = Request.Form("AdditionalChargeVal10")
	rs("ShipperID") = CheckNum(Request.Form("ShipperID"))
	rs("ConsignerID") = CheckNum(Request.Form("ConsignerID"))
	rs("AgentID") = CheckNum(Request.Form("AgentID"))
	rs("SalespersonID") = CheckNum(Request.Form("SalespersonID"))
	rs("ShipperAddrID") = CheckNum(Request.Form("ShipperAddrID"))
	rs("ConsignerAddrID") = CheckNum(Request.Form("ConsignerAddrID"))
	rs("AgentAddrID") = CheckNum(Request.Form("AgentAddrID"))
	rs("AdditionalChargeName11") = Request.Form("AdditionalChargeName11")
	rs("AdditionalChargeVal11") = Request.Form("AdditionalChargeVal11")
	rs("AdditionalChargeName12") = Request.Form("AdditionalChargeName12")
	rs("AdditionalChargeVal12") = Request.Form("AdditionalChargeVal12")
	rs("AdditionalChargeName13") = Request.Form("AdditionalChargeName13")
	rs("AdditionalChargeVal13") = Request.Form("AdditionalChargeVal13")
	rs("AdditionalChargeName14") = Request.Form("AdditionalChargeName14")
	rs("AdditionalChargeVal14") = Request.Form("AdditionalChargeVal14")
	rs("AdditionalChargeName15") = Request.Form("AdditionalChargeName15")
	rs("AdditionalChargeVal15") = Request.Form("AdditionalChargeVal15")
	rs("Voyage") = CheckNum(Request.Form("Voyage"))
	rs("PickUp") = Request.Form("PickUp")
	rs("Intermodal") = Request.Form("Intermodal")
	rs("SedFilingFee") = Request.Form("SedFilingFee")
	rs("CalcAdminFee") = CheckNum(Request.Form("CalcAdminFee"))
	rs("RoutingID") = CheckNum(Request.Form("RoutingID"))
	rs("Routing") = Request.Form("Routing")

    'if rs("Countries") = "GT" and AwbType = 1 then   '1 export 2 import  
    '
    'else        
    '    'Bloqueando el RO para que no lo puedan borrar, porque ya esta asociado al aereo
    '    Dim Conn2
    '    OpenConn2 Conn2 ////////////////////////////////////////////////////////////////////////////// ESTE BLOQUE PASO A InsertData.asp al final del insert
    '        Conn2.Execute("update routings set activo=false where routing='" & Request.Form("Routing") & "'")
    '    CloseOBJ Conn2
    'end if
	rs("CTX") = Request.Form("CTX")
	rs("TCTX") = CheckNum(Request.Form("TCTX"))
	rs("TPTX") = CheckNum(Request.Form("TPTX"))
	rs("UserID") = Session("OperatorID")
	rs("Closed") = CheckNum(Request.Form("Closed"))
    rs("ConsignerColoader") = CheckNum(Request.Form("ConsignerColoader"))
    rs("ShipperColoader") = CheckNum(Request.Form("ShipperColoader"))
    rs("AgentNeutral") = CheckNum(Request.Form("AgentNeutral"))
    rs("ManifestNo") = Request.Form("ManifestNumber")

    if Action = 1 Then 'insert 2017-12-08

        'los inserts de house solo se haran por medio de express
        if rs("HAWBNumber") = "" then 'cuando viene en blanco es master consolidado, lo bloquea, en espera de input de houses
            rs("flg_master") = "1"
            rs("flg_totals") = "1"
        'else 
        '    rs("flg_master") = "0"
        '    rs("flg_totals") = "0"
        end if

    end if

    rs("file") = Request.Form("file")
            
    if Action = 2 Then '2020-05-05
        SQLQuery = "UPDATE Awb_IE_Expansion SET aiee_TipoCarga = '" & Request("TipoCarga2") & "', aiee_Fecha_Update=NOW()  where aiee_AwbID_fk='" & Request("OID") & "' and aiee_ImpExp='" & IIf(AwbType = 1,0,1) & "'"
        'response.write SQLQuery & "<br>"
        Conn.Execute(SQLQuery)
    end if
           
    rs("id_coloader") = CheckNum(Request.Form("id_coloader"))
    rs("id_cliente_order") = CheckNum(Request.Form("id_cliente_order"))
    rs("id_cliente_orderData") = Request.Form("id_cliente_orderData")

    if AwbType = 1 then   '1 export 2 import  
        
		rs("replica") = Request.Form("replica")

        if Request.Form("replica") = "Master-Hija" then
            rs("replica") = "Consolidado"
        end if

        if Request.Form("replica") = "Hija-Directa" then
            rs("replica") = "Directo"
        end if

        if Request.Form("replica") = "Master-Master-Hija" then
            rs("replica") = "Consolidado"
        end if

    end if
    

    'response.write "(0)(" & Request.Form("iMinimo") & ")"
    'if Request.Form("iMinimo") <> "" then
    '    if Request.Form("iMinimo") = "1" then
    '        rs("Comment2") = "MIN"
    '    else
    '        rs("Comment2") = Request.Form("iMinimo")
    '    end if    
    'else
    '    rs("Comment2") = ""
    'end if

    'response.write ( "Saliendo del SaveInfo(" & rs("Countries") & ")(" & AwbType & ")(" & rs("AWBNumber") & ")(" & rs("HAWBNumber") & ")<br>" )

    if rs("Countries") = "GT" and AwbType = 1 then   '1 export 2 import  
        
    else
	    'Seteando el ID del viaje en los House, con el ID de la guia Master
	    Dim MAWBID, rst '////////////////////////////////////////////////////////////////////////////// ESTE BLOQUE PASO A InsertData.asp al final del insert
	    MAWBID = 0
	    if AwbType <> 1 then 'Import
		    Set rst = Conn.Execute("select AWBID from Awbi where AWBNumber='" & Request.Form("AWBNumber") & "' and HAWBNumber=''")
	    else 'Export
		    Set rst = Conn.Execute("select AWBID from Awb where AWBNumber='" & Request.Form("AWBNumber") & "' and HAWBNumber=''")
	    end if
	    if Not rst.EOF then
		    MAWBID = CheckNum(rst(0))
	    end if
	    CloseOBJ rst
	    rs("MAWBID") = MAWBID

	    'Actualizando la BD Master indicando la fecha y tipo de servicio que realizo el cliente y el shipper
	    OpenConn2 ConnMaster
		    'response.write "update clientes set ultima_fecha_descarga='" & ConvertDate(Now,2) & "', ultimo_tipo_movimiento=1 where id_cliente in (" & CheckNum(Request.Form("ShipperID")) & "," & CheckNum(Request.Form("ConsignerID")) & ")<br>"
            '2016-09-05 se agrego id_estatus=1 solicitado por Cesar en reunion del dia de hoy
		    ConnMaster.Execute("update clientes set id_estatus=1, ultima_fecha_descarga='" & ConvertDate(Now,2) & "', ultimo_tipo_movimiento=1 where id_cliente in (" & CheckNum(Request.Form("ShipperID")) & "," & CheckNum(Request.Form("ConsignerID")) & ")")
	    CloseOBJ ConnMaster

    end if

Case 2 'Transportistas - Shippers
	Dim CarrierID, Name, Countries, CarrierCode, ObjectID, rs2
	Name = PurgeData(Request.Form("Name"))
	Countries = Request.Form("Countries")
	CarrierCode = PurgeData(Request.Form("CarrierCode"))
	ObjectID = CheckNum(Request("OID"))
	
	OpenConn2 ConnMaster
	select case Action
	case 1
		ConnMaster.execute("insert into Carriers (name, countries, carriercode) values ('" & Name & "', '" & Countries & "', '" & CarrierCode & "')")
		set rs2 = ConnMaster.execute("select carrier_id from Carriers where name='" & Name & "' and countries='" & Countries & "' and carriercode='" & CarrierCode & "'")
		if Not rs2.EOF then
			ObjectID = rs2(0)
		end if
		CloseOBJ rs2
	Case 2
		ConnMaster.execute("update Carriers set name='" & Name & "', countries='" & Countries & "', carriercode='" & CarrierCode & "' where carrier_id=" & ObjectID)
	end select
	CloseOBJ ConnMaster
	
	rs("CarrierID") = ObjectID
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("Name") = Name
	rs("CarrierCode") = CarrierCode
	rs("Countries") = Countries
	rs("ComisionRate") = CheckNum(Request.Form("ComisionRate"))
	If Action = 2 then
		rs("ver_AirportCode") = CheckNum(Request.Form("ver_AirportCode"))
		rs("hor_AirportCode") = CheckNum(Request.Form("hor_AirportCode"))
		rs("ver_AWBNumber") = CheckNum(Request.Form("ver_AWBNumber"))
		rs("hor_AWBNumber") = CheckNum(Request.Form("hor_AWBNumber"))

		rs("ver_AWBNumber2") = CheckNum(Request.Form("ver_AWBNumber2"))
		rs("hor_AWBNumber2") = CheckNum(Request.Form("hor_AWBNumber2"))

		rs("ver_AWBNumber3") = CheckNum(Request.Form("ver_AWBNumber3"))
		rs("hor_AWBNumber3") = CheckNum(Request.Form("hor_AWBNumber3"))

		rs("ver_id_cliente_order") = CheckNum(Request.Form("ver_id_cliente_order"))
		rs("hor_id_cliente_order") = CheckNum(Request.Form("hor_id_cliente_order"))

		rs("ver_AccountShipperNo") = CheckNum(Request.Form("ver_AccountShipperNo"))
		rs("hor_AccountShipperNo") = CheckNum(Request.Form("hor_AccountShipperNo"))
		rs("ver_ShipperData") = CheckNum(Request.Form("ver_ShipperData"))
		rs("hor_ShipperData") = CheckNum(Request.Form("hor_ShipperData"))
		rs("ver_AccountConsignerNo") = CheckNum(Request.Form("ver_AccountConsignerNo"))
		rs("hor_AccountConsignerNo") = CheckNum(Request.Form("hor_AccountConsignerNo"))
		rs("ver_ConsignerData") = CheckNum(Request.Form("ver_ConsignerData"))
		rs("hor_ConsignerData") = CheckNum(Request.Form("hor_ConsignerData"))
		rs("ver_AgentData") = CheckNum(Request.Form("ver_AgentData"))
		rs("hor_AgentData") = CheckNum(Request.Form("hor_AgentData"))
		rs("ver_AccountInformation") = CheckNum(Request.Form("ver_AccountInformation"))
		rs("hor_AccountInformation") = CheckNum(Request.Form("hor_AccountInformation"))
		rs("ver_IATANo") = CheckNum(Request.Form("ver_IATANo"))
		rs("hor_IATANo") = CheckNum(Request.Form("hor_IATANo"))
		rs("ver_AccountAgentNo") = CheckNum(Request.Form("ver_AccountAgentNo"))
		rs("hor_AccountAgentNo") = CheckNum(Request.Form("hor_AccountAgentNo"))
		rs("ver_AirportDepID") = CheckNum(Request.Form("ver_AirportDepID"))
		rs("hor_AirportDepID") = CheckNum(Request.Form("hor_AirportDepID"))
		rs("ver_RequestedRouting") = CheckNum(Request.Form("ver_RequestedRouting"))
		rs("hor_RequestedRouting") = CheckNum(Request.Form("hor_RequestedRouting"))
		rs("ver_AirportToCode1") = CheckNum(Request.Form("ver_AirportToCode1"))
		rs("hor_AirportToCode1") = CheckNum(Request.Form("hor_AirportToCode1"))
		rs("ver_CarrierID") = CheckNum(Request.Form("ver_CarrierID"))
		rs("hor_CarrierID") = CheckNum(Request.Form("hor_CarrierID"))
		rs("ver_AirportToCode2") = CheckNum(Request.Form("ver_AirportToCode2"))
		rs("hor_AirportToCode2") = CheckNum(Request.Form("hor_AirportToCode2"))
		rs("ver_AirportToCode3") = CheckNum(Request.Form("ver_AirportToCode3"))
		rs("hor_AirportToCode3") = CheckNum(Request.Form("hor_AirportToCode3"))
		rs("ver_CarrierCode2") = CheckNum(Request.Form("ver_CarrierCode2"))
		rs("hor_CarrierCode2") = CheckNum(Request.Form("hor_CarrierCode2"))
		rs("ver_CarrierCode3") = CheckNum(Request.Form("ver_CarrierCode3"))
		rs("hor_CarrierCode3") = CheckNum(Request.Form("hor_CarrierCode3"))
		rs("ver_CurrencyID") = CheckNum(Request.Form("ver_CurrencyID"))
		rs("hor_CurrencyID") = CheckNum(Request.Form("hor_CurrencyID"))
		rs("ver_ChargeType") = CheckNum(Request.Form("ver_ChargeType"))
		rs("hor_ChargeType") = CheckNum(Request.Form("hor_ChargeType"))
		rs("ver_ValChargeType") = CheckNum(Request.Form("ver_ValChargeType"))
		rs("hor_ValChargeType") = CheckNum(Request.Form("hor_ValChargeType"))
		rs("ver_OtherChargeType") = CheckNum(Request.Form("ver_OtherChargeType"))
		rs("hor_OtherChargeType") = CheckNum(Request.Form("hor_OtherChargeType"))
		rs("ver_DeclaredValue") = CheckNum(Request.Form("ver_DeclaredValue"))
		rs("hor_DeclaredValue") = CheckNum(Request.Form("hor_DeclaredValue"))
		rs("ver_AduanaValue") = CheckNum(Request.Form("ver_AduanaValue"))
		rs("hor_AduanaValue") = CheckNum(Request.Form("hor_AduanaValue"))
		rs("ver_AirportDesID") = CheckNum(Request.Form("ver_AirportDesID"))
		rs("hor_AirportDesID") = CheckNum(Request.Form("hor_AirportDesID"))
		rs("ver_FlightDate1") = CheckNum(Request.Form("ver_FlightDate1"))
		rs("hor_FlightDate1") = CheckNum(Request.Form("hor_FlightDate1"))
		rs("ver_FlightDate2") = CheckNum(Request.Form("ver_FlightDate2"))
		rs("hor_FlightDate2") = CheckNum(Request.Form("hor_FlightDate2"))
		rs("ver_SecuredValue") = CheckNum(Request.Form("ver_SecuredValue"))
		rs("hor_SecuredValue") = CheckNum(Request.Form("hor_SecuredValue"))
		rs("ver_HandlingInformation") = CheckNum(Request.Form("ver_HandlingInformation"))
		rs("hor_HandlingInformation") = CheckNum(Request.Form("hor_HandlingInformation"))
		rs("ver_Observations") = CheckNum(Request.Form("ver_Observations"))
		rs("hor_Observations") = CheckNum(Request.Form("hor_Observations"))
		rs("ver_NoOfPieces") = CheckNum(Request.Form("ver_NoOfPieces"))
		rs("hor_NoOfPieces") = CheckNum(Request.Form("hor_NoOfPieces"))
		rs("ver_Weights") = CheckNum(Request.Form("ver_Weights"))
		rs("hor_Weights") = CheckNum(Request.Form("hor_Weights"))
		rs("ver_WeightsSymbol") = CheckNum(Request.Form("ver_WeightsSymbol"))
		rs("hor_WeightsSymbol") = CheckNum(Request.Form("hor_WeightsSymbol"))
		rs("ver_Commodities") = CheckNum(Request.Form("ver_Commodities"))
		rs("hor_Commodities") = CheckNum(Request.Form("hor_Commodities"))
		rs("ver_ChargeableWeights") = CheckNum(Request.Form("ver_ChargeableWeights"))
		rs("hor_ChargeableWeights") = CheckNum(Request.Form("hor_ChargeableWeights"))
		rs("ver_CarrierRates") = CheckNum(Request.Form("ver_CarrierRates"))
		rs("hor_CarrierRates") = CheckNum(Request.Form("hor_CarrierRates"))
		rs("ver_CarrierSubTot") = CheckNum(Request.Form("ver_CarrierSubTot"))
		rs("hor_CarrierSubTot") = CheckNum(Request.Form("hor_CarrierSubTot"))
		rs("ver_NatureQtyGoods") = CheckNum(Request.Form("ver_NatureQtyGoods"))
		rs("hor_NatureQtyGoods") = CheckNum(Request.Form("hor_NatureQtyGoods"))
		rs("ver_TotNoOfPieces") = CheckNum(Request.Form("ver_TotNoOfPieces"))
		rs("hor_TotNoOfPieces") = CheckNum(Request.Form("hor_TotNoOfPieces"))
		rs("ver_TotWeight") = CheckNum(Request.Form("ver_TotWeight"))
		rs("hor_TotWeight") = CheckNum(Request.Form("hor_TotWeight"))
		rs("ver_TotCarrierRate") = CheckNum(Request.Form("ver_TotCarrierRate"))
		rs("hor_TotCarrierRate") = CheckNum(Request.Form("hor_TotCarrierRate"))
		rs("ver_TotChargeWeightPrepaid") = CheckNum(Request.Form("ver_TotChargeWeightPrepaid"))
		rs("hor_TotChargeWeightPrepaid") = CheckNum(Request.Form("hor_TotChargeWeightPrepaid"))
		rs("ver_TotChargeWeightCollect") = CheckNum(Request.Form("ver_TotChargeWeightCollect"))
		rs("hor_TotChargeWeightCollect") = CheckNum(Request.Form("hor_TotChargeWeightCollect"))
		rs("ver_TotChargeValuePrepaid") = CheckNum(Request.Form("ver_TotChargeValuePrepaid"))
		rs("hor_TotChargeValuePrepaid") = CheckNum(Request.Form("hor_TotChargeValuePrepaid"))
		rs("ver_TotChargeValueCollect") = CheckNum(Request.Form("ver_TotChargeValueCollect"))
		rs("hor_TotChargeValueCollect") = CheckNum(Request.Form("hor_TotChargeValueCollect"))
		rs("ver_TotChargeTaxPrepaid") = CheckNum(Request.Form("ver_TotChargeTaxPrepaid"))
		rs("hor_TotChargeTaxPrepaid") = CheckNum(Request.Form("hor_TotChargeTaxPrepaid"))
		rs("ver_TotChargeTaxCollect") = CheckNum(Request.Form("ver_TotChargeTaxCollect"))
		rs("hor_TotChargeTaxCollect") = CheckNum(Request.Form("hor_TotChargeTaxCollect"))
		rs("ver_AnotherChargesAgentPrepaid") = CheckNum(Request.Form("ver_AnotherChargesAgentPrepaid"))
		rs("hor_AnotherChargesAgentPrepaid") = CheckNum(Request.Form("hor_AnotherChargesAgentPrepaid"))
		rs("ver_AnotherChargesAgentCollect") = CheckNum(Request.Form("ver_AnotherChargesAgentCollect"))
		rs("hor_AnotherChargesAgentCollect") = CheckNum(Request.Form("hor_AnotherChargesAgentCollect"))
		rs("ver_AnotherChargesCarrierPrepaid") = CheckNum(Request.Form("ver_AnotherChargesCarrierPrepaid"))
		rs("hor_AnotherChargesCarrierPrepaid") = CheckNum(Request.Form("hor_AnotherChargesCarrierPrepaid"))
		rs("ver_AnotherChargesCarrierCollect") = CheckNum(Request.Form("ver_AnotherChargesCarrierCollect"))
		rs("hor_AnotherChargesCarrierCollect") = CheckNum(Request.Form("hor_AnotherChargesCarrierCollect"))
		rs("ver_TotPrepaid") = CheckNum(Request.Form("ver_TotPrepaid"))
		rs("hor_TotPrepaid") = CheckNum(Request.Form("hor_TotPrepaid"))
		rs("ver_TotCollect") = CheckNum(Request.Form("ver_TotCollect"))
		rs("hor_TotCollect") = CheckNum(Request.Form("hor_TotCollect"))
		rs("ver_TerminalFee") = CheckNum(Request.Form("ver_TerminalFee"))
		rs("hor_TerminalFee") = CheckNum(Request.Form("hor_TerminalFee"))
		rs("ver_CustomFee") = CheckNum(Request.Form("ver_CustomFee"))
		rs("hor_CustomFee") = CheckNum(Request.Form("hor_CustomFee"))
		rs("ver_FuelSurcharge") = CheckNum(Request.Form("ver_FuelSurcharge"))
		rs("hor_FuelSurcharge") = CheckNum(Request.Form("hor_FuelSurcharge"))
		rs("ver_SecurityFee") = CheckNum(Request.Form("ver_SecurityFee"))
		rs("hor_SecurityFee") = CheckNum(Request.Form("hor_SecurityFee"))
		rs("ver_PBA") = CheckNum(Request.Form("ver_PBA"))
		rs("hor_PBA") = CheckNum(Request.Form("hor_PBA"))
		rs("ver_TAX") = CheckNum(Request.Form("ver_TAX"))
		rs("hor_TAX") = CheckNum(Request.Form("hor_TAX"))
		rs("ver_AdditionalChargeName1") = CheckNum(Request.Form("ver_AdditionalChargeName1"))
		rs("hor_AdditionalChargeName1") = CheckNum(Request.Form("hor_AdditionalChargeName1"))
		rs("ver_AdditionalChargeVal1") = CheckNum(Request.Form("ver_AdditionalChargeVal1"))
		rs("hor_AdditionalChargeVal1") = CheckNum(Request.Form("hor_AdditionalChargeVal1"))
		rs("ver_AdditionalChargeName2") = CheckNum(Request.Form("ver_AdditionalChargeName2"))
		rs("hor_AdditionalChargeName2") = CheckNum(Request.Form("hor_AdditionalChargeName2"))
		rs("ver_AdditionalChargeVal2") = CheckNum(Request.Form("ver_AdditionalChargeVal2"))
		rs("hor_AdditionalChargeVal2") = CheckNum(Request.Form("hor_AdditionalChargeVal2"))
		rs("ver_Invoice") = CheckNum(Request.Form("ver_Invoice"))
		rs("hor_Invoice") = CheckNum(Request.Form("hor_Invoice"))
		rs("ver_ExportLic") = CheckNum(Request.Form("ver_ExportLic"))
		rs("hor_ExportLic") = CheckNum(Request.Form("hor_ExportLic"))
		rs("ver_AgentContactSignature") = CheckNum(Request.Form("ver_AgentContactSignature"))
		rs("hor_AgentContactSignature") = CheckNum(Request.Form("hor_AgentContactSignature"))
		rs("ver_Instructions") = CheckNum(Request.Form("ver_Instructions"))
		rs("hor_Instructions") = CheckNum(Request.Form("hor_Instructions"))
		rs("ver_AgentSignature") = CheckNum(Request.Form("ver_AgentSignature"))
		rs("hor_AgentSignature") = CheckNum(Request.Form("hor_AgentSignature"))
		rs("ver_AWBDate") = CheckNum(Request.Form("ver_AWBDate"))
		rs("hor_AWBDate") = CheckNum(Request.Form("hor_AWBDate"))
		rs("ver_AdditionalChargeName3") = CheckNum(Request.Form("ver_AdditionalChargeName3"))
		rs("hor_AdditionalChargeName3") = CheckNum(Request.Form("hor_AdditionalChargeName3"))
		rs("ver_AdditionalChargeVal3") = CheckNum(Request.Form("ver_AdditionalChargeVal3"))
		rs("hor_AdditionalChargeVal3") = CheckNum(Request.Form("hor_AdditionalChargeVal3"))
		rs("ver_AdditionalChargeName4") = CheckNum(Request.Form("ver_AdditionalChargeName4"))
		rs("hor_AdditionalChargeName4") = CheckNum(Request.Form("hor_AdditionalChargeName4"))
		rs("ver_AdditionalChargeVal4") = CheckNum(Request.Form("ver_AdditionalChargeVal4"))
		rs("hor_AdditionalChargeVal4") = CheckNum(Request.Form("hor_AdditionalChargeVal4"))
		rs("ver_ChargeType2") = CheckNum(Request.Form("ver_ChargeType2"))
		rs("hor_ChargeType2") = CheckNum(Request.Form("hor_ChargeType2"))
		rs("ver_ValChargeType2") = CheckNum(Request.Form("ver_ValChargeType2"))
		rs("hor_ValChargeType2") = CheckNum(Request.Form("hor_ValChargeType2"))
		rs("ver_OtherChargeType2") = CheckNum(Request.Form("ver_OtherChargeType2"))
		rs("hor_OtherChargeType2") = CheckNum(Request.Form("hor_OtherChargeType2"))	
		rs("OtherChargesPrintType") = Request.Form("OtherChargesPrintType")
	end if
Case 3 'Transportistas-Salida
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("CarrierID") = CheckNum(Request.Form("CarrierID"))
	rs("AirportID") = CheckNum(Request.Form("AirportID"))
	rs("TerminalFeePD") = CheckNum(Request.Form("TerminalFeePD"))
	rs("TerminalFeeCS") = CheckNum(Request.Form("TerminalFeeCS"))
	rs("CustomFee") = CheckNum(Request.Form("CustomFee"))
	rs("FuelSurcharge") = CheckNum(Request.Form("FuelSurcharge"))
	rs("SecurityFee") = CheckNum(Request.Form("SecurityFee"))
Case 4 	'Confirmacion de Reserva
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("ReservationDate") = Request.Form("ReservationDate")	
	rs("DeliveryDate") = Request.Form("DeliveryDate")	
	rs("DepartureDate") = Request.Form("DepartureDate")	
	rs("Comment") = Request.Form("Comment")
Case 5 'Transportistas-Rango
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("CarrierID") = CheckNum(Request.Form("CarrierID"))
	rs("RangeID") = CheckNum(Request.Form("RangeID"))
Case 6 	'House Cargo Manifiesto
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("Comment2") = Request.Form("Comment2")
Case 7, 10 'Destinatarios - Consigner / Embarcadores - Shippers
	rs("hora_creacion") = CreatedTime 
    rs("id_estatus") = Request.Form("Expired")
	rs("nombre_cliente") = PurgeData(Request.Form("Name"))
	rs("nombre_facturar") = PurgeData(Request.Form("BillName"))
	'rs("codigo_tributario") = PurgeData(Request.Form("Tax"))
	rs("es_consigneer") = SetActive(Request.Form("isConsigneer")) 
	rs("es_shipper") = SetActive(Request.Form("isShipper"))
	rs("id_grupo") = CheckNum(Request.Form("BGID"))
	rs("id_pais") = Request.Form("CreatedIn")
	Select Case Action
	Case 1
		rs("id_usuario_creacion") = CheckNum(Session("OperatorID"))
	Case 2
		rs("id_usuario_modificacion") = CheckNum(Session("OperatorID"))
	End Select
Case 8	'Agentes
	rs("hora_creacion") = CreatedTime 
	rs("activo") = SetActive(Request.Form("Expired"))
	rs("agente") = PurgeData(Request.Form("Name"))
	rs("direccion") = PurgeData(Request.Form("Address"))
	rs("telefono") = PurgeData(Request.Form("Phone1"))
	rs("fax") = PurgeData(Request.Form("Phone2"))
	rs("contacto") = PurgeData(Request.Form("Attn"))
	rs("correo") = PurgeData(Request.Form("Email"))
	rs("id_grupo") = CheckNum(Request.Form("BGID"))
	Select Case Action
	Case 1
		rs("id_usuario_creacion") = CheckNum(Session("OperatorID"))
	Case 2
		rs("id_usuario_modificacion") = CheckNum(Session("OperatorID"))
	end Select
	rs("accountno") = PurgeData(Request.Form("AccountNo"))
	rs("iatano") = PurgeData(Request.Form("IATANo"))
	If Request.Form("DefaultVal") = "on" Then
    	rs("defaultval") = 1
	Else
		rs("defaultval") = 0
	End If
	rs("countries") = Request.Form("Countries")
Case 9	'Aeropuertos
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("AirportCode") = PurgeData(Request.Form("AirportCode"))
	rs("Name") = PurgeData(Request.Form("Name"))
    rs("Country") = Request.Form("Countries")
Case 11	'Commodities
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("CommodityCode") = Request.Form("CommodityCode")
	rs("NameES") = PurgeData(Request.Form("NameES"))
	rs("NameEN") = PurgeData(Request.Form("NameEN"))
	rs("TypeVal") = Request.Form("TypeVal")
	rs("ReqAuth") = 0
	rs("Arancel_GT") = PurgeData(Request.Form("Arancel_GT"))
	rs("Arancel_SV") = PurgeData(Request.Form("Arancel_SV"))
	rs("Arancel_HN") = PurgeData(Request.Form("Arancel_HN"))
	rs("Arancel_NI") = PurgeData(Request.Form("Arancel_NI"))
	rs("Arancel_CR") = PurgeData(Request.Form("Arancel_CR"))
	rs("Arancel_PA") = PurgeData(Request.Form("Arancel_PA"))
	rs("Arancel_BZ") = PurgeData(Request.Form("Arancel_BZ"))
Case 12	'Monedas
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("CurrencyCode") = PurgeData(Request.Form("CurrencyCode"))
	rs("Name") = PurgeData(Request.Form("Name"))
	rs("Xchange") = Request.Form("Xchange")
	rs("Countries") = Request.Form("Countries")
	rs("Symbol") = Request.Form("Symbol")
Case 13	'Rangos
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("Val") = PurgeData(Request.Form("Val"))
Case 14	'Impuestos
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("Tax") = Request.Form("Tax")
	rs("Countries") = Request.Form("Countries")
Case 15	'Arrival
	rs("CreatedTime") = CreatedTime
    rs("Expired") = SetOn(Request.Form("Expired"))
	rs("ArrivalDate") = Request.Form("ArrivalDate")
    rs("MonitorArrivalDate") = Request.Form("ArrivalDate")
	rs("HDepartureDate") = Request.Form("HDepartureDate")
	rs("Cont") = Request.Form("Cont")
	rs("Destinity") = Request.Form("Destinity")
	rs("TotalToPay") = Request.Form("TotalToPay")
	rs("Concept") = Request.Form("Concept")
	rs("FiscalFactory") = Request.Form("FiscalFactory")
	rs("ArrivalAttn") = Request.Form("ArrivalAttn")
	rs("ArrivalFlight") = Request.Form("ArrivalFlight")
	rs("Comment3") = Request.Form("Comment3")
	rs("ManifestNumber") = Request.Form("ManifestNumber")
Case 18
	rs("CreatedTime") = CreatedTime 
	rs("AWBID") = CheckNum(Request.Form("AWBID"))
	rs("ClientID") = CheckNum(Request.Form("CID"))
	rs("Comment") = Request.Form("Comment")
	rs("OperatorID") = Session("OperatorID")
	rs("BLStatus") = CheckNum(Request.Form("BLStatus"))
	rs("BLStatusName") = Request.Form("BLStatusName")	
    rs("DocTyp") = CheckNum(Request.Form("AT"))	
Case 21        
    rs("GuideNumber") = guia99
    rs("GuideActive") = SetActive(Request.Form("Expired"))
    rs("GuideType") = SetActive(Request.Form("Iniciada"))
    rs("GuideCountry") = Session("OperatorCountry")
    rs("GuideCarrierID") = CheckNum(Request.Form("CarrierID"))
    rs("Comentarios") = Request.Form("Comentarios")
    rs("GuideStatus") = SetActive(Request.Form("Estatus"))

    if Action = 1 then
        rs("CreatedUser") = CheckNum(Session("OperatorID"))
        rs("CreatedDate") = CreatedDate 
        rs("CreatedTime") = CreatedTime
    else        
        'rs("CreatedDate") = ConvertDate(rs("CreatedDate"),6)
        'rs("UpdatedDate") = CreatedDate            
        rs("UpdatedUser") = CheckNum(Session("OperatorID"))
        rs("UpdatedTime") = Year(date) & TwoDigits(Month(date)) & TwoDigits(Day(date)) & TwoDigits(Hour(time)) & TwoDigits(Minute(time)  & TwoDigits(Second(time)) )
    end if

Case 22 '2016-02-11
    '                       0           1           2           3       4           5           6       7       8           9       10      11          12      13          14          15          16              17              18      19          20          21              22          23          24                          
    'QuerySelect = "SELECT MedicionID, CreatedDate, CreatedTime, AwbID, AwbNumber, HAwbNumber, AwbType, DateUno, TimeUno, DateDos, TimeDos, DateTres, TimeTres, DateCuatro, TimeCuatro, MedicionUno, MedicionDos, TotNoOfPieces, TotWeight, Destinity, ShipperData, UserInsert, UserUpdate, DateUpdate, Status FROM "

    rs("AwbID") = Request.Form("OID")
    rs("AwbNumber") = Request.Form("AwbNumber")
    rs("HAwbNumber") = Request.Form("HAwbNumber")
    rs("AwbType") = AwbType

    rs("DateUno") = ConvertDate (Request.Form("DateTo1"), 6)    
    rs("TimeUno") = Request.Form("Hrs1") & ":" & Request.Form("Min1") & ":" & Request.Form("Sec1")
    
    if Request.Form("DateTo2") = "" then
        rs("DateDos") = Request.Form("DateTo1")
    else
        rs("DateDos") = ConvertDate (Request.Form("DateTo2"), 6)
    end if

    rs("TimeDos") = Request.Form("Hrs2") & ":" & Request.Form("Min2") & ":" & Request.Form("Sec2")
    
    if Request.Form("DateTo3") = "" then
        rs("DateTres") = Request.Form("DateTo1")
    else
        rs("DateTres") = ConvertDate (Request.Form("DateTo3"), 6)
    end if
        
    rs("TimeTres") = Request.Form("Hrs3") & ":" & Request.Form("Min3") & ":" & Request.Form("Sec3")

    
    if Request.Form("DateTo4") = "" then
        rs("DateCuatro") = Request.Form("DateTo1")
    else
        rs("DateCuatro") = ConvertDate (Request.Form("DateTo4"), 6)
    end if
    
    rs("TimeCuatro") = Request.Form("Hrs4") & ":" & Request.Form("Min4") & ":" & Request.Form("Sec4")
    
    if Request.Form("DateTo5") = "" then
        rs("DateCinco") = Request.Form("DateTo1")
    else
        rs("DateCinco") = ConvertDate (Request.Form("DateTo5"), 6)
    end if

    rs("TimeCinco") = Request.Form("Hrs5") & ":" & Request.Form("Min5") & ":" & Request.Form("Sec5")
    
    
    
    





    if Request.Form("DateTo1") <> "" AND Request.Form("DateTo2") <> "" AND Request.Form("DateTo3") <> "" AND Request.Form("DateTo4") <> "" then 

        dim DifUno, DifDos, hh, mm, ss, dd, tdias1, tdias2 

        tdias1 = (YEAR(Request.Form("DateTo1")) * 365 * 24 * 60 * 60) + (MONTH(Request.Form("DateTo1")) * 30 * 60 * 60) + (DAY(Request.Form("DateTo1")) * 24 * 60 * 60) 
        tdias2 = (YEAR(Request.Form("DateTo2")) * 365 * 24 * 60 * 60) + (MONTH(Request.Form("DateTo2")) * 30 * 60 * 60) + (DAY(Request.Form("DateTo2")) * 24 * 60 * 60) 
    
        DifUno = (Request.Form("Hrs1") * 60 * 60) + (Request.Form("Min1") * 60) + Request.Form("Sec1") 
        DifDos = (Request.Form("Hrs2") * 60 * 60) + (Request.Form("Min2") * 60) + Request.Form("Sec2") 

        'response.write( "tdias1=" & tdias1 & "<br>" )
        'response.write( "tdias2=" & tdias2 & "<br>" )
        'response.write( "DifUno=" & DifUno & "<br>" )
        'response.write( "DifDos=" & DifDos & "<br>" )
        'response.write( ss & "<BR>" )








        ss = (tdias2 + DifDos) - (tdias1 + DifUno) 
        if ss > 0then                
            mm = Int(ss / 60)
            ss = ss - (mm * 60)
            hh = Int(mm / 60)
            mm = mm - (hh * 60)
            dd = Int(hh / 24)
            hh = hh - (dd * 24)        
            'response.write( "dd=" & dd & "<br>" )        
            'response.write( "hh=" & hh & "<br>" )
            'response.write( "mm=" & mm & "<br>" )
            'response.write( "ss=" & ss & "<br>" )             
            if dd > 0 then
                rs("MedicionUno") = dd & " " & "Dia(s) " & TwoDigits(hh) & ":" & TwoDigits(mm) & ":" & TwoDigits(ss) & " hrs."
            else
                rs("MedicionUno") = TwoDigits(hh) & ":" & TwoDigits(mm) & ":" & TwoDigits(ss) & " hrs."
            end if           
        else
            rs("MedicionUno") = "Revise fechas / horas"
        end if















        tdias1 = (YEAR(Request.Form("DateTo3")) * 365 * 24 * 60 * 60) + (MONTH(Request.Form("DateTo3")) * 30 * 60 * 60) + (DAY(Request.Form("DateTo3")) * 24 * 60 * 60) 
        tdias2 = (YEAR(Request.Form("DateTo4")) * 365 * 24 * 60 * 60) + (MONTH(Request.Form("DateTo4")) * 30 * 60 * 60) + (DAY(Request.Form("DateTo4")) * 24 * 60 * 60) 
    
        DifUno = (Request.Form("Hrs3") * 60 * 60) + (Request.Form("Min3") * 60) + Request.Form("Sec3") 
        DifDos = (Request.Form("Hrs4") * 60 * 60) + (Request.Form("Min4") * 60) + Request.Form("Sec4") 

        ss = (tdias2 + DifDos) - (tdias1 + DifUno) 
        if ss > 0then                
            mm = Int(ss / 60)
            ss = ss - (mm * 60)
            hh = Int(mm / 60)
            mm = mm - (hh * 60)
            dd = Int(hh / 24)
            hh = hh - (dd * 24)        
          
            if dd > 0 then
                rs("MedicionDos") = dd & " " & "Dia(s) " & TwoDigits(hh) & ":" & TwoDigits(mm) & ":" & TwoDigits(ss) & " hrs."
            else
                rs("MedicionDos") = TwoDigits(hh) & ":" & TwoDigits(mm) & ":" & TwoDigits(ss) & " hrs."
            end if           
        else
            rs("MedicionDos") = "Revise fechas / horas"
        end if




    else
        rs("MedicionUno") = "Revise fechas / horas"
        rs("MedicionDos") = "Revise fechas / horas"
    end if

    rs("TotNoOfPieces") = Request.Form("TotNoOfPieces")
    rs("TotWeight") = Request.Form("TotWeight")
    rs("Destinity") = Request.Form("Destinity")
    rs("ShipperData") = Request.Form("ShipperData")    
    rs("Status") = 1    

    CreatedDate = Year(date) & "-" & TwoDigits(Month(date)) & "-" & TwoDigits(Day(date))         
    CreatedTime = TwoDigits(Hour(time)) & TwoDigits(Minute(time)  & TwoDigits(Second(time)) )

    if Action = 1 then
        rs("UserInsert") = Session("OperatorID")               
        rs("CreatedDate") = CreatedDate 
        rs("CreatedTime") = CreatedTime
    else        
        rs("UserUpdate") = Session("OperatorID")
    end if
   
End Select

'response.write(" (Updated1)") 

rs.Update

'response.write(" (Updated2)") 


End Sub

Function SaveMaster (Conn, ObjectID)
Dim Address, AddressT, Phone1, Phone1T, Phone2, Phone2T, Attn, AttnT, Region, rrs
Dim AccoutNo, AccountNoT, IATANo, IATANoT, AddressID, PhoneID, AttnID

	Address = PurgeData(Request.Form("Address"))
	AddressT = PurgeData(Request.Form("AddressT"))
	Phone1 = PurgeData(Request.Form("Phone1"))
	Phone1T = PurgeData(Request.Form("Phone1T"))
	Phone2 = PurgeData(Request.Form("Phone2"))
	Phone2T = PurgeData(Request.Form("Phone2T"))
	Attn = PurgeData(Request.Form("Attn"))
	AttnT = PurgeData(Request.Form("AttnT"))
	AccountNo = PurgeData(Request.Form("AccountNo"))
	AccountNoT = PurgeData(Request.Form("AccountNoT"))
	IATANo = PurgeData(Request.Form("IATANo"))
	IATANoT = PurgeData(Request.Form("IATANoT"))
	AddressID = CheckNum(Request.Form("AddressID"))
	PhoneID = CheckNum(Request.Form("PhoneID"))
	AttnID = CheckNum(Request.Form("AttnID"))
	Region = CheckNum(Request.Form("Region"))
	if Region =-1 then Region = 0 end if
	
	if (AddressT="" and Phone1T="") and (Address<>"" or Phone1<>"") then
		Conn.Execute("insert into direcciones (id_cliente, direccion_completa, ""phone number"", id_nivel_geografico) values (" & ObjectID & ", '" & Address & "', '" & Phone1 & "', " & Region & ")")
	else
		if (AddressT<>"" or Phone1T<>"") and (Address<>"" or Phone1<>"") then
			Conn.Execute("update direcciones set direccion_completa='" & Address & "', ""phone number""='" & Phone1 & "', id_nivel_geografico=" & Region & " where id_cliente=" & ObjectID & " and id_direccion=" & AddressID)
		else
			if (AddressT<>"" or Phone1T<>"") and (Address="" and Phone1="") then
				Conn.Execute("delete from direcciones where id_cliente=" & ObjectID & " and id_direccion=" & AddressID)
			end if
		end if
	end if

	if Phone2T = "" and Phone2 <> "" then
		Conn.Execute("insert into cli_telefonos (id_cliente, numero_telefono) values (" & ObjectID & ", '" & Phone2 & "')")
	else
		if Phone2T <> "" and Phone2 <> "" then
			Conn.Execute("update cli_telefonos set numero_telefono='" & Phone2 & "' where id_cliente=" & ObjectID & " and id_telefono=" & PhoneID)
		else
			if Phone2T <> "" and Phone2 = "" then
				Conn.Execute("delete from cli_telefonos where id_cliente=" & ObjectID & " and id_telefono=" & PhoneID)
			end if
		end if
	end if

	if AttnT = "" and Attn <> "" then
		Conn.Execute("insert into contactos (id_cliente, nombres) values (" & ObjectID & ", '" & Attn & "')")
	else
		if AttnT <> "" and Attn <> "" then
			Conn.Execute("update contactos set nombres='" & Attn & "' where id_cliente=" & ObjectID & " and contacto_id=" & AttnID)
		else
			if AttnT <> "" and Attn = "" then
				Conn.Execute("delete from contactos where id_cliente=" & ObjectID & " and contacto_id=" & AttnID)
			end if
		end if
	end if

	if (AccountNoT="" and IATANoT="") and (AccountNo<>"" or IATANo<>"") then
		Conn.Execute("insert into clientes_aereo (id_cliente, no_cuenta, no_iata) values (" & ObjectID & ", '" & AccountNo & "', '" & IATANo & "')")
		'response.write "insert into clientes_aereo (id_cliente, no_cuenta, no_iata) values (" & ObjectID & ", '" & AccountNo & "', '" & IATANo & "')<br>"
	else
		if (AccountNoT<>"" or IATANoT<>"") and (AccountNo<>"" or IATANo<>"") then
			Conn.Execute("update clientes_aereo set no_cuenta='" & AccountNo & "', no_iata='" & IATANo & "' where id_cliente=" & ObjectID)
			'response.write "update clientes_aereo set no_cuenta='" & AccountNo & "', no_iata='" & IATANo & "' where id_cliente=" & ObjectID & "<br>"
		else
			if (AccountNoT<>"" or IATANoT<>"") and (AccountNo="" and IATANo="") then
				Conn.Execute("delete from clientes_aereo where id_cliente=" & ObjectID)
				'response.write "delete from clientes_aereo where id_cliente=" & ObjectID & "<br>"
			end if
		end if
	end if
	
	'response.write "select id_direccion from direcciones where id_cliente=" & ObjectID & " and direccion_completa='"& Address & "' and ""phone number""='" & Phone1 & "' and id_nivel_geografico=" & Region
	Set rrs = Conn.Execute("select id_direccion from direcciones where id_cliente=" & ObjectID & " and direccion_completa='"& Address & "' and ""phone number""='" & Phone1 & "' and id_nivel_geografico=" & Region)
	if Not rrs.EOF then
		SaveMaster = CheckNum(rrs(0))
	else
		SaveMaster = 0
	end if
	CloseOBJ rrs
End Function

Function ConvertDate (Data, Format)
Dim ConvertDay, ConvertMonth, ConvertYear
'response.write Month(Data) & "/" & Day(Data) & "/" & Year(Data) & "<br>"
if Data <> "" then
	 select case Format
	 case 1 'formato dd/mm/yyyy
	   if Day(Data) < 13 then		 		
				 ConvertDate = TwoDigits(Month(Data)) & "/" & TwoDigits(Day(Data)) & "/" & Year(Data)		 		 		
		 else
				 ConvertDate = TwoDigits(Day(Data)) & "/" & TwoDigits(Month(Data)) & "/" & Year(Data)		 		
		 end if
	 case 2 'formato yyyy/mm/dd para strings
		 ConvertDate = Year(Data) & "/" & TwoDigits(Month(Data)) & "/" & TwoDigits(Day(Data))
	 case 3 'formato yyyy/mm/dd para dates
		 if Day(Data) < 13 then
		 		ConvertDate = Year(Data) & "/" & TwoDigits(Month(Data)) & "/" & TwoDigits(Day(Data))
		 else 
				ConvertDate = Year(Data) & "/" & TwoDigits(Day(Data)) & "/" & TwoDigits(Month(Data))				
		 end if	
	 case 4 'formato yyyy-mm-dd para strings
		 'ConvertDate = Year(Data) & "-" & TwoDigits(Month(Data)) & "-" & TwoDigits(Day(Data))
	   	 if Day(Data) < 13 then		 		
				 ConvertDate =Year(Data) & "-" & TwoDigits(Day(Data)) & "-" & TwoDigits(Month(Data))
		 else
				 ConvertDate =Year(Data) & "-" & TwoDigits(Month(Data)) & "-" & TwoDigits(Day(Data))
		 end if
	 case 5 'formato dd/mm/yyyy
		 ConvertDate = TwoDigits(Day(Data)) & "/" & TwoDigits(Month(Data)) & "/" & Year(Data)
     case 6
        ConvertDate = Year(Data) & "-" & TwoDigits(Month(Data)) & "-" & TwoDigits(Day(Data))
	 end select
end if
End Function


Function ConvertDate2 (Data, Format)
Dim ConvertDay, ConvertMonth, ConvertYear
'response.write Month(Data) & "/" & Day(Data) & "/" & Year(Data) & "<br>"
select case Format
case 1 'formato dd/mm/yyyy
	   if Day(Data) < 13 then		 		
				ConvertDate = Day(Data) & "/" & Month(Data) & "/" & Year(Data)		 		
		 else
		 		 ConvertDate = Month(Data) & "/" & Day(Data) & "/" & Year(Data) 
		 end if
case 2 'formato yyyy/mm/dd para strings
		 if Day(Data) < 13 then
		 		ConvertDate = Year(Data) & "/" & Day(Data) & "/" & Month(Data)
		 else 
				ConvertDate = Year(Data) & "/" & Month(Data) & "/" & Day(Data)
		 end if	
case 3 'formato yyyy/mm/dd para dates
		 ConvertDate = Year(Data) & "/" & Month(Data) & "/" & Day(Data)
end select			
End Function


Function CreateSearchQuery(QuerySelect, OptionX, ByRef MoreOptions, oper)
	if OptionX <> "" then
			 if MoreOptions = 0 then
			 		QuerySelect = QuerySelect & " where "
			 else
					QuerySelect = QuerySelect & oper
			 end if			 
			 QuerySelect = QuerySelect & OptionX
			 MoreOptions = 1
	end if
end Function

Function DisplayBirthDate(BirthDate, LN)
Dim i, BirthDay, BirthMonth, BirthYear, HTMLSelect, selected
Dim SelectOption

Dim Matchh, Match, Matches
if isDate(BirthDate) then
	 BirthDate = ConvertDate(BirthDate,1)
	 Set Match = FRegExp("([0-9]*)\/([0-9]*)\/([0-9]*)", BirthDate, "", 1)
	 For Each Matchh In Match
     Set Matches = Match(0)
     BirthDay = CInt(Matches.SubMatches(0))
     BirthMonth = CInt(Matches.SubMatches(1))
     BirthYear = CInt(Matches.SubMatches(2))
   Next
else
		BirthDay = ""
		BirthMonth = ""
		BirthYear = ""
end if
'Desplegando los Dias
HTMLSelect = "<select class=label name=BirthDay><option value='00'>" & TranslateName(LN, "Día") & "</option>"
for i = 1 to 31
		selected = ">"
		if i = BirthDay then
			 selected = " selected>"
		end if
		if i < 10 then
			 SelectOption = "0" & i
		else
			 SelectOption = i				  
		end if
		HTMLSelect = HTMLSelect & "<option value='" & SelectOption & "' " & selected & SelectOption & "</option>"
next
HTMLSelect = HTMLSelect & "</select>"

'Desplegando los Meses
HTMLSelect = HTMLSelect & "<select class=label name=BirthMonth><option value='00'>" & TranslateName(LN, "Mes") & "</option>"
for i = 1 to 12
		selected = ">"
		if i = BirthMonth then
			 selected = " selected>"
		end if
		if i < 10 then
			 SelectOption = "0" & i
		else
			 SelectOption = i				  
		end if
		HTMLSelect = HTMLSelect & "<option value='" & SelectOption & "' " & selected & SelectOption & "</option>"
next
'Desplegando el Año
HTMLSelect = HTMLSelect & "</select>" & _
					  "<INPUT name=BirthYear maxlength=4 TYPE=text size=4 class=label value=" & BirthYear & ">"
DisplayBirthDate = HTMLSelect 
end Function

Function PurgeData(byval AllData)
Dim Data
 Data = replace(AllData,"&","&amp;",1,-1)
 'Data = replace(Data,"?","",1,-1)
 Data = replace(Data,">","&gt;",1,-1)
 Data = replace(Data,"<","&lt;",1,-1)
 Data = replace(Data,chr(13) & chr(10)," ",1,-1)
 'Data = replace(Data,"'","",1,-1)
 'Data = replace(Data,chr(34),"",1,-1)
 'Data = replace(Data,"=","",1,-1)
 'Data = replace(Data,"|","",1,-1)
 'Data = replace(Data,"^","",1,-1)
 'Data = replace(Data,"$","",1,-1)
 'Data = replace(Data,"Ñ","N",1,-1)
 'Data = replace(Data,"ñ","n",1,-1)
 'Data = replace(Data,"À","A",1,-1)
 'Data = replace(Data,"Á","A",1,-1)
 'Data = replace(Data,"Â","A",1,-1)
 'Data = replace(Data,"Ã","A",1,-1)
 'Data = replace(Data,"Ä","A",1,-1)
 'Data = replace(Data,"Å","A",1,-1)
 'Data = replace(Data,"Æ","",1,-1)
 'Data = replace(Data,"Ç","",1,-1)
 'Data = replace(Data,"È","E",1,-1)
 'Data = replace(Data,"É","E",1,-1)
 'Data = replace(Data,"Ê","E",1,-1)
 'Data = replace(Data,"Ë","E",1,-1)
 'Data = replace(Data,"Ì","I",1,-1)
 'Data = replace(Data,"Í","I",1,-1)
 'Data = replace(Data,"Î","I",1,-1)
 'Data = replace(Data,"Ï","I",1,-1)
 'Data = replace(Data,"Ð","",1,-1)
 'Data = replace(Data,"Ñ","N",1,-1)
 'Data = replace(Data,"Ò","O",1,-1)
 'Data = replace(Data,"Ó","O",1,-1)
 'Data = replace(Data,"Ô","O",1,-1)
 'Data = replace(Data,"Õ","O",1,-1)
 'Data = replace(Data,"Ö","O",1,-1)
 'Data = replace(Data,"×","",1,-1)
 'Data = replace(Data,"Ø","",1,-1)
 'Data = replace(Data,"Ù","U",1,-1)
 'Data = replace(Data,"Ú","U",1,-1)
 'Data = replace(Data,"Û","U",1,-1)
 'Data = replace(Data,"Ü","U",1,-1)
 'Data = replace(Data,"Ý","",1,-1)
 'Data = replace(Data,"Þ","",1,-1)
 'Data = replace(Data,"ß","",1,-1)
 'Data = replace(Data,"à","a",1,-1)
 'Data = replace(Data,"á","a",1,-1)
 'Data = replace(Data,"â","a",1,-1)
 'Data = replace(Data,"ã","a",1,-1)
 'Data = replace(Data,"ä","a",1,-1)
 'Data = replace(Data,"å","a",1,-1)
 'Data = replace(Data,"æ","",1,-1)
 'Data = replace(Data,"ç","",1,-1)
 'Data = replace(Data,"è","e",1,-1)
 'Data = replace(Data,"é","e",1,-1)
 'Data = replace(Data,"ê","e",1,-1)
 'Data = replace(Data,"ë","e",1,-1)
 'Data = replace(Data,"ì","i",1,-1)
 'Data = replace(Data,"í","i",1,-1)
 'Data = replace(Data,"î","i",1,-1)
 'Data = replace(Data,"ï","i",1,-1)
 'Data = replace(Data,"ð","",1,-1)
 'Data = replace(Data,"ñ","n",1,-1)
 'Data = replace(Data,"ò","o",1,-1)
 'Data = replace(Data,"ó","o",1,-1)
 'Data = replace(Data,"ô","o",1,-1)
 'Data = replace(Data,"õ","o",1,-1)
 'Data = replace(Data,"ö","o",1,-1)
 'Data = replace(Data,"÷","",1,-1)
 'Data = replace(Data,"ø","",1,-1)
 'Data = replace(Data,"ù","u",1,-1)
 'Data = replace(Data,"ú","u",1,-1)
 'Data = replace(Data,"û","u",1,-1)
 'Data = replace(Data,"ü","u",1,-1)
 'Data = replace(Data,"ý","",1,-1)
 'Data = replace(Data,"þ","",1,-1)
 'Data = replace(Data,"ÿ","",1,-1) 
 PurgeData = Data
End Function

Sub Checking (OL)
	'Revisando Permisos de los Usuarios
	If Not FRegExp(OL, Session("OperatorLevel"), "", 2) Then
		Response.Redirect "redirect.asp?MS=4"
	end if
End Sub

Function TwoDigits(Val)
	if Val <= 9 then
		 TwoDigits = "0" & Val
	else 
		 TwoDigits = Val
	end if
End Function


Function UpdateReference(Val,AwbType)
Dim Consecutive, NewRef, Mt, Yr, Yr2
	'Val = "EA002-11-2007" 'Ejemplo 
	Mt = TwoDigits(Month(Date))
	Yr = Year(Date)
	NewRef = "001"
	If Len(Val) >= 13 then
		Consecutive = (Mid(Val,3,3)*1)+1 'Incrementando el Consecutivo
		
		Yr2 = Mid(Val,10,4)*1 'Obteniendo el anio actual
		if Yr2 < Yr then 'Verificando si el anio ha cambiado, cada anio se resetea el Consecutivo a 1
			Consecutive = 1
		end if
		
		if Consecutive <= 9 then
			NewRef = "00" & Consecutive
		elseif Consecutive <= 99 then
			NewRef = "0" & Consecutive
		else 
			NewRef = Consecutive		
		end if
	end if
	if AwbType = 1 then
		UpdateReference = "EA" & NewRef & "-" & Mt & "-" & Yr
	else
		UpdateReference = "IA" & NewRef & "-" & Mt & "-" & Yr
	end if	
End Function
	
Sub FormatTime (ByRef CreatedDate, ByRef CreatedTime) 
		If Not isDate(CreatedDate) or Not isNumeric(CreatedTime) then
			 CreatedDate = Date 
			 CreatedTime = Time
			 CreatedTime = Hour(CreatedTime) & TwoDigits(Minute(CreatedTime)) & TwoDigits(Second(CreatedTime))
			 'CreatedDate = Year(CreatedDate) & "/" & Month(CreatedDate) & "/" & Day(CreatedDate) 
		end if
		CreatedDate = ConvertDate(CreatedDate,2)
		'CreatedDate = FormatDateTime(Year(CreatedDate) & "/" & Month(CreatedDate) & "/" & day(CreatedDate))
end sub

Function NameOfDay (DayValue)
select case DayValue
case 1
		 NameOfDay = "Domingo"
case 2
		 NameOfDay = "Lunes"
case 3
		 NameOfDay = "Martes"
case 4
		 NameOfDay = "Miercoles"
case 5
		 NameOfDay = "Jueves"
case 6
		 NameOfDay = "Viernes"
case 7
		 NameOfDay = "Sabado"
end select
End Function

Function NameOfMonth (MonthValue)
select case MonthValue
case 1
		 NameOfMonth = "Enero"
case 2
		 NameOfMonth = "Febrero"
case 3
		 NameOfMonth = "Marzo"
case 4
		 NameOfMonth = "Abril"
case 5
		 NameOfMonth = "Mayo"
case 6
		 NameOfMonth = "Junio"
case 7
		 NameOfMonth = "Julio"
case 8
		 NameOfMonth = "Agosto"
case 9
		 NameOfMonth = "Septiembre"
case 10
		 NameOfMonth = "Octubre"
case 11
		 NameOfMonth = "Noviembre"
case 12
		 NameOfMonth = "Diciembre"
end select
End Function

Function NameOfMonth2 (MonthValue)
select case MonthValue
case 1
		 NameOfMonth2 = "ENE"
case 2
		 NameOfMonth2 = "FEB"
case 3
		 NameOfMonth2 = "MAR"
case 4
		 NameOfMonth2 = "ABR"
case 5
		 NameOfMonth2 = "MAY"
case 6
		 NameOfMonth2 = "JUN"
case 7
		 NameOfMonth2 = "JUL"
case 8
		 NameOfMonth2 = "AGO"
case 9
		 NameOfMonth2 = "SEP"
case 10
		 NameOfMonth2 = "OCT"
case 11
		 NameOfMonth2 = "NOV"
case 12
		 NameOfMonth2 = "DIC"
end select
End Function


Sub DisplaySearchAdminResults (HTMLCode) 
Dim Conn, rs, rs2, j, HTMLCode2, HTMLCode3, ListColor, Expired
select case GroupID 
case 7, 8, 10, 11
	OpenConn2 Conn
case else
	OpenConn Conn
end select
	'Buscando los archivos que coinciden con el query de Busqueda
	'response.write QuerySelect  & OrderName & "<br>"
	Set rs = Conn.Execute(QuerySelect & OrderName)
	if Not rs.EOF then
		'Obteniendo la cantidad de resultados por busqueda
		rs.PageSize = Session("SearchResults")
		'Saltando a la pagina seleccionada
		rs.AbsolutePage = AbsolutePage
		PageCount = rs.PageCount
		'Desplegando los resultados de la pagina
        dim nfont
		for i=1 to rs.PageSize
			CD = ConvertDate(rs(2),2)'Day(rs(3)) & "/" & Month(rs(3)) & "/" & Year(rs(3))

            nfont = ""

			for j = 2 to elements
				select case GroupID
				case 4, 6, 15, 18
					if rs(elements) <> "" then
						ListColor = "list"
						Expired = 0
					else
						ListColor = "listwarning"
						Expired = 1
					end if

                        if GroupID = 18 and nfont = "" then '2017-04-11                                                        
                            nfont = "class=" & ListColor
                            if rs(3) <> "" and rs(3) <> rs(6)  then                                
                                Set rs2 = Conn.Execute("SELECT count(*) FROM Tracking WHERE AWBID = " & rs(0))
	                            'if Not rs2.EOF then                                
                                 if rs2(0) > "0" then
                                    'response.write "SELECT count(*) FROM Tracking WHERE AWBID = " & rs(0) & " (" & j & ")(" & rs(3) & ")(" & rs(6) & ")(" & rs2(0) & ")<br>"                                    
                                else
                                    nfont = "style=background:red"
                                end if
                            end if                            
                            'ListColor = ""                            
                            'if rs(8) = "-1" then este campo ya no existe 2017-04-18
                            '    nfont = "style=background:red"
                            'else
                            '    nfont = "style=background:#333333"
                            'end if                                                        
                        end if

				case 7, 8, 10
					if rs(elements) = 1 then
						ListColor = "list"
						Expired = 0
					else
						ListColor = "listwarning"
						Expired = 1
					end if
				case else
					if rs(elements) = 0 then
						ListColor = "list"
						Expired = 0
					else
						ListColor = "listwarning"
						Expired = 1
					end if
				end select

				 Select case j
				 case 2 'cuando es fecha se le da formato español
						'HTMLCode3 = CD & "</a></td>"
                        HTMLCode3 = rs(0) & "</a></td>"                        

				 case 4 'cuando es titulo, nombre o descripcion, se le da formato acotado
						HTMLCode3 = Mid(rs(j),1,40) & "...</a></td>"
				 case elements 'cuando es titulo, nombre o descripcion, se le da formato acotado
						if Expired=0 then
							HTMLCode3 = "Activo</a></td>"
						else
							HTMLCode3 = "Inactivo</a></td>"
						end if
				 case else
						Select case GroupID
						Case 1
							HTMLCode3 = rs(j) & " - " & rs(6) & "</a></td>"
						Case 15, 18
							HTMLCode3 = rs(6) & " - " & rs(j) & "</a></td>"
						Case else
							HTMLCode3 = rs(j) & "</a></td>"
						end select
				end select
				Select case GroupID
				case 3
					HTMLCode2 = HTMLCode2 & "<td class=" & ListColor & "><a class=labellist href=InsertData.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & "&CarrierID=" & rs(7) & "&AirportID=" & rs(8) & "&AT=" & AwbType & ">" & HTMLCode3
				case 5
					HTMLCode2 = HTMLCode2 & "<td class=" & ListColor & "><a class=labellist href=InsertData.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & "&CarrierID=" & rs(6) & "&RangeID=" & rs(7) & "&AT=" & AwbType & ">" & HTMLCode3
				case 7, 10
					HTMLCode2 = HTMLCode2 & "<td class=" & ListColor & "><a class=labellist href=InsertData.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & "&AT=" & AwbType & "&AID=" & rs(7) & ">" & HTMLCode3 
				case 17
					HTMLCode2 = HTMLCode2 & "<td class=" & ListColor & "><a class=labellist href=Costs.asp?OID=" & rs(0) & "&DocType=" & AwbType & ">" & HTMLCode3 
                case 18
					HTMLCode2 = HTMLCode2 & "<td " & nfont & "><a class=labellist href=InsertData.asp?AWBID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & "&AT=" & AwbType & "&CID=" & rs(7) & ">" & HTMLCode3 
                case 1
					HTMLCode2 = HTMLCode2 & "<td class=" & ListColor & "><a class=labellist href=InsertData.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & "&AT=" & AwbType & "&awb_frame2=2>" & HTMLCode3 '2017-12-13 se agrego awb_frame2
				case else
					HTMLCode2 = HTMLCode2 & "<td class=" & ListColor & "><a class=labellist href=InsertData.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & "&AT=" & AwbType & ">" & HTMLCode3 
				End Select
			next
			HTMLCode = HTMLCode & "<tr>" & HTMLCode2 & "</tr>"
			HTMLCode2 = ""
			rs.MoveNext
			If rs.EOF Then Exit For 				
		 next
	else
		JavaMsg = "No Hay Resultados para esta busqueda"
	end if
CloseOBJs rs, Conn
End Sub

Sub DisplayStats () 
Dim Conn, rs, CountYears, QueryFilter
Dim Tot1, Tot2, Tot3, Tot4, Tot5, Tot6, Tot7, Tot8, Tot9, Tot10, Tot11, Tot12, Tot13, Tot14, Tot15, Tot16, Tot17, Tot18, Tot19
Dim HAirFuelSec, HIntermodal, HPickUp, ChWeight, Complete, prevAWBNumber, NetProfit, AdminFee, HasAdminFee, ConvKgLbs
Select Case ReportType
Case 0
	CountYears = YYTo-YYFrom
	if CountYears = 0 then
		QueryFilter = " and a.AWBDate>='" & YYFrom & "-" & TwoDigits(MMFrom) & "-01' and a.AWBDate<'" & SetFilterMonth(MMTo, YYTo) & "' "
	Else
		QueryFilter = " and a.AWBDate>='" & YYFrom & "-" & TwoDigits(MMFrom) & "-01' and a.AWBDate<='" & YYFrom & "-12-31' "
	End If

	OpenConn Conn
	for i=YYFrom to YYTo
		'response.write QuerySelect & QueryFilter & OrderName & "<br>"
		Set rs = Conn.Execute(QuerySelect & QueryFilter & OrderName)
		'Set rs = Conn.Execute("select month(a.CreatedDate), count(a.AWBID), sum(a.TotNoOfPieces*1), sum(a.TotWeightChargeable*1), sum(a.TotChargeWeightPrepaid*1), sum(a.TotChargeWeightPrepaid*1*b.ComisionRate/100), sum(a.TotChargeValuePrepaid*1), sum(a.TotChargeTaxPrepaid*1), sum(a.AnotherChargesAgentPrepaid*1), sum(a.AnotherChargesCarrierPrepaid*1), sum(a.TotPrepaid*1), sum(a.TotChargeWeightCollect*1), sum(a.TotChargeWeightCollect*1*b.ComisionRate/100), sum(a.TotChargeValueCollect*1), sum(a.TotChargeTaxCollect*1), sum(a.AnotherChargesAgentCollect*1), sum(a.AnotherChargesCarrierCollect*1), sum(a.TotCollect*1) from Awbi a, Carriers b where a.CarrierID=b.CarrierID and a.Countries in ('GT','SV','HN','NI','CR','PA','BZ') and a.HAWBNumber='' and a.CreatedDate>='2008-10-01' and a.CreatedDate<'2008-11-01' group by month(a.CreatedDate)")
		if Not rs.EOF then
			do while Not rs.EOF
				if round(rs(0),2) <> 0 then
					HTMLCode = HTMLCode & "<tr><td class=label align=right><a href=# onclick='javascript:DisplayDetailStats("& rs(0) &", " & i & ");'>" & NameOfMonth2(CDbl(rs(0))) & "," & i & "</a></td>" & _
							"<td class=label align=right>" & round(rs(1),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(2),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(3),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(4),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(5),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(6),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(7),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(8),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(9),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(10),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(5)+rs(8),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(11),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(12),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(13),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(14),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(15),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(16),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(17),2) & "</td>" & _
							"<td class=label align=right>" & round(rs(12)+rs(15),2) & "</td></tr>"
							Tot1 = Tot1 + Cdbl(rs(1))
							Tot2 = Tot2 + Cdbl(rs(2))
							Tot3 = Tot3 + Cdbl(rs(3))
							Tot4 = Tot4 + Cdbl(rs(4))
							Tot5 = Tot5 + Cdbl(rs(5))
							Tot6 = Tot6 + Cdbl(rs(6))
							Tot7 = Tot7 + Cdbl(rs(7))
							Tot8 = Tot8 + Cdbl(rs(8))
							Tot9 = Tot9 + Cdbl(rs(9))
							Tot10 = Tot10 + Cdbl(rs(10))
							Tot11 = Tot11 + Cdbl(rs(5)+rs(8))
							Tot12 = Tot12 + Cdbl(rs(11))
							Tot13 = Tot13 + Cdbl(rs(12))
							Tot14 = Tot14 + Cdbl(rs(13))
							Tot15 = Tot15 + Cdbl(rs(14))
							Tot16 = Tot16 + Cdbl(rs(15))
							Tot17 = Tot17 + Cdbl(rs(16))
							Tot18 = Tot18 + Cdbl(rs(17))
							Tot19 = Tot19 + Cdbl(rs(12)+rs(15))
				end if
				rs.MoveNext
			loop
		end if
		CloseOBJ rs
		'Comprobando si es el ultimo ciclo = ultimo anio
		CountYears = i+1

		if CountYears = YYTo then
			QueryFilter = " and a.AWBDate>='" & YYTo & "-01-01' and a.AWBDate<'" & SetFilterMonth(MMTo, YYTo) & "' "
		else
			QueryFilter = " and year(a.AWBDate)=" & i & " "
		end if
	next
	CloseOBJ Conn
	HTMLCode = HTMLCode & "<tr><td class=label align=right><b>TOTALES</b></td>" & _
					"<td class=label align=right><b>" & round(Tot1,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot2,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot3,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot4,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot5,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot6,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot7,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot8,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot9,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot10,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot11,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot12,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot13,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot14,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot15,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot16,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot17,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot18,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot19,2) & "</b></td></tr>"
Case 1
	QueryFilter = " and a.AWBDate>='" & YYFrom & "-" & TwoDigits(MMFrom) & "-01' and a.AWBDate<'" & SetFilterMonth(MMTo, YYTo) & "' "
	'QuerySelect = "select a.AWBID, a.CreatedDate, a.Voyage, a.AWBNumber, a.HAWBNumber, round(sum(a.TotCarrierRate+a.FuelSurcharge+a.SecurityFee),2), round(sum(a.Intermodal),2), round(sum(a.PickUp),2), round(sum(a.TotWeightChargeable)*2.204621,2), 1, CalcAdminFee" & _
	'	QuerySelect & QueryFilter & " and a.HAWBNumber<>'' group by a.AWBNumber union " & _
	'	"select a.AWBID, a.CreatedDate, a.Voyage, a.AWBNumber, a.HAWBNumber, round(sum(a.TotCarrierRate+a.FuelSurcharge+a.SecurityFee),2), round(a.Intermodal,2), round(a.PickUp,2), round(a.TotWeightChargeable*2.204621,2), 0, CalcAdminFee" & _
	'	QuerySelect & QueryFilter & " and a.HAWBNumber='' group by a.AWBNumber order by Voyage, AWBNumber, HAWBNumber Desc"

    'Comentado el 16.03.12
	'QuerySelect = "select a.AWBID, a.AWBDate, a.Voyage, a.AWBNumber, a.HAWBNumber, round(sum(a.TotCarrierRate+a.FuelSurcharge+a.SecurityFee),2), round(sum(a.Intermodal),2), round(sum(a.PickUp),2), round(sum(a.TotWeightChargeable),2), 1, CalcAdminFee" & _
	'	QuerySelect & QueryFilter & " and a.HAWBNumber<>'' group by a.AWBNumber union " & _
	'	"select a.AWBID, a.AWBDate, a.Voyage, a.AWBNumber, a.HAWBNumber, round(sum(a.TotCarrierRate+a.FuelSurcharge+a.SecurityFee),2), round(a.Intermodal,2), round(a.PickUp,2), round(a.TotWeightChargeable,2), 0, CalcAdminFee" & _
	'	QuerySelect & QueryFilter & " and a.HAWBNumber='' group by a.AWBNumber order by AWBDate, Voyage, AWBNumber, HAWBNumber Desc"

	QuerySelect = "select a.AWBID, a.AWBDate, a.Voyage, a.AWBNumber, a.HAWBNumber, round(sum(a.TotCarrierRate+a.FuelSurcharge+a.SecurityFee),2), 0, 0, round(sum(a.TotWeightChargeable),2), 1, CalcAdminFee" & _
		QuerySelect & QueryFilter & " and a.HAWBNumber<>'' group by a.AWBNumber union " & _
		"select a.AWBID, a.AWBDate, a.Voyage, a.AWBNumber, a.HAWBNumber, round(sum(a.TotCarrierRate+a.FuelSurcharge+a.SecurityFee),2), 0, 0, round(a.TotWeightChargeable,2), 0, CalcAdminFee" & _
		QuerySelect & QueryFilter & " and a.HAWBNumber='' group by a.AWBNumber order by AWBDate, Voyage, AWBNumber, HAWBNumber Desc"
    
    'response.write QuerySelect & "<br>"
	
    'Para convertir Kgs a Lbs
    ConvKgLbs = 2.2046
	
    OpenConn Conn
	Set rs = Conn.Execute(QuerySelect)
	do while Not rs.EOF
		select case CInt(rs(9))
		case 0
			if prevAWBNumber=rs(3) then	'Completando el registro con los datos de la Master y calculos
				if rs(10) = 1 then
					'AdminFee = rs(8)*ConvKgLbs*0.04
                    AdminFee = rs(8)*0.09
					if AdminFee < 25 then AdminFee=25 end if
				else
					AdminFee = 0
				end if
				NetProfit = (HAirFuelSec+HIntermodal+HPickUp)-(CDbl(rs(5))+CDbl(rs(6))+CDbl(rs(7))+AdminFee)
				HTMLCode = HTMLCode & "<td class=label align=right>" & round(HAirFuelSec,2) & "</td>" & _
					"<td class=label align=right>" & round(HIntermodal,2) & "</td>" & _
					"<td class=label align=right>" & round(HPickUp,2) & "</td>" & _
					"<td class=label align=right>" & round(rs(5),2) & "</td>" & _
					"<td class=label align=right>" & round(rs(6),2) & "</td>" & _
					"<td class=label align=right>" & round(rs(7),2) & "</td>" & _
					"<td class=label align=right>" & round(rs(8),2) & "</td>" & _
					"<td class=label align=right>" & round(rs(8)*ConvKgLbs,2) & "</td>" & _
					"<td class=label align=right>" & round(AdminFee,2) & "</td>" & _
					"<td class=label align=right>" & round(NetProfit,2) & "</td>" & _
					"<td class=label align=right>" & round(NetProfit/2,2) & "</td></tr>"
				Tot4 = Tot4 + CDbl(rs(5))
				Tot5 = Tot5 + CDbl(rs(6))
				Tot6 = Tot6 + CDbl(rs(7))
				Tot7 = Tot7 + CDbl(rs(8)*ConvKgLbs)
				Tot8 = Tot8 + AdminFee
				Tot9 = Tot9 + NetProfit
			else
				if rs(10) = 1 then
					'AdminFee = ChWeight*ConvKgLbs*0.04
                    AdminFee = ChWeight*0.09
					if AdminFee < 25 then AdminFee=25 end if
				else
					AdminFee = 0
				end if
				if Complete=1 then 'Hay un registro pendiente de completar que no tiene master pero si tiene calculos
					NetProfit = (HAirFuelSec+HIntermodal+HPickUp)-(AdminFee)
					HTMLCode = HTMLCode & "<td class=label align=right>" & round(HAirFuelSec,2) & "</td>" & _
						"<td class=label align=right>" & round(HIntermodal,2) & "</td>" & _
						"<td class=label align=right>" & round(HPickUp,2) & "</td>" & _
						"<td class=label align=right>0.00</td>" & _
						"<td class=label align=right>0.00</td>" & _
						"<td class=label align=right>0.00</td>" & _
						"<td class=label align=right>" & round(ChWeight,2) & "</td>" & _
						"<td class=label align=right>" & round(ChWeight*ConvKgLbs,2) & "</td>" & _
						"<td class=label align=right>" & round(AdminFee,2) & "</td>" & _
						"<td class=label align=right>" & round(NetProfit,2) & "</td>" & _
						"<td class=label align=right>" & round(NetProfit/2,2) & "</td></tr>"
						Tot7 = Tot7 + ChWeight
						Tot8 = Tot8 + AdminFee
						Tot9 = Tot9 + NetProfit
				end if
				NetProfit = -(CDbl(rs(5))+CDbl(rs(6))+CDbl(rs(7)))
				HTMLCode = HTMLCode & "<tr><td class=label align=right>" & rs(1) & "</td>" & _
					"<td class=label align=right>" & rs(2) & "</td>" & _
					"<td class=label align=right>" & rs(3) & "</td>" & _
					"<td class=label align=right>0.00</td>" & _
					"<td class=label align=right>0.00</td>" & _
					"<td class=label align=right>0.00</td>" & _
					"<td class=label align=right>" & round(rs(5),2) & "</td>" & _
					"<td class=label align=right>" & round(rs(6),2) & "</td>" & _
					"<td class=label align=right>" & round(rs(7),2) & "</td>" & _
					"<td class=label align=right>0.00</td>" & _
					"<td class=label align=right>0.00</td>" & _
					"<td class=label align=right>0.00</td>" & _
					"<td class=label align=right>" & round(NetProfit,2) & "</td>" & _
					"<td class=label align=right>" & round(NetProfit/2,2) & "</td></tr>"
					Tot4 = Tot4 + CDbl(rs(5))
					Tot5 = Tot5 + CDbl(rs(6))
					Tot6 = Tot6 + CDbl(rs(7))
					Tot9 = Tot9 + NetProfit
			end if
			Complete = 0
		case 1
			if Complete=1 then 'Hay un registro pendiente de completar que no tiene master pero si tiene calculos
				if rs(10) = 1 then
					'AdminFee = ChWeight*ConvKgLbs*0.04
                    AdminFee = ChWeight*0.09
					if AdminFee < 25 then AdminFee=25 end if
				else
					AdminFee = 0
				end if
				NetProfit = (HAirFuelSec+HIntermodal+HPickUp)-(AdminFee)
				HTMLCode = HTMLCode & "<td class=label align=right>" & round(HAirFuelSec,2) & "</td>" & _
					"<td class=label align=right>" & round(HIntermodal,2) & "</td>" & _
					"<td class=label align=right>" & round(HPickUp,2) & "</td>" & _
					"<td class=label align=right>0.00</td>" & _
					"<td class=label align=right>0.00</td>" & _
					"<td class=label align=right>0.00</td>" & _
					"<td class=label align=right>" & round(ChWeight,2) & "</td>" & _
					"<td class=label align=right>" & round(ChWeight*ConvKgLbs,2) & "</td>" & _
					"<td class=label align=right>" & round(AdminFee,2) & "</td>" & _
					"<td class=label align=right>" & round(NetProfit,2) & "</td>" & _
					"<td class=label align=right>" & round(NetProfit/2,2) & "</td></tr>"
					Tot7 = Tot7 + ChWeight
					Tot8 = Tot8 + AdminFee
					Tot9 = Tot9 + NetProfit
			end if
				
			'HTMLCode = HTMLCode & "<tr><td class=label align=right><a href=# onclick='javascript:DisplayDetailStats(" & rs(0) & ");'>" & rs(1) & "</a></td>" & _
			HTMLCode = HTMLCode & "<tr><td class=label align=right>" & rs(1) & "</td>" & _
				"<td class=label align=right>" & rs(2)& "</td>" & _
				"<td class=label align=right>" & rs(3)& "</td>"
				HAirFuelSec = CDbl(rs(5))
				HIntermodal = CDbl(rs(6))
				HPickUp = CDbl(rs(7))
				ChWeight = CDbl(rs(8))
				HasAdminFee = rs(9)
				Complete = 1
				Tot1 = Tot1 + HAirFuelSec
				Tot2 = Tot2 + HIntermodal
				Tot3 = Tot3 + HPickUp
		end select

		prevAWBNumber=rs(3)
		rs.MoveNext
	loop
	CloseOBJs rs, Conn
	if Complete=1 then 'Hay un registro pendiente de completar que no tiene master pero si tiene calculos
		if HasAdminFee then
			'AdminFee = ChWeight*ConvKgLbs*0.04
            AdminFee = ChWeight*0.09
			if AdminFee < 25 then AdminFee=25 end if
		else
			AdminFee = 0
		end if

		NetProfit = (HAirFuelSec+HIntermodal+HPickUp)-(AdminFee)
		HTMLCode = HTMLCode & "<td class=label align=right>" & round(HAirFuelSec,2) & "</td>" & _
			"<td class=label align=right>" & round(HIntermodal,2) & "</td>" & _
			"<td class=label align=right>" & round(HPickUp,2) & "</td>" & _
			"<td class=label align=right>0.00</td>" & _
			"<td class=label align=right>0.00</td>" & _
			"<td class=label align=right>0.00</td>" & _
			"<td class=label align=right>" & round(ChWeight,2) & "</td>" & _
			"<td class=label align=right>" & round(AdminFee,2) & "</td>" & _
			"<td class=label align=right>" & round(NetProfit,2) & "</td>" & _
			"<td class=label align=right>" & round(NetProfit/2,2) & "</td></tr>"
			Tot7 = Tot7 + ChWeight
			Tot8 = Tot8 + AdminFee
			Tot9 = Tot9 + NetProfit
	end if
	HTMLCode = HTMLCode & "<tr><td class=label align=right colspan=3><b>TOTALES</b></td>" & _
					"<td class=label align=right><b>" & round(Tot1,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot2,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot3,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot4,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot5,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot6,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot7,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot7*ConvKgLbs,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot8,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot9,2) & "</b></td>" & _
					"<td class=label align=right><b>" & round(Tot9/2,2) & "</b></td></tr>"
End Select
End Sub

Sub DisplayCargoStats () 
Dim Conn, i, rs, aTableValues, CountTableValues, DiasTransito, Agent, Shipper, Consignee

    'response.write QuerySelect & OrderName
    CountTableValues=-1 
	OpenConn Conn
        Set rs = Conn.Execute(QuerySelect & OrderName)
        if Not rs.EOF then
            aTableValues = rs.GetRows
			CountTableValues = rs.RecordCount - 1
        end if
    CloseOBJs rs, Conn

	for i=0 to CountTableValues
        'a.AgentData, b.Name, a.HAWBNumber, a.Weights, c.Name, a.HDepartureDate, a.ArrivalDate, a.ShipperData, a.ConsignerData, a.Routing, a.RoutingID
    
    	DiasTransito="&nbsp;" 
        Agent = "&nbsp;"
        Shipper = "&nbsp;"
        Consignee = "&nbsp;"

        if aTableValues(1,i)="" then
            aTableValues(1,i) = "&nbsp;"
        end if
        if aTableValues(2,i)="" then
            aTableValues(2,i) = "&nbsp;"
        end if
        if aTableValues(4,i)="" then
            aTableValues(4,i) = "&nbsp;"
        end if
        if IsNull(aTableValues(5,i)) then
            aTableValues(5,i) = "&nbsp;"
        end if
        if IsNull(aTableValues(6,i)) then
            aTableValues(6,i) = "&nbsp;"
        end if
        if aTableValues(9,i)="" then
            aTableValues(9,i) = "&nbsp;"
        end if
        if aTableValues(10,i)="" then
            aTableValues(10,i) = "&nbsp;"
        end if

        if aTableValues(0,i) <> "" then
            On Error Resume Next    
            'Se hace Split del dato para obtener solo el nombre del Agente        
            Agent = Split(aTableValues(0,i),chr(13))
            If Err.Number <> 0 Then
                response.write  Err.Number & " " & Err.description & "(" & aTableValues(0,i) & ")"
            end if
        end if


        'Si los campos son fechas se calcula su diferencia de tiempo
        if aTableValues(5,i)<>"&nbsp;" and aTableValues(6,i)<>"&nbsp;" then
            DiasTransito = DateDiff("d",ConvertDate(aTableValues(5,i),4),ConvertDate(aTableValues(6,i),4))
        end if
        'Se hace Split del dato para obtener solo el nombre del Shipper
        Shipper = Split(aTableValues(7,i),chr(13))
        'Se hace Split del dato para obtener solo el nombre del Consignatario
        Consignee = Split(aTableValues(8,i),chr(13))

        'Se busca el vendedor del Routing
        OpenConn2 Conn
            Set rs = Conn.Execute("select pw_name from usuarios_empresas, routings where vendedor_id=id_usuario and id_routing=" & aTableValues(10,i))
            if Not rs.EOF then
                aTableValues(10,i) = rs(0)
            end if
        CloseOBJs rs, Conn
        
		HTMLCode = HTMLCode & "<tr>" & _
			"<td class=label align=right>" & Agent(0) & "</td>" & _
			"<td class=label align=right>" & aTableValues(1,i) & "</td>" & _
			"<td class=label align=right>" & aTableValues(2,i) & "</td>" & _
			"<td class=label align=right>" & CheckNum(aTableValues(3,i)) & "</td>" & _
			"<td class=label align=right>" & aTableValues(4,i) & "</td>" & _
			"<td class=label align=right>" & aTableValues(5,i) & "</td>" & _
			"<td class=label align=right>" & aTableValues(6,i) & "</td>" & _
			"<td class=label align=right>" & DiasTransito & "</td>" & _
			"<td class=label align=right>" & Shipper(0) & "</td>" & _
			"<td class=label align=right>" & Consignee(0) & "</td>" & _
			"<td class=label align=right>" & aTableValues(9,i) & "</td>" & _
			"<td class=label align=right>" & aTableValues(10,i) & "</td>" & _
			"</tr>"
	next	
End Sub

Function IIf(i,j,k)
If i = True Then IIf = j Else IIf = k
End Function

Sub DisplayBitacora()
    Dim ConnBaw1, rst, rs_bw, documento, replica, iSQL
    Dim sTotal_Aereo_Quetzales, sTotal_Aereo_USD, sTotal_Baw_SubTotal, sTotal_Baw_Iva, sTotal_Baw_Total, Acumula_Quetzales, Acumula_Dolares
    Acumula_Quetzales = 0
    Acumula_Dolares = 0
    documento = ""
    sTotal_Baw_SubTotal = 0
    sTotal_Baw_Iva = 0
    sTotal_Baw_Total = 0
    i = 0
    openConnBAW ConnBaw1
    '                               0           1               2               3           4           5           6           7           8       9           10              11          12              13          14          15          16          17
    'QuerySelect = "SELECT distinct a.AWBID, a.CreatedDate, b.CreatedTime, a.AwbNumber, a.HAwbNumber, a.UserID, b.DocTyp, b.CurrencyID, b.ItemID, b.ItemName, b.Value, b.PrepaidCollect, b.ServiceID, b.ServiceName, b.InvoiceID, b.ChargeID, b.DocType, c.FirstName, c.LastName FROM Awb a INNER JOIN ChargeItems b ON a.AwbID = b.AWBID AND b.Expired = '0' AND b.DocType IN (1,4) AND b.DocTyp = '0' AND b.InvoiceID > 0 LEFT JOIN Operators c ON c.OperatorID = a.UserID WHERE a.Countries = 'GT' AND YEAR(a.CreatedDate) = 2017 ORDER BY a.AWBID, b.ItemID LIMIT 100"
    'response.write QuerySelect & "<br>"

    OpenConn Conn

    Set rst = Conn.Execute(QuerySelect & OrderName)

    HTMLCode = HTMLCode & "<tbod>"

        On Error Resume Next    


    Do While Not rst.EOF

        'response.write rst(3) & "<br>"

	    if rst(3) <> "" then

            if rst(14) <> documento then

                if rst(6) = 0 then
                    replica = "Export"
                else
                    replica = "Import"                
                end if


                if rst(16) = 1 then 'factura
                    iSQL = "SELECT tfa_serie, tfa_correlativo, tfa_sub_total_eq, tfa_impuesto_eq, tfa_total_eq FROM tbl_facturacion WHERE tfa_id = '" & rst(14) & "' AND tfa_ted_id != '3' LIMIT 1"
                end if                    
                if rst(16) = 4 then 'nota credito
                    iSQL = "SELECT tnc_serie, tnc_correlativo, tnc_monto, 0, tnc_monto FROM tbl_nota_credito WHERE tnc_id = '" & rst(14) & "' AND tnc_ted_id != '3' LIMIT 1"
                end if
                Set rs_bw = ConnBaw1.Execute(iSQL)  
                if Not rst.EOF then
                    if rs_bw(0) <> "" then

                        if Acumula_Quetzales + Acumula_Dolares > 0 then
                            HTMLCode = HTMLCode & "<tr>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td align=right><b>" & IIf(Acumula_Quetzales > 0, FormatNumber(Acumula_Quetzales), "") & "</th>"
                            HTMLCode = HTMLCode & "<td align=right><b>" & IIf(Acumula_Dolares > 0, FormatNumber(Acumula_Dolares), "") & "</th>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "<td></td>"
                            HTMLCode = HTMLCode & "</tr>"    

                            HTMLCode = HTMLCode & "<tr>"    
                            HTMLCode = HTMLCode & "<td colspan=16><hr></td>"
                            HTMLCode = HTMLCode & "</tr>"    

                        end if 

                        i = i + 1

                        HTMLCode = HTMLCode & "<tr>"
                        HTMLCode = HTMLCode & "<td><b>" & i & "</td>"
                        HTMLCode = HTMLCode & "<td>" & rst(1) & "</td>"
                        HTMLCode = HTMLCode & "<td>" & left(rst(2),2) & ":" & mid(rst(2),2,2) & ":" & right(rst(2),2) & "</td>"
                        HTMLCode = HTMLCode & "<td>" & rst(3) & "</td>"
                        HTMLCode = HTMLCode & "<td>" & rst(4) & "</td>"
                        HTMLCode = HTMLCode & "<td>" & replica & "</td>"
                        HTMLCode = HTMLCode & "<td>" & rst(11) & "</td>"                
                        HTMLCode = HTMLCode & "<td>" & rst(17) & " " & rst(18) & "</td>"
                        HTMLCode = HTMLCode & "<td>" & rst(8) & " " & rst(9) & "</td>"

                        HTMLCode = HTMLCode & "<td align=right>" & IIf(rst(7) = "GTQ", FormatNumber(rst(10)), "") & "</td>"
                        HTMLCode = HTMLCode & "<td align=right>" & IIf(rst(7) = "USD", FormatNumber(rst(10)), "") & "</td>"

                        HTMLCode = HTMLCode & "<td rowspan=3 >" & rs_bw(0) & " - " & rs_bw(1) & "</td>"
                        HTMLCode = HTMLCode & "<td rowspan=3 align=right>" & FormatNumber(rs_bw(2)) & "</td>"
                        HTMLCode = HTMLCode & "<td rowspan=3 align=right>" & FormatNumber(rs_bw(3)) & "</td>"
                        HTMLCode = HTMLCode & "<td rowspan=3 align=right>" & FormatNumber(rs_bw(4)) & "</td>"
                        HTMLCode = HTMLCode & "</tr>"    

                        sTotal_Baw_SubTotal = sTotal_Baw_SubTotal + CDBL(rs_bw(2))
                        sTotal_Baw_Iva = sTotal_Baw_Iva + CDBL(rs_bw(3))
                        sTotal_Baw_Total = sTotal_Baw_Total + CDBL(rs_bw(4))

                        Acumula_Quetzales = 0   
                        Acumula_Dolares = 0   
                        
                    end if

                    documento = rst(14)

                end if
                

            else

                HTMLCode = HTMLCode & "<tr>"
                HTMLCode = HTMLCode & "<td></td>"
                HTMLCode = HTMLCode & "<td>" & rst(1) & "</td>"
                HTMLCode = HTMLCode & "<td>" & left(rst(2),2) & ":" & mid(rst(2),2,2) & ":" & right(rst(2),2) & "</td>"
                HTMLCode = HTMLCode & "<td>" & rst(3) & "</td>"
                HTMLCode = HTMLCode & "<td>" & rst(4) & "</td>"
                HTMLCode = HTMLCode & "<td>" & replica & "</td>"
                HTMLCode = HTMLCode & "<td>" & rst(11) & "</td>"
                HTMLCode = HTMLCode & "<td>" & rst(17) & " " & rst(18) & "</td>"
                HTMLCode = HTMLCode & "<td>" & rst(8) & " " & rst(9) & "</td>"
                HTMLCode = HTMLCode & "<td align=right>" & IIf(rst(7) = "GTQ", FormatNumber(rst(10)), "") & "</td>"
                HTMLCode = HTMLCode & "<td align=right>" & IIf(rst(7) = "USD", FormatNumber(rst(10)), "") & "</td>"
                HTMLCode = HTMLCode & "</tr>"

            end if        			

            
            IF rst(7) = "GTQ" then
                sTotal_Aereo_Quetzales = sTotal_Aereo_Quetzales + rst(10)
                Acumula_Quetzales = Acumula_Quetzales + rst(10)
            else
                sTotal_Aereo_USD = sTotal_Aereo_USD + rst(10)
                Acumula_Dolares = Acumula_Dolares + rst(10)
            end if

            

        end if
    
    HTMLCode = HTMLCode & "</tbod>"

	    rst.MoveNext
    Loop

    If Err.Number <> 0 Then
        response.write  Err.Number & " " & Err.description    
    end if

    if Acumula_Quetzales+Acumula_Dolares > 0 then
        HTMLCode = HTMLCode & "<tr>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td align=right><b>" & IIf(Acumula_Quetzales > 0, FormatNumber(Acumula_Quetzales), "") & "</th>"
        HTMLCode = HTMLCode & "<td align=right><b>" & IIf(Acumula_Dolares > 0, FormatNumber(Acumula_Dolares), "") & "</th>"
        HTMLCode = HTMLCode & "<th></th>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<th></th>"
        HTMLCode = HTMLCode & "<th></th>"
        HTMLCode = HTMLCode & "<th></th>"
        HTMLCode = HTMLCode & "</tr>"    

        HTMLCode = HTMLCode & "<tr>"    
        HTMLCode = HTMLCode & "<td colspan=16><hr></td>"
        HTMLCode = HTMLCode & "</tr>"    
    end if 

        HTMLCode = HTMLCode & "<thead>"
        HTMLCode = HTMLCode & "<tr>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<th align=right>" & FormatNumber(sTotal_Aereo_Quetzales) & "</th>"
        HTMLCode = HTMLCode & "<th align=right>" & FormatNumber(sTotal_Aereo_USD) & "</th>"
        HTMLCode = HTMLCode & "<td></td>"
        HTMLCode = HTMLCode & "<th align=right>" & FormatNumber(sTotal_Baw_SubTotal) & "</th>"
        HTMLCode = HTMLCode & "<th align=right>" & FormatNumber(sTotal_Baw_Iva) & "</th>"
        HTMLCode = HTMLCode & "<th align=right>" & FormatNumber(sTotal_Baw_Total) & "</th>"
        HTMLCode = HTMLCode & "</tr>"    
        HTMLCode = HTMLCode & "</thead>"
	
    CloseOBJ Conn

    CloseOBJ ConnBaw1

End Sub


Sub DisplayMediciones (Segmentos) 
Dim Conn, i, rs, aTableValues, CountTableValues, tmp1, tmp2, tmp3, tmp4, tmp5, nTitle, ncore, json

    CountTableValues=-1 
	OpenConn Conn

    'response.write QuerySelect & OrderName

        Set rs = Conn.Execute(QuerySelect & OrderName)
        'response.write QuerySelect & OrderName
        if Not rs.EOF then
            aTableValues = rs.GetRows
			CountTableValues = rs.RecordCount - 1
        end if

    CloseOBJs rs, Conn

	for i=0 to CountTableValues

        tmp1 = aTableValues(7,i) & " " & aTableValues(8,i) 'Mid(aTableValues(8,i), 1, 2) & ":" & Mid(aTableValues(8,i), 4, 2) & ":" & Right(aTableValues(8,i), 2)
        tmp2 = aTableValues(9,i) & " " & aTableValues(10,i) 'Mid(aTableValues(10,i), 1, 2) & ":" & Mid(aTableValues(10,i), 4, 2) & ":" & Right(aTableValues(10,i), 2)
        tmp3 = aTableValues(11,i) & " " & aTableValues(12,i) 'Mid(aTableValues(12,i), 1, 2) & ":" & Mid(aTableValues(12,i), 4, 2) & ":" & Right(aTableValues(12,i), 2)
        tmp4 = aTableValues(13,i) & " " & aTableValues(14,i) 'Mid(aTableValues(14,i), 1, 2) & ":" & Mid(aTableValues(14,i), 4, 2) & ":" & Right(aTableValues(14,i), 2)
        tmp5 = aTableValues(25,i) & " " & aTableValues(26,i) 'Mid(aTableValues(26,i), 1, 2) & ":" & Mid(aTableValues(14,i), 4, 2) & ":" & Right(aTableValues(14,i), 2)

		HTMLCode = HTMLCode & "<tr>" & _			
			"<td>" & aTableValues(4,i) & "</td>" & _
			"<td>" & aTableValues(5,i) & "</td>" & _			
			"<td>" & aTableValues(17,i) & "</td>" & _
			"<td>" & aTableValues(18,i) & "</td>" & _
			"<td>" & aTableValues(19,i) & "</td>" & _
            "<td>" & aTableValues(20,i) & "</td>" 

        If InStr(1,Segmentos, "1") > 0 Then
            HTMLCode = HTMLCode & "<td>" & tmp1 & "</td>"
        End If
        
        If InStr(1,Segmentos, "2") > 0 Then
            HTMLCode = HTMLCode & "<td>" & tmp2 & "</td>"
        End If

        If InStr(1,Segmentos, "3") > 0 Then
            HTMLCode = HTMLCode & "<td>" & tmp3 & "</td>"
        End If

        If InStr(1,Segmentos, "4") > 0 Then
            HTMLCode = HTMLCode & "<td>" & tmp4 & "</td>"
        End If

        If InStr(1,Segmentos, "5") > 0 Then
            HTMLCode = HTMLCode & "<td>" & tmp5 & "</td>"
        End If
    
		HTMLCode = HTMLCode & "<td>" & aTableValues(15,i) & "</td>" & _			
        "<td>" & aTableValues(16,i) & "</td>" & _				
		"</tr>"
	next	

    'HTMLCode = HTMLCode & "</tbody>"
    'response.write ("<" & "script" & " language='JavaScript'> function openmed() { " )
    'response.write (" window.open('MedicionesReporte.asp', 'Seleccionar', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');")
    'response.write ("} <" & "/" & "script" & ">")
    'HTMLCode = HTMLCode & "<a href='#' onclick=openmed()>Excel</a>" 

End Sub




Function SetFilterMonth (MM, YY)
if MM = 12 then
	SetFilterMonth = YY+1 & "-01-01"
else 
	SetFilterMonth = YY & "-" & TwoDigits(MM+1) & "-01"
end if
End Function

Function GetNameUser (UserID, UserType)
Dim Conn, rs, QuerySelect
		select case UserType
		case 0
				 QuerySelect = "select Name from Users where UserID=" & UserID				 
		case 1
				 QuerySelect = "select FirstName, LastName from Operators where OperatorID=" & UserID
		end select
		
		OpenConn Conn
		Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			 select case UserType
			 case 0
				 GetNameUser = rs(0)
		   case 1
				 GetNameUser = rs(0) & " " & rs(1)
			 end select
		end if
		CloseOBJs rs, Conn
End Function

Function FormatClock (TimeVal)
FormatClock = ""
Select Case Len(TimeVal)
Case 5
		 FormatClock = " " & Mid(TimeVal, 1, 1) & ":" & Mid(TimeVal, 2, 2) & ":" & Mid(TimeVal, 4, 2)
Case 6
		 FormatClock = " " & Mid(TimeVal, 1, 2) & ":" & Mid(TimeVal, 3, 2) & ":" & Mid(TimeVal, 5, 2) 
end select
End Function



Function DisplayLogo(Country)
	   select case Country
	   Case "GT", "SV", "HN", "NI", "CR", "PA", "BZ"
		DisplayLogo = "<img src='img/aimar.jpg' border=0 width=237 height=59>"
	   Case "N1"
	   	DisplayLogo = "<img src='img/grh.bmp' border=0>"
	   Case "GTLTF","SVLTF","HNLTF","NILTF","CRLTF","PALTF"
	   	DisplayLogo = "<img src='img/LatinFreight.png' border=0>"
	   Case "GT2", "SV2", "HN2", "NI2", "CR2", "PA2", "BZ2"
	   	DisplayLogo = "<img src='img/craft.bmp' border=0 width=208 height=62>"
	   Case "GTRMR"
	   	DisplayLogo = "<img src='img/LatinFreight.png' border=0>"
	   Case Else
	   	DisplayLogo = "<img src='img/aimar.jpg' border=0 width=237 height=59>"
	   End Select
End Function

Function DisplayCountries(Country)			
Dim MatchCountries, Match, HTML
	Set MatchCountries = FRegExp(PtrnCountries, Session("Countries"),  "", 1)
	For Each Match in MatchCountries
	   select case Match.value
	   Case "'GT'"
	   		HTML = HTML & "<option value='GT'"
			if Country = "GT" then HTML = HTML & " selected" end if
			HTML = HTML & ">GUATEMALA-AIMAR</option>" 
	   Case "'SV'"
	   		HTML = HTML & "<option value='SV'"
			if Country = "SV" then HTML = HTML & " selected" end if
			HTML = HTML & ">EL SALVADOR-AIMAR</option>" 
	   Case "'HN'"
	   		HTML = HTML & "<option value='HN'"
			if Country = "HN" then HTML = HTML & " selected" end if
			HTML = HTML & ">HONDURAS-AIMAR</option>" 
	   Case "'NI'"
	   		HTML = HTML & "<option value='NI'"
			if Country = "NI" then HTML = HTML & " selected" end if
			HTML = HTML & ">NICARAGUA-AIMAR</option>" 
	   Case "'N1'"
	   		HTML = HTML & "<option value='N1'"
			if Country = "N1" then HTML = HTML & " selected" end if
			HTML = HTML & ">NICARAGUA-GRH</option>" 
	   Case "'CR'"
	   		HTML = HTML & "<option value='CR'"
			if Country = "CR" then HTML = HTML & " selected" end if
			HTML = HTML & ">COSTA RICA-AIMAR</option>" 
	   Case "'PA'"
	   		HTML = HTML & "<option value='PA'"
			if Country = "PA" then HTML = HTML & " selected" end if
			HTML = HTML & ">PANAMA-AIMAR</option>" 
	   Case "'BZ'"
	   		HTML = HTML & "<option value='BZ'"
			if Country = "BZ" then HTML = HTML & " selected" end if
			HTML = HTML & ">BELICE-AIMAR</option>" 
	   Case "'GT2'"
	   		HTML = HTML & "<option value='GT2'"
			if Country = "GT2" then HTML = HTML & " selected" end if
			HTML = HTML & ">GUATEMALA-CRAFT</option>" 
	   Case "'SV2'"
	   		HTML = HTML & "<option value='SV2'"
			if Country = "SV2" then HTML = HTML & " selected" end if
			HTML = HTML & ">EL SALVADOR-CRAFT</option>" 
	   Case "'HN2'"
	   		HTML = HTML & "<option value='HN2'"
			if Country = "HN2" then HTML = HTML & " selected" end if
			HTML = HTML & ">HONDURAS-CRAFT</option>" 
	   Case "'NI2'"
	   		HTML = HTML & "<option value='NI2'"
			if Country = "NI2" then HTML = HTML & " selected" end if
			HTML = HTML & ">NICARAGUA-CRAFT</option>" 
	   Case "'CR2'"
	   		HTML = HTML & "<option value='CR2'"
			if Country = "CR2" then HTML = HTML & " selected" end if
			HTML = HTML & ">COSTA RICA-CRAFT</option>" 
	   Case "'PA2'"
	   		HTML = HTML & "<option value='PA2'"
			if Country = "PA2" then HTML = HTML & " selected" end if
			HTML = HTML & ">PANAMA-CRAFT</option>" 
	   Case "'BZ2'"
	   		HTML = HTML & "<option value='BZ2'"
			if Country = "BZ2" then HTML = HTML & " selected" end if
			HTML = HTML & ">BELICE-CRAFT</option>"
       Case "'GTLTF'"
	   		HTML = HTML & "<option value='GTLTF'"
			if Country = "GTLTF" then HTML = HTML & " selected" end if
			HTML = HTML & ">LATIN FREIGHT GT</option>" 
       Case "'SVLTF'"
	   		HTML = HTML & "<option value='SVLTF'"
			if Country = "SVLTF" then HTML = HTML & " selected" end if
			HTML = HTML & ">LATIN FREIGHT SV</option>"
       Case "'HNLTF'"
	   		HTML = HTML & "<option value='HNLTF'"
			if Country = "HNLTF" then HTML = HTML & " selected" end if
			HTML = HTML & ">LATIN FREIGHT HN</option>"
       Case "'NILTF'"
	   		HTML = HTML & "<option value='NILTF'"
			if Country = "NILTF" then HTML = HTML & " selected" end if
			HTML = HTML & ">LATIN FREIGHT NI</option>"
       Case "'CRLTF'"
	   		HTML = HTML & "<option value='CRLTF'"
			if Country = "CRLTF" then HTML = HTML & " selected" end if
			HTML = HTML & ">LATIN FREIGHT CR</option>"
       Case "'PALTF'"
	   		HTML = HTML & "<option value='PALTF'"
			if Country = "PALTF" then HTML = HTML & " selected" end if
			HTML = HTML & ">LATIN FREIGHT PA</option>"
       Case "'CN'"
	   		HTML = HTML & "<option value='CN'"
			if Country = "CN" then HTML = HTML & " selected" end if
			HTML = HTML & ">CHINA</option>"
       Case "'BR'"
	   		HTML = HTML & "<option value='BR'"
			if Country = "BR" then HTML = HTML & " selected" end if
			HTML = HTML & ">BRASIL</option>"
       Case "'GTRMR'"
	   		HTML = HTML & "<option value='GTRMR'"
			if Country = "GTRMR" then HTML = HTML & " selected" end if
			HTML = HTML & ">REIMAR</option>"
       Case "'BE'"
	   		HTML = HTML & "<option value='BE'"
			if Country = "BE" then HTML = HTML & " selected" end if
			HTML = HTML & ">BELGICA</option>"
       Case "'ES'"
	   		HTML = HTML & "<option value='ES'"
			if Country = "ES" then HTML = HTML & " selected" end if
			HTML = HTML & ">ESPAÑA</option>"
       Case "'GTTLA'"
	   		HTML = HTML & "<option value='GTTLA'"
			if Country = "GTTLA" then HTML = HTML & " selected" end if
			HTML = HTML & ">TLA GT</option>" 
       Case "'SVTLA'"
	   		HTML = HTML & "<option value='SVTLA'"
			if Country = "SVTLA" then HTML = HTML & " selected" end if
			HTML = HTML & ">TLA SV</option>"
       Case "'HNTLA'"
	   		HTML = HTML & "<option value='HNTLA'"
			if Country = "HNTLA" then HTML = HTML & " selected" end if
			HTML = HTML & ">TLA HN</option>"
       Case "'NITLA'"
	   		HTML = HTML & "<option value='NITLA'"
			if Country = "NITLA" then HTML = HTML & " selected" end if
			HTML = HTML & ">TLA NI</option>"
       Case "'CRTLA'"
	   		HTML = HTML & "<option value='CRTLA'"
			if Country = "CRTLA" then HTML = HTML & " selected" end if
			HTML = HTML & ">TLA CR</option>"
       Case "'PATLA'"
	   		HTML = HTML & "<option value='PATLA'"
			if Country = "PATLA" then HTML = HTML & " selected" end if
			HTML = HTML & ">TLA PA</option>"
       Case "'BZTLA'"
	   		HTML = HTML & "<option value='BZTLA'"
			if Country = "BZTLA" then HTML = HTML & " selected" end if
			HTML = HTML & ">TLA BZ</option>"
		Case "'GTLGX'"
	   		HTML = HTML & "<option value='GTLGX'"
			if Country = "GTLGX" then HTML = HTML & " selected" end if
			HTML = HTML & ">LOGISTICS GT</option>" 
       Case "'CRLGX'"
	   		HTML = HTML & "<option value='CRLGX'"
			if Country = "CRLGX" then HTML = HTML & " selected" end if
			HTML = HTML & ">LOGISTICS CR</option>"
       Case "'PALGX'"
	   		HTML = HTML & "<option value='PALGX'"
			if Country = "PALGX" then HTML = HTML & " selected" end if
			HTML = HTML & ">LOGISTICS PA</option>"

	   end select
	Next
	response.Write HTML
End Function

Function ListCountries (CountriesChecked, CountriesAssigned)
Dim MatchCountries, Match, GT, SV, HN, NI, CR, PA, BZ, N1, GT2, SV2, HN2, NI2, CR2, PA2, BZ2, GTLTF,SVLTF,HNLTF,NILTF,CRLTF,PALTF,CN,BR,GTRMR,BE,ES,GTTLA,SVTLA,HNTLA,NITLA,CRTLA,PATLA,BZTLA,GTLGX,CRLGX,PALGX
Dim HTML, TRTDB, BTDTDI, I1, I2, TDTR	
	Set MatchCountries = FRegExp(PtrnCountries, CountriesChecked, "", 1)
	For Each Match in MatchCountries
	   select case Match.value
	   Case "'GT'"
	   		GT = "checked"			
	   Case "'SV'"
	   		SV = "checked"
	   Case "'HN'"
	   		HN = "checked"
	   Case "'NI'"
	   		NI = "checked"
	   Case "'CR'"
	   		CR = "checked"
	   Case "'PA'"
	   		PA = "checked"
	   Case "'BZ'"
	   		BZ = "checked"
	   Case "'N1'"
	   		N1 = "checked"
	   Case "'GT2'"
	   		GT2 = "checked"			
	   Case "'SV2'"
	   		SV2 = "checked"
	   Case "'HN2'"
	   		HN2 = "checked"
	   Case "'NI2'"
	   		NI2 = "checked"
	   Case "'CR2'"
	   		CR2 = "checked"
	   Case "'PA2'"
	   		PA2 = "checked"
	   Case "'BZ2'"
	   		BZ2 = "checked"
       Case "'GTLTF'"
            GTLTF = "checked"
       Case "'SVLTF'"
	   		SVLTF = "checked"
       Case "'HNLTF'"
	   		HNLTF = "checked"
       Case "'NILTF'"
	   		NILTF = "checked"
       Case "'CRLTF'"
	   		CRLTF = "checked"
       Case "'PALTF'"
	   		PALTF = "checked"
       Case "'CN'"
	   		CN = "checked"
       Case "'BR'"
	   		BR = "checked"
       Case "'GTRMR'"
	   		GTRMR = "checked"
       Case "'BE'"
	   		BE = "checked"
       Case "'ES'"
	   		ES = "checked"
	   Case "'GTTLA'"
            GTTLA = "checked"
       Case "'SVTLA'"
	   		SVTLA = "checked"
       Case "'HNTLA'"
	   		HNTLA = "checked"
       Case "'NITLA'"
	   		NITLA = "checked"
       Case "'CRTLA'"
	   		CRTLA = "checked"
       Case "'PATLA'"
	   		PATLA = "checked"
       Case "'BZTLA'"
	   		BZTLA = "checked"

	   Case "'GTLGX'"
            GTLGX = "checked"
	   Case "'CRLGX'"
	   		CRLGX = "checked"
       Case "'PALGX'"
	   		PALGX = "checked"
        end select
   Next
	   
   TRTDB = "<TR><TD class=label align=right><b>"
   BTDTDI = "</b></TD><TD class=label align=left><INPUT name='"
   I1 = "' value='"
   I2 = "' type=checkbox class=label "
   TDTR = "></TD></TR>"
   HTML = "<TABLE align=center>"	   
   Set MatchCountries = FRegExp(PtrnCountries, CountriesAssigned, "", 1)
   For Each Match in MatchCountries
	   select case Match.value
	   Case "'GT'"
			HTML = HTML & TRTDB & "Guatemala-Aimar:" & BTDTDI & "GT" & I1 & "GT" & I2 & GT & TDTR
	   Case "'SV'"
			HTML = HTML & TRTDB & "El Salvador-Aimar:" & BTDTDI & "SV" & I1 & "SV" & I2 & SV & TDTR
	   Case "'HN'"
			HTML = HTML & TRTDB & "Honduras-Aimar:" & BTDTDI & "HN" & I1 & "HN" & I2 &  HN & TDTR
	   Case "'NI'"
			HTML = HTML & TRTDB & "Nicaragua-Aimar:" & BTDTDI & "NI" & I1 & "NI" & I2 & NI & TDTR
	   Case "'CR'"
			HTML = HTML & TRTDB & "Costa Rica-Aimar:" & BTDTDI & "CR" & I1 & "CR" & I2 & CR & TDTR
	   Case "'PA'"
			HTML = HTML & TRTDB & "Panama-Aimar:" & BTDTDI & "PA" & I1 & "PA" & I2 & PA & TDTR
	   Case "'BZ'"
			HTML = HTML & TRTDB & "Belice-Aimar:" & BTDTDI & "BZ" & I1 & "BZ" & I2 & BZ & TDTR
	   Case "'N1'"
			HTML = HTML & TRTDB & "Nicaragua-GRH:" & BTDTDI & "N1" & I1 & "N1" & I2 & N1 & TDTR
	   Case "'GT2'"
			HTML = HTML & TRTDB & "Guatemala-Craft:" & BTDTDI & "GT2" & I1 & "GT2" & I2 & GT2 & TDTR
	   Case "'SV2'"
			HTML = HTML & TRTDB & "El Salvador-Craft:" & BTDTDI & "SV2" & I1 & "SV2" & I2 & SV2 & TDTR
	   Case "'HN2'"
			HTML = HTML & TRTDB & "Honduras-Craft:" & BTDTDI & "HN2" & I1 & "HN2" & I2 & HN2 & TDTR
	   Case "'NI2'"
			HTML = HTML & TRTDB & "Nicaragua-Craft:" & BTDTDI & "NI2" & I1 & "NI2" & I2 & NI2 & TDTR
	   Case "'CR2'"
			HTML = HTML & TRTDB & "Costa Rica-Craft:" & BTDTDI & "CR2" & I1 & "CR2" & I2 & CR2 & TDTR
	   Case "'PA2'"
			HTML = HTML & TRTDB & "Panama-Craft:" & BTDTDI & "PA2" & I1 & "PA2" & I2 & PA2 & TDTR
	   Case "'BZ2'"
			HTML = HTML & TRTDB & "Belice-Craft:" & BTDTDI & "BZ2" & I1 & "BZ2" & I2 & BZ2 & TDTR
       Case "'GTLTF'"
			HTML = HTML & TRTDB & "Latin Freight GT:" & BTDTDI & "GTLTF" & I1 & "GTLTF" & I2 & GTLTF & TDTR            
       Case "'SVLTF'"
			HTML = HTML & TRTDB & "Latin Freight-SV:" & BTDTDI & "SVLTF" & I1 & "SVLTF" & I2 & SVLTF & TDTR  
       Case "'HNLTF'"
			HTML = HTML & TRTDB & "Latin Freight-HN:" & BTDTDI & "HNLTF" & I1 & "HNLTF" & I2 & HNLTF & TDTR  
       Case "'NILTF'"
			HTML = HTML & TRTDB & "Latin Freight-NI:" & BTDTDI & "NILTF" & I1 & "NILTF" & I2 & NILTF & TDTR  
       Case "'CRLTF'"
			HTML = HTML & TRTDB & "Latin Freight-CR:" & BTDTDI & "CRLTF" & I1 & "CRLTF" & I2 & CRLTF & TDTR  
       Case "'PALTF'"
			HTML = HTML & TRTDB & "Latin Freight-PA:" & BTDTDI & "PALTF" & I1 & "PALTF" & I2 & PALTF & TDTR  
       Case "'CN'"
			HTML = HTML & TRTDB & "CHINA:" & BTDTDI & "CN" & I1 & "CN" & I2 & CN & TDTR  
       Case "'BR'"
			HTML = HTML & TRTDB & "BRASIL:" & BTDTDI & "BR" & I1 & "BR" & I2 & BR & TDTR  
       Case "'GTRMR'"
			HTML = HTML & TRTDB & "REIMAR GT:" & BTDTDI & "GTRMR" & I1 & "GTRMR" & I2 & GTRMR & TDTR  
       Case "'BE'"
			HTML = HTML & TRTDB & "BELGICA:" & BTDTDI & "BE" & I1 & "BE" & I2 & BR & TDTR  
       Case "'ES'"
			HTML = HTML & TRTDB & "ESPAÑA:" & BTDTDI & "ES" & I1 & "ES" & I2 & ES & TDTR  
	   Case "'GTTLA'"
			HTML = HTML & TRTDB & "TLA GT:" & BTDTDI & "GTTLA " & I1 & "GTTLA " & I2 & GTTLA  & TDTR            
       Case "'SVTLA'"
			HTML = HTML & TRTDB & "TLA SV:" & BTDTDI & "SVTLA " & I1 & "SVTLA " & I2 & SVTLA  & TDTR  
       Case "'HNTLA'"
			HTML = HTML & TRTDB & "TLA HN:" & BTDTDI & "HNTLA " & I1 & "HNTLA " & I2 & HNTLA  & TDTR  
       Case "'NITLA'"
			HTML = HTML & TRTDB & "TLA NI:" & BTDTDI & "NITLA " & I1 & "NITLA " & I2 & NITLA  & TDTR  
       Case "'CRTLA'"
			HTML = HTML & TRTDB & "TLA CR:" & BTDTDI & "CRTLA " & I1 & "CRTLA " & I2 & CRTLA  & TDTR  
       Case "'PATLA'"
			HTML = HTML & TRTDB & "TLA PA:" & BTDTDI & "PATLA " & I1 & "PATLA " & I2 & PATLA  & TDTR 			
		Case "'BZTLA'"
			HTML = HTML & TRTDB & "TLA BZ:" & BTDTDI & "BZTLA " & I1 & "BZTLA " & I2 & BZTLA  & TDTR 
       Case "'GTLGX'"
			HTML = HTML & TRTDB & "LOGISTICS GT:" & BTDTDI & "GTLGX " & I1 & "GTLGX " & I2 & GTLGX  & TDTR            
	   Case "'CRLGX'"
			HTML = HTML & TRTDB & "LOGISTICS CR:" & BTDTDI & "CRLGX " & I1 & "CRLGX " & I2 & CRLGX  & TDTR  
       Case "'PALGX'"
			HTML = HTML & TRTDB & "LOGISTICS PA:" & BTDTDI & "PALGX " & I1 & "PALGX " & I2 & PALGX  & TDTR 			

	   end select
	Next
   response.Write HTML & "</TABLE>"
end Function

Function TranslateCompany(Country)
   select case Country
   Case "GT"
		TranslateCompany = "Aimar GT"
   Case "SV"
		TranslateCompany = "Aimar SV"
   Case "HN"
		TranslateCompany = "Aimar HN"
   Case "NI"
		TranslateCompany = "Aimar NI"
   Case "CR"
		TranslateCompany = "Aimar CR"
   Case "PA"
		TranslateCompany = "Aimar PA"
   Case "BZ"
		TranslateCompany = "Aimar BZ"
   Case "N1"
		TranslateCompany = "GRH NI"
   Case "GT2"
		TranslateCompany = "Craft GT"
   Case "SV2"
		TranslateCompany = "Craft SV"
   Case "HN2"
		TranslateCompany = "Craft HN"
   Case "NI2"
		TranslateCompany = "Craft NI"
   Case "CR2"
		TranslateCompany = "Craft CR"
   Case "PA2"
		TranslateCompany = "Craft PA"
   Case "BZ2"
		TranslateCompany = "Craft BZ"
   Case "GTLTF"
		TranslateCompany = "Latin Freight GT"
   Case "SVLTF"
		TranslateCompany = "Latin Freight SV"
   Case "HNLTF"
		TranslateCompany = "Latin Freight HN"
   Case "NILTF"
		TranslateCompany = "Latin Freight NI"
   Case "CRLTF"
		TranslateCompany = "Latin Freight CR"
   Case "PALTF"
		TranslateCompany = "Latin Freight PA"
   Case "CN"
		TranslateCompany = "CHINA"
   Case "BR"
		TranslateCompany = "BRASIL"
   Case "GTRMR"
		TranslateCompany = "REIMAR GT"
   Case "GTTLA"
		TranslateCompany = "TLA GT"
   Case "SVTLA"
		TranslateCompany = "TLA SV"
   Case "HNTLA"
		TranslateCompany = "TLA HN"
   Case "NITLA"
		TranslateCompany = "TLA NI"
   Case "CRTLA"
		TranslateCompany = "TLA CR"
   Case "PATLA"
		TranslateCompany = "TLA PA"
   Case "BZTLA"
		TranslateCompany = "TLA BZ"
				
   Case "GTLGX"
		TranslateCompany = "LGX GT"
   Case "CRLGX"
		TranslateCompany = "LGX CR"
   Case "PALGX"
		TranslateCompany = "LGX PA"

   end select
End Function

Function TranslateCountry (Country)
Select Case Country
	Case "AF" TranslateCountry = "AFGANISTAN"
	Case "AL" TranslateCountry = "ALBANIA"
	Case "DZ" TranslateCountry = "ARGELIA"
	Case "AD" TranslateCountry = "ANDORRA"
	Case "AO" TranslateCountry = "ANGOLA"
	Case "AQ" TranslateCountry = "ANTÁRTIDA"
	Case "AG" TranslateCountry = "ANTIGUA Y BARBUDA"
	Case "AR" TranslateCountry = "ARGENTINA"
	Case "AM" TranslateCountry = "ARMENIA"
	Case "AW" TranslateCountry = "ARUBA"
	Case "AU" TranslateCountry = "AUSTRALIA"
	Case "AT" TranslateCountry = "AUSTRIA"
	Case "AZ" TranslateCountry = "AZERBAIJÁN"
	Case "BS" TranslateCountry = "BAHAMAS"
	Case "BH" TranslateCountry = "BAHRAYN"
	Case "BD" TranslateCountry = "BANGLADESH"
	Case "BB" TranslateCountry = "BARBADOS ISLAS DE BARLOVENTO"
	Case "BY" TranslateCountry = "BELARUS"
	Case "BE" TranslateCountry = "BELGICA"
	Case "BZ" TranslateCountry = "BELICE"
	Case "BJ" TranslateCountry = "BENÍN"
	Case "BM" TranslateCountry = "BERMUDAS"
	Case "BT" TranslateCountry = "BUTÁN"
	Case "BO" TranslateCountry = "BOLIVIA"
	Case "BA" TranslateCountry = "BOSNIA Y HERZEGOVINA"
	Case "BW" TranslateCountry = "BOTSWANA ESTADO DE  AFRICA AUSTRAL"
	Case "BV" TranslateCountry = "BOUVET ISLA NORUETA DEL ATLANTICO SUR"
	Case "BR" TranslateCountry = "BRASIL"
	Case "IO" TranslateCountry = "INDIAS BRITANICAS TERRITORIO DEL OCENO INDICO"
	Case "BN" TranslateCountry = "BRUNEI DARUSSALAM"
	Case "BG" TranslateCountry = "BULGARIA"
	Case "BF" TranslateCountry = "BURKINA FASO"
	Case "BI" TranslateCountry = "BURUNDI"
	Case "KH" TranslateCountry = "CAMBOYA"
	Case "CM" TranslateCountry = "CAMERÚN"
	Case "CA" TranslateCountry = "CANADÁ"
	Case "CV" TranslateCountry = "CAPE VERDE"
	Case "KY" TranslateCountry = "ISLAS CAYMAN"
	Case "CF" TranslateCountry = "REPÚBLICA DE AFRICA CENTRAL"
	Case "TD" TranslateCountry = "CHAD"
	Case "CL" TranslateCountry = "CHILE"
	Case "CN" TranslateCountry = "CHINA"
	Case "CX" TranslateCountry = "ISLA DE NAVIDAD"
	Case "CC" TranslateCountry = "ISLAS DE LOS COCOS"
	Case "CO" TranslateCountry = "COLOMBIA"
	Case "KM" TranslateCountry = "COMORES"
	Case "CG" TranslateCountry = "CONGO"
	Case "CD" TranslateCountry = "REPÚBLICA DEMOCRÁTICA DEL CONGO"
	Case "CK" TranslateCountry = "ISLAS DE COOK"
	Case "CR" TranslateCountry = "COSTA RICA"
	Case "CI" TranslateCountry = "CÔTE D'IVOIRE"
	Case "HR" TranslateCountry = "CROACIA"
	Case "CU" TranslateCountry = "CUBA"
	Case "CY" TranslateCountry = "CHIPRE"
	Case "CZ" TranslateCountry = "REPÚBLICA CHECA"
	Case "DK" TranslateCountry = "DINAMARCA"
	Case "DJ" TranslateCountry = "DJIBOUTI"
	Case "DM" TranslateCountry = "DOMINICA"
	Case "DO" TranslateCountry = "REPÚBLICA DOMINICANA"
	Case "EC" TranslateCountry = "ECUADOR"
	Case "EG" TranslateCountry = "EGYPTO"
	Case "SV" TranslateCountry = "EL SALVADOR"
	Case "VA" TranslateCountry = "EL VATICANO"
	Case "GQ" TranslateCountry = "GUINEA ECUATORIAL"
	Case "ER" TranslateCountry = "ERITREA"
	Case "EE" TranslateCountry = "ESTONIA"
	Case "ET" TranslateCountry = "ETIOPÍA"
	Case "FK" TranslateCountry = "ISLAS DE FALKAND MALVINAS"
	Case "FO" TranslateCountry = "ISLAS DE FAROE"
	Case "FJ" TranslateCountry = "FIDJI"
	Case "FI" TranslateCountry = "FINLANDIA"
	Case "FR" TranslateCountry = "FRANCIA"
	Case "GF" TranslateCountry = "GUINEA FRANCESA"
	Case "PF" TranslateCountry = "POLINESIA FRANCESA"
	Case "GA" TranslateCountry = "GABÓN"
	Case "GM" TranslateCountry = "GAMBIA"
	Case "GE" TranslateCountry = "GEORGIA"
	Case "DE" TranslateCountry = "GERMANIA"
	Case "GH" TranslateCountry = "GHÁNA"
	Case "GI" TranslateCountry = "GIBRALTAR"
	Case "GR" TranslateCountry = "GRECIA"
	Case "GL" TranslateCountry = "GROENLANDIA"
	Case "GD" TranslateCountry = "GRANADA"
	Case "GP" TranslateCountry = "GUADALUPE"
	Case "GU" TranslateCountry = "GUAM"
	Case "GT" TranslateCountry = "GUATEMALA"
	Case "GN" TranslateCountry = "GUINEA"
	Case "GW" TranslateCountry = "GUINEA PORTUGUESA"
	Case "GY" TranslateCountry = "GUYANA"
	Case "HT" TranslateCountry = "HAITÍ"
	Case "HN" TranslateCountry = "HONDURAS"
	Case "HK" TranslateCountry = "HONG KONG"
	Case "HU" TranslateCountry = "HUNGRIA"
	Case "IS" TranslateCountry = "ISLANDIA"
	Case "IN" TranslateCountry = "INDIA"
	Case "ID" TranslateCountry = "INDONESIA"
	Case "IR" TranslateCountry = "IRAN"
	Case "IQ" TranslateCountry = "IRAQ"
	Case "IE" TranslateCountry = "IRLANDIA"
	Case "IL" TranslateCountry = "ISRAEL"
	Case "IT" TranslateCountry = "ITALIA"
	Case "JM" TranslateCountry = "JAMAICA"
	Case "JP" TranslateCountry = "JAPÓN"
	Case "JO" TranslateCountry = "JORDANIA"
	Case "KZ" TranslateCountry = "KASAJISTÁN"
	Case "KE" TranslateCountry = "KENYA"
	Case "KI" TranslateCountry = "KIRIBATI"
	Case "KP" TranslateCountry = "REPÚBLICAS DEMOCRÁTICAS DE COREA"
	Case "KR" TranslateCountry = "REPÚBLICA DE COREA"
	Case "KW" TranslateCountry = "KUWAIT"
	Case "KG" TranslateCountry = "KIRGUIZISTÁN"
	Case "LA" TranslateCountry = "LAOS"
	Case "LV" TranslateCountry = "ESTADO RUSO DE LATVIA"
	Case "LB" TranslateCountry = "LÍBANO"
	Case "LS" TranslateCountry = "LESOTHO O BASUTOLANDIA"
	Case "LR" TranslateCountry = "LIBERIA"
	Case "LY" TranslateCountry = "LIBIA ARABE JAMAHIRYA"
	Case "LI" TranslateCountry = "LIECHTENSTEIN"
	Case "LT" TranslateCountry = "LITUNIA"
	Case "LU" TranslateCountry = "LUXEMBURGO"
	Case "MO" TranslateCountry = "MACAO"
	Case "MK" TranslateCountry = "MACEDONIA, ANTIGUA REPÚBLICA DE YUGOESLAVIA"
	Case "MG" TranslateCountry = "MADAGASCAR"
	Case "MW" TranslateCountry = "MALAWI"
	Case "MY" TranslateCountry = "MALASYA"
	Case "MV" TranslateCountry = "MALDIVAS"
	Case "ML" TranslateCountry = "MALÍ"
	Case "MT" TranslateCountry = "MALTA"
	Case "MH" TranslateCountry = "ISLAS MARSHALL"
	Case "MQ" TranslateCountry = "ISLA DE MARTINICA"
	Case "MR" TranslateCountry = "MAURITANIA"
	Case "MU" TranslateCountry = "ISLA MAURICIO"
	Case "YT" TranslateCountry = "MAYOTTÉ"
	Case "MX" TranslateCountry = "MÉXICO"
	Case "MD" TranslateCountry = "REPÚBLICA DE MOLDOVIA"
	Case "MC" TranslateCountry = "MÓNACO"
	Case "MN" TranslateCountry = "MONGOLIA"
	Case "MS" TranslateCountry = "MONTSERRAT"
	Case "MA" TranslateCountry = "MARRUECOS"
	Case "MZ" TranslateCountry = "MOZAMBIQUE"
	Case "MM" TranslateCountry = "BIRMANIA"
	Case "NA" TranslateCountry = "NAMIBIA"
	Case "NR" TranslateCountry = "NAURU"
	Case "NP" TranslateCountry = "NEPAL"
	Case "NL" TranslateCountry = "PAISES BAJOS"
	Case "AN" TranslateCountry = "ANTILLAS DE LOS PAISES BAJOS"
	Case "NC" TranslateCountry = "NUEVA CALEDONIA"
	Case "NZ" TranslateCountry = "NUEVA ZELANDA"
	Case "NI" TranslateCountry = "NICARAGUA"
	Case "N1" TranslateCountry = "NICARAGUA (GRH)"
	Case "NE" TranslateCountry = "NÍGER"
	Case "NG" TranslateCountry = "NIGERIA"
	Case "NU" TranslateCountry = "SAVAGE ISLA DEL PACÍFICO"
	Case "NF" TranslateCountry = "ISLA NORFOLK"
	Case "MP" TranslateCountry = "ISLAS MARIANAS DEL NORTE"
	Case "NO" TranslateCountry = "NORUEGA"
	Case "OM" TranslateCountry = "OMÁN"
	Case "PK" TranslateCountry = "PAKISTÁN"
	Case "PW" TranslateCountry = "PALAOS"
	Case "PS" TranslateCountry = "PALESTINA"
	Case "PA" TranslateCountry = "PANAMA"
	Case "PG" TranslateCountry = "NUEVA GUINEA - PAPUASIA"
	Case "PY" TranslateCountry = "PARAGUAY"
	Case "PE" TranslateCountry = "PERU"
	Case "PH" TranslateCountry = "FILIPINAS"
	Case "PN" TranslateCountry = "ISLA PITCAIRN"
	Case "PL" TranslateCountry = "POLONIA"
	Case "PT" TranslateCountry = "PORTUGAL"
	Case "PR" TranslateCountry = "PUERTO RICO"
	Case "QA" TranslateCountry = "QATAR"
	Case "RE" TranslateCountry = "RÉUNION"
	Case "RO" TranslateCountry = "RUMANIA"
	Case "RU" TranslateCountry = "FEDERACIÓN RUSIA"
	Case "RW" TranslateCountry = "RUANDA"
	Case "SH" TranslateCountry = "SANTA HELENA"
	Case "KN" TranslateCountry = "SANTA KITTS Y NEVIS "
	Case "LC" TranslateCountry = "SANTA LUCIA"
	Case "PM" TranslateCountry = "SANTO PIER Y MIKELON"
	Case "VC" TranslateCountry = "SAN VICENTE Y LAS GRANADINAS"
	Case "WS" TranslateCountry = "SAMOA"
	Case "SM" TranslateCountry = "SAN MARINO"
	Case "ST" TranslateCountry = "TOMO DE SAO Y PRÍNCIPE"
	Case "SA" TranslateCountry = "ARABIA SAUDITA"
	Case "SN" TranslateCountry = "SENEGAL"
	Case "CS" TranslateCountry = "SERBIA Y MONTENEGRO"
	Case "SC" TranslateCountry = "SEYCHELLES"
	Case "SL" TranslateCountry = "SIERRA LEONA"
	Case "SG" TranslateCountry = "SINGAPUR"
	Case "SK" TranslateCountry = "ESLOVAQUIA"
	Case "SI" TranslateCountry = "ESLOVENIA"
	Case "SB" TranslateCountry = "ISLAS SALOMON"
	Case "SO" TranslateCountry = "SOMALIA"
	Case "ZA" TranslateCountry = "SUR AFRICA"
	Case "GS" TranslateCountry = "GEORGIA DEL SUR Y LAS ISLAS DEL SUR DE SANDWICH"
	Case "ES" TranslateCountry = "ESPAÑA"
	Case "LK" TranslateCountry = "SRI LANKA"
	Case "SD" TranslateCountry = "SUDAN"
	Case "SR" TranslateCountry = "SURINAM"
	Case "SJ" TranslateCountry = "SVALBARD Y ENERO MAYEN"
	Case "SZ" TranslateCountry = "SWAZILANDIA"
	Case "SE" TranslateCountry = "SUECIA"
	Case "CH" TranslateCountry = "SUIZA"
	Case "SY" TranslateCountry = "REPÚBLICA ARABE DE SIRIA"
	Case "TW" TranslateCountry = "TAIWAN PROVINCIA DE CHINA"
	Case "TJ" TranslateCountry = "TAJIKISTAN"
	Case "TZ" TranslateCountry = "REPÚBLICA UNIDA DE TANZANIA"
	Case "TH" TranslateCountry = "TAILANDIA"
	Case "TL" TranslateCountry = "TIMOR-LESTE"
	Case "TG" TranslateCountry = "TOGO"
	Case "TK" TranslateCountry = "TOKELAU"
	Case "TO" TranslateCountry = "TONGA"
	Case "TT" TranslateCountry = "TRINIDAD Y TOBAGO"
	Case "TN" TranslateCountry = "TUNISIA"
	Case "TR" TranslateCountry = "TURKESTÁN"
	Case "TM" TranslateCountry = "TURKMENISTÁN"
	Case "TC" TranslateCountry = "ISLAS TURCAS Y TURCOS"
	Case "TV" TranslateCountry = "TUVALU"
	Case "UG" TranslateCountry = "UGANDA"
	Case "UA" TranslateCountry = "UCRANIA"
	Case "AE" TranslateCountry = "EMIRATOS ÁRABES UNIDOS"
	Case "GB" TranslateCountry = "REINGO UNIDO O INGLATERRA"
	Case "US" TranslateCountry = "ESTADOS UNIDOS"
	Case "UM" TranslateCountry = "ISLAS MENORES Y PERÍFERICAS DE ESTADOS UNIDOS"
	Case "UY" TranslateCountry = "URUGUAY"
	Case "UZ" TranslateCountry = "UZBEKISTÁN"
	Case "VU" TranslateCountry = "VANUATU"
	Case "VE" TranslateCountry = "VENEZUELA"
	Case "VN" TranslateCountry = "VIETNAM"
	Case "VG" TranslateCountry = "ISLAS BRITÁNICAS VIRGINIA"
	Case "VI" TranslateCountry = "ISLAS DE ESTADOS UNIDOS VIRGINIA"
	Case "WF" TranslateCountry = "WALLIS Y FUTUNA"
	Case "EH" TranslateCountry = "SAHARA OCCIDENTAL"
	Case "YE" TranslateCountry = "YEMEN"
	Case "ZM" TranslateCountry = "ZAMBIA"
	Case "ZW" TranslateCountry = "ZIMBABWE"
	Case "GTLTF" TranslateCountry = "LATIN FREIGHT GUATEMALA"
    Case "SVLTF" TranslateCountry = "LATIN FREIGHT SV"
    Case "HNLTF" TranslateCountry = "LATIN FREIGHT HN"
    Case "NILTF" TranslateCountry = "LATIN FREIGHT NI"
    Case "CRLTF" TranslateCountry = "LATIN FREIGHT CR"
    Case "PALTF" TranslateCountry = "LATIN FREIGHT PA"
	Case "14.PAIS PROCEDENCIA" TranslateCountry = "14.PAIS PROCEDENCIA"
	Case "15.PAIS DESTINO" TranslateCountry = "15.PAIS DESTINO"
    Case "XX" TranslateCountry = "PAIS"
    Case "GTTLA" TranslateCountry = "TLA GT"
    Case "SVTLA" TranslateCountry = "TLA SV"
    Case "HNTLA" TranslateCountry = "TLA HN"
    Case "NITLA" TranslateCountry = "TLA NI"
    Case "CRTLA" TranslateCountry = "TLA CR"
    Case "PATLA" TranslateCountry = "TLA PA"
	Case "BZTLA" TranslateCountry = "TLA BZ"	
	Case "GTLGX" TranslateCountry = "LGX GT"
	Case "CRLGX" TranslateCountry = "LGX CR"
    Case "PALGX" TranslateCountry = "LGX PA"
	Case Else TranslateCountry = ""
End Select
End Function

Sub SearchSimilars (WSName, WSVal, GID, Separator, SearchOption)
Dim Chain, LenChain, MoreOptions, Lik, OrdType, sep, TempSQLQuery
	WSVal = UCase(WSVal)
	WSVal = Replace(WSVal, "GUATEMALA", "")
	WSVal = Replace(WSVal, "EL SALVADOR", "")
	WSVal = Replace(WSVal, "HONDURAS", "")
	WSVal = Replace(WSVal, "NICARAGUA", "")
	WSVal = Replace(WSVal, "COSTA RICA", "")
	WSVal = Replace(WSVal, "PANAMA", "")
	WSVal = Replace(WSVal, "BRASIL", "")
	WSVal = Replace(WSVal, "MEXICO", "")
	WSVal = Replace(WSVal, "BELICE", "")
	WSVal = Replace(WSVal, "OF.", "")
	WSVal = Replace(WSVal, "S.", " ")
	WSVal = Replace(WSVal, "A.", " ")
	WSVal = Replace(WSVal, ".", "")
	WSVal = Replace(WSVal, ",", "")
	WSVal = Replace(WSVal, " EL ", " ")
	WSVal = Replace(WSVal, " DE ", " ")
	WSVal = Replace(WSVal, " E ", " ")
	WSVal = Replace(WSVal, " DO ", " ")
	WSVal = Replace(WSVal, " DEL ", " ")
	WSVal = Replace(WSVal, " POR ", " ")
    WSVal = Replace(WSVal, " PARA ", " ")
	WSVal = Replace(WSVal, " Y ", " ")
	WSVal = Replace(WSVal, " & ", " ")
	WSVal = Replace(WSVal, "/", "")
	WSVal = Replace(WSVal, " LA ", " ")
	WSVal = Replace(WSVal, " LOS ", " ")
	WSVal = Replace(WSVal, " lAS ", " ")
	WSVal = Replace(WSVal, " SAN ", " ")
	WSVal = Replace(WSVal, " SA", " ")
	WSVal = Replace(WSVal, " S A", " ")
	WSVal = Replace(WSVal, " CA", " ")
	WSVal = Replace(WSVal, " LTDA", "")
	WSVal = Replace(WSVal, " CO ", "")
	WSVal = Replace(WSVal, " CO", "")
	WSVal = Replace(WSVal, " LTD", "")
	WSVal = Replace(WSVal, "LIMITADA", "")
	WSVal = Replace(WSVal, "SOCIEDAD", "")
	WSVal = Replace(WSVal, "ANONIMA", "")
	WSVal = Replace(WSVal, "INC", "")
	WSVal = Replace(WSVal, " SV", "")
	WSVal = Replace(WSVal, " CV", "")
	WSVal = Replace(WSVal, " RL", "")
	WSVal = Replace(WSVal, "   ", " ")
	WSVal = Replace(WSVal, "  ", " ")
    WSVal = Replace(WSVal, " EN ", "")
	WSVal = Trim (WSVal)

	LenChain = -1
	Chain = split(WSVal, Separator)
	LenChain = ubound(Chain)
	Lik = " like "
	OrdType = " order by a.Countries, a." & WSName
	
	Select Case GID
	Case 7,10
		SQLQuery = GetSQLSearch (GroupID) & " where a.id_cliente = d.id_cliente " & _
					"and d.id_nivel_geografico = n.id_nivel " & _
					"and n.id_pais = p.codigo "
		'SQLQuery =  "select a.id_cliente, a.hora_creacion, a.fecha_creacion, a.id_estatus, a.nombre_cliente, b.id_direccion,  " & _
		'			"from clientes a, direcciones b where a.id_cliente=b.id_cliente and a.id_estatus<>0 "
					
		Lik = " ilike "
		OrdType = " order by p.codigo, a." & WSName
	Case 8
		SQLQuery = GetSQLSearch (GroupID) & " where (a.activo=false or a.activo=true) "
		Lik = " ilike "
		OrdType = " order by a." & WSName
	Case 9,11 '9=Airports 11=Commodities/Productos
		SQLQuery = GetSQLSearch (GroupID) & " where (a.Expired=0 or a.Expired=1) "
		OrdType = " order by a." & WSName
	Case Else
		SQLQuery = GetSQLSearch (GroupID) & " where a.Countries in " & Session("Countries")
	End Select
	
	'if LenChain >= 0 then
	'	MoreOptions = 1
	'	SQLQuery = SQLQuery & " and ("
	'	CreateSearchQuery SQLQuery, WSName & lik & "'%" & Chain(0) & "%'", MoreOptions, ""
	'	for i = 1 to LenChain
	'		CreateSearchQuery SQLQuery, WSName & lik & "'%" & Chain(i) & "%'", MoreOptions, " or "
	'	next
	'	SQLQuery = SQLQuery & ")"
	'end if
	
	if LenChain >= 0 then
		MoreOptions = 1
		sep = ""
		for i = 0 to LenChain
			if Len(Chain(i))>1 then
				CreateSearchQuery TempSQLQuery, WSName & lik & "'%" & Chain(i) & "%'", MoreOptions, sep
				sep = " or "
			end if
		next
		if TempSQLQuery <> "" then
			SQLQuery = SQLQuery & " and (" & TempSQLQuery & ")"
		end if
	end if	

	'response.write SQLQuery
	Set rs = Conn.Execute(SQLQuery & OrdType)
	If Not rs.EOF Then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	End If
	CloseOBJs rs, Conn
	'CloseOBJ Conn
	
	HTMLCode = ""
	if CountTableValues >= 0 then
		select case GID
		case  7, 10
			for i = 0 to CountTableValues
			HTMLCode = HTMLCode & "<tr><td class='list'><a class=labellist href=InsertData.asp?OID=" & aTableValues(0, i) & "&GID=" & GroupID & "&CD=" & aTableValues(2, i) & "&CT=" & aTableValues(1, i) & "&SO=" & SearchOption & "&AID=" & aTableValues(7, i) & ">" & aTableValues(3, i) & " - " & aTableValues(0, i) & " - " & aTableValues(5, i) & "</a></td></tr>"
			next
		case 8
			for i = 0 to CountTableValues
			HTMLCode = HTMLCode & "<tr><td class='list'><a class=labellist href=InsertData.asp?OID=" & aTableValues(0, i) & "&GID=" & GroupID & "&CD=" & aTableValues(2, i) & "&CT=" & aTableValues(1, i) & "&SO=" & SearchOption & ">" & aTableValues(0, i) & " - " & aTableValues(4, i) & "</a></td></tr>"
			next
		case else
			for i = 0 to CountTableValues
			HTMLCode = HTMLCode & "<tr><td class='list'><a class=labellist href=InsertData.asp?OID=" & aTableValues(0, i) & "&GID=" & GroupID & "&CD=" & aTableValues(2, i) & "&CT=" & aTableValues(1, i) & "&SO=" & SearchOption & ">" & aTableValues(4, i) & " - " & aTableValues(0, i) & " - " & aTableValues(5, i) & "</a></td></tr>"
			next
		end select
	end if
End Sub

Function VirtualForm
Dim FormElements, LenFormElements, Elements, ntr
	VirtualForm = ""
	ntr = chr(13)
	FormElements = Split(Request.form, "&")
	LenFormElements = ubound(FormElements)
	for i = 0 to LenFormElements
		Elements = Split(URLDecode(FormElements(i)),"=")
		VirtualForm = VirtualForm & "<INPUT name='" & Elements(0) & "' type=hidden value='" & Elements(1) & "'>" & ntr
	next
End Function

Function URLDecode(sConvert)
    Dim aSplit, sOutput, i, lenaSplit
	
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If
	
    ' convert all pluses to spaces
    sOutput = Replace(sConvert, "+", " ")
	
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")
	lenaSplit = UBound(aSplit) - 1
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to lenaSplit
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If	
    URLDecode = sOutput
End Function

Function GetSQLSearch (GroupID)
	Select Case GroupID
	Case 2
		GetSQLSearch = "select a.CarrierID, a.CreatedTime, a.CreatedDate, a.CarrierCode, a.Name, a.Expired from Carriers a"
	Case 7, 10
		GetSQLSearch = 	"select a.id_cliente, a.hora_creacion, a.fecha_creacion, p.codigo, a.id_cliente, a.nombre_cliente, a.id_estatus, d.id_direccion " & _
							"from clientes a, direcciones d, niveles_geograficos n, paises p "						
	'Case 7
	'	GetSQLSearch = "select a.ConsignerID, a.CreatedTime, a.CreatedDate, a.AccountNo, a.Name, a.Expired from Consigners a"
	Case 8
		GetSQLSearch = "select a.agente_id, a.fecha_creacion, a.hora_creacion, a.agente_id, a.agente, a.contacto, a.activo from agentes a"
		'GetSQLSearch = "select a.AgentID, a.CreatedTime, a.CreatedDate, a.AccountNo, a.Name, a.IATANo, a.Expired from Agents a"
	Case 9
		GetSQLSearch = "select a.AirportID, a.CreatedTime, a.CreatedDate, a.AirportCode, a.Name, a.Expired from Airports a"
	'Case 10
	'	GetSQLSearch = "select a.ShipperID, a.CreatedTime, a.CreatedDate, a.AccountNo, a.Name, a.Expired from Shippers a"
	Case 11
		GetSQLSearch = "select a.CommodityID, a.CreatedTime, a.CreatedDate, a.CommodityCode, a.NameES, a.Expired from Commodities a"
	Case 12
		GetSQLSearch = "select a.CurrencyID, a.CreatedTime, a.CreatedDate, a.CurrencyCode, a.Name, a.Expired from Currencies a"
	End Select
End Function


Function limpiarCadenaNombreFichero(cadenaTexto, sustituirPor) 'creada version nueva 
  Dim tamanoCadena, i, cadenaResultado, caracteresValidos, caracterActual 
  
  tamanoCadena = Len(cadenaTexto)
  If tamanoCadena > 0 Then
    caracteresValidos = " 0123456789abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ.-/"
    For i = 1 To tamanoCadena
      caracterActual = Mid(cadenaTexto, i, 1)
      If InStr(caracteresValidos, caracterActual) Then
        cadenaResultado = cadenaResultado & caracterActual
      Else
        cadenaResultado = cadenaResultado & sustituirPor
      End If
    Next
  End If
  
  limpiarCadenaNombreFichero = cadenaResultado
End Function

Sub MasterData (Conn, ObjectID, AddressID, Name, Address, Phone1, Phone2, Attn, AccountNo, IATANo, Coloader)
Dim rs
	Name = ""
	Address = ""
	Phone1 = ""
	Phone2 = ""
	Attn = ""
	AccountNo = ""
	IATANo= ""
    Coloader = 0
	QuerySelect = "select a.nombre_cliente, d.direccion_completa, d.""phone_number"", a.id_cliente, a.es_coloader " & _
							"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
							" where a.id_cliente = d.id_cliente" & _
							" and d.id_nivel_geografico = n.id_nivel" & _
							" and n.id_pais = p.codigo" & _
							" and a.id_cliente = " & ObjectID  & _
							" and d.id_direccion = " & AddressID

	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
        Name = limpiarCadenaNombreFichero(rs(0), "")
        Address = limpiarCadenaNombreFichero(rs(1), "")
		Phone1 = rs(2)
        Coloader = CheckNum(rs(4))
	end if
	CloseOBJ rs

	set rs = Conn.Execute("select numero_telefono from cli_telefonos where id_cliente=" & ObjectID)
	if Not rs.EOF then
		Phone2 = rs(0)
	end if
	CloseOBJ rs

	set rs = Conn.Execute("select nombres from contactos where id_cliente=" & ObjectID)
	if Not rs.EOF then
        Attn = limpiarCadenaNombreFichero(rs(0), "")
	end if
	CloseOBJ rs

	set rs = Conn.Execute("select no_cuenta, no_iata from clientes_aereo where id_cliente=" & ObjectID)
	if Not rs.EOF then
		AccountNo = rs(0)
		IATANo = rs(1)
	end if
	CloseOBJ rs
End Sub

Sub SaveChargeItems (Conn, AWBID, DocTyp)
Dim CrtDate, CrtTime
	FormatTime CrtDate, CrtTime
	Conn.Execute("update ChargeItems set Expired=1 where AWBID=" & AWBID & " and DocTyp=" & DocTyp & " and (InvoiceID=0 or DocType=9) and InterProviderType<>5 and InterChargeType<>2")
    'response.write "update ChargeItems set Expired=1 where AWBID=" & AWBID & " and DocTyp=" & DocTyp & " and InvoiceID=0 and InterProviderType<>5 and InterChargeType<>2" & "<br><br>"
	'Valores Fijos de Agente y Transportista
	if DocTyp = 0 then 'Export
		if CheckNum(CustomFee) > 0 then
			if CheckNum(Request.Form("INVCF"))=0 then
				CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CCF") & "', 14, " & CheckNum(Request.Form("CustomFee"))  & ", " & CheckNum(Request.Form("TCCF"))  & ", 0, " & DocTyp & ", 'CUSTOM FEE', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", 3, 'AEREO', " & CheckNum(Request.Form("TPCF"))  & ", " & CheckNum(Request.Form("CustomFee_Routing")) & ", '" & Request.Form("CF_Tarifa") & "', '" & Request.Form("CF_Regimen") & "', '" & Request.Form("CF_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"
			end if
		end if
		
		if CheckNum(TerminalFee) > 0 then
			if CheckNum(Request.Form("INVTF"))=0 then
				CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CTF") & "', 15, " & CheckNum(Request.Form("TerminalFee"))  & ", " & CheckNum(Request.Form("TCTF"))  & ", 0, " & DocTyp & ", 'TERMINAL FEE', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", 3, 'AEREO', " & CheckNum(Request.Form("TPTF"))  & ", " & CheckNum(Request.Form("TerminalFee_Routing")) & ", '" & Request.Form("TF_Tarifa") & "', '" & Request.Form("TF_Regimen") & "', '" & Request.Form("TF_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"
			end if
		end if
	end if
	if CheckNum(TotCarrierRate) > 0 and TotCarrierRate <> "AS AGREED" then
		if CheckNum(Request.Form("INVAF"))=0 then
			CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CAF") & "', 11, " & CheckNum(Request.Form("TotCarrierRate"))  & ", " & CheckNum(Request.Form("TCAF"))  & ", 0, " & DocTyp & ", 'AIR FREIGHT', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", 3, 'AEREO', " & CheckNum(Request.Form("TPAF"))  & ", " & CheckNum(Request.Form("TotCarrierRate_Routing")) & ", '" & Request.Form("AF_Tarifa") & "', '" & Request.Form("AF_Regimen") & "', '" & Request.Form("AF_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"
		end if
	end if
	if CheckNum(FuelSurcharge) > 0 then
		if CheckNum(Request.Form("INVFS"))=0 then
			CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CFS") & "', 12, " & CheckNum(Request.Form("FuelSurcharge"))  & ", " & CheckNum(Request.Form("TCFS"))  & ", 0, " & DocTyp & ", 'FUEL SURCHARGE', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", 3, 'AEREO', " & CheckNum(Request.Form("TPFS"))  & ", " & CheckNum(Request.Form("FuelSurcharge_Routing")) & ", '" & Request.Form("FS_Tarifa") & "', '" & Request.Form("FS_Regimen") & "', '" & Request.Form("FS_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"
		end if
	end if
	if CheckNum(SecurityFee) > 0 then
		if CheckNum(Request.Form("INVSF"))=0 then
			CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CSF") & "', 13, " & CheckNum(Request.Form("SecurityFee"))  & ", " & CheckNum(Request.Form("TCSF"))  & ", 0, " & DocTyp & ", 'SECURITY FEE', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", 3, 'AEREO', " & CheckNum(Request.Form("TPSF"))  & ", " & CheckNum(Request.Form("SecurityFee_Routing")) & ", '" & Request.Form("SF_Tarifa") & "', '" & Request.Form("SF_Regimen") & "', '" & Request.Form("SF_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"
		end if
	end if
	if CheckNum(PickUp) > 0 then
		if CheckNum(Request.Form("INVPU"))=0 then
			CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CPU") & "', 31, " & CheckNum(Request.Form("PickUp"))  & ", " & CheckNum(Request.Form("TCPU"))  & ", 1, " & DocTyp & ", 'PICK UP', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", 3, 'AEREO', " & CheckNum(Request.Form("TPPU"))  & ", " & CheckNum(Request.Form("PickUp_Routing")) & ", '" & Request.Form("PU_Tarifa") & "', '" & Request.Form("PU_Regimen") & "', '" & Request.Form("PU_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"
		end if
	end if
	if CheckNum(SedFilingFee) > 0 then
		if CheckNum(Request.Form("INVFF"))=0 then
			CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CFF") & "', 38, " & CheckNum(Request.Form("SedFilingFee"))  & ", " & CheckNum(Request.Form("TCFF"))  & ", 1, " & DocTyp & ", 'SED FILING FEE', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", 3, 'AEREO', " & CheckNum(Request.Form("TPFF"))  & ", " & CheckNum(Request.Form("SedFilingFee_Routing")) & ", '" & Request.Form("FF_Tarifa") & "', '" & Request.Form("FF_Regimen") & "', '" & Request.Form("FF_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"			
		end if
	end if
	if CheckNum(Intermodal) > 0 then
		if CheckNum(Request.Form("INVIM"))=0 then
			CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CIM") & "', 115, " & CheckNum(Request.Form("Intermodal"))  & ", " & CheckNum(Request.Form("TCIM"))  & ", 1, " & DocTyp & ", 'INTERMODAL', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", 5, 'TERRESTRE', " & CheckNum(Request.Form("TPIM"))  & ", " & CheckNum(Request.Form("Intermodal_Routing")) & ", '" & Request.Form("IM_Tarifa") & "', '" & Request.Form("IM_Regimen") & "', '" & Request.Form("IM_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"			
		end if
	end if
	if CheckNum(PBA) > 0 then
		if CheckNum(Request.Form("INVPB"))=0 then
			CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, ServiceID, ServiceName, PrepaidCollect, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CPB") & "', 116, " & CheckNum(Request.Form("PBA"))  & ", " & CheckNum(Request.Form("TCPB"))  & ", 1, " & DocTyp & ", 'PBA', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", 3, 'AEREO', " & CheckNum(Request.Form("TPPB"))  & ", " & CheckNum(Request.Form("PBA_Routing")) & ", '" & Request.Form("PB_Tarifa") & "', '" & Request.Form("PB_Regimen") & "', '" & Request.Form("PB_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"			
		end if
	end if

    dim tmp_j
	'Valores del Transportista
	for i=0 to 3
    'for i=2 to 7

        'If InStr(1, "3,4,5,8", i+1) > 0 Then

        select case i
        case 0 tmp_j = 3
        case 1 tmp_j = 4
        case 3 tmp_j = 5
        case 4 tmp_j = 8
        end select

		if CheckNum(Request.Form("VC"&(i+1))) > 0 then
			if CheckNum(Request.Form("INVC"&(i+1))) = 0 then
				CrtTime = CrtTime+1
                'response.write i & "a<br>"
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, Pos, ServiceID, ServiceName, PrepaidCollect, CalcInBL, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CC"&(i+1)) & "', " & CheckNum(Request.Form("C"&(i+1))) & ", " & CheckNum(Request.Form("VC"&(i+1)))  & ", " & CheckNum(Request.Form("TCC"&(i+1)))  & ", 0, " & DocTyp & ", '" & Request.Form("NC"&(i+1)) & "', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", " & i+1 & ", " & Request.Form("SVIC"&(i+1)) & ", '" & Request.Form("SVNC"&(i+1)) & "', " & CheckNum(Request.Form("TPC"&(i+1)))  & ", " & CheckNum(Request.Form("CCBLC"&(i+1))) & ", " & CheckNum(Request.Form("AdditionalChargeName" & tmp_j & "_Routing")) & ", '" & Request.Form("C"&(i+1)&"_Tarifa") & "', '" & Request.Form("C"&(i+1)&"_Regimen") & "', '" & Request.Form("C"&(i+1)&"_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"			
			end if
		end if

        'End If
	next
	'Valores del Agente
  	for i=0 to 10
    'for i=0 to 14

        'If InStr(1, "1,2,6,7,9,10,11,12,13,14,15", i+1) > 0 Then

        select case i
        case 0 tmp_j = 1
        case 1 tmp_j = 2
        case 2 tmp_j = 6
        case 3 tmp_j = 7
        case 4 tmp_j = 9
        case 5 tmp_j = 10
        case 6 tmp_j = 11
        case 7 tmp_j = 12
        case 8 tmp_j = 13
        case 9 tmp_j = 14
        case 10 tmp_j = 15        
        end select
         
        'response.write "(" & i & ")<br>"
        'response.write "(" & CheckNum(Request.Form("VA"&(i+1))) & ")"
        'response.write "(" & CheckNum(Request.Form("INVA"&(i+1))) & ")<br>"

		if CheckNum(Request.Form("VA"&(i+1))) > 0 then
			if CheckNum(Request.Form("INVA"&(i+1))) = 0 then
				CrtTime = CrtTime+1
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, Pos, ServiceID, ServiceName, PrepaidCollect, CalcInBL, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CA"&(i+1)) & "', " & CheckNum(Request.Form("A"&(i+1))) & ", " & CheckNum(Request.Form("VA"&(i+1)))  & ", " & CheckNum(Request.Form("TCA"&(i+1)))  & ", 1, " & DocTyp & ", '" & Request.Form("NA"&(i+1)) & "', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", " & i+1 & ", " & Request.Form("SVIA"&(i+1)) & ", '" & Request.Form("SVNA"&(i+1)) & "', " & CheckNum(Request.Form("TPA"&(i+1)))  & ", " & CheckNum(Request.Form("CCBLA"&(i+1))) & ", " & CheckNum(Request.Form("AdditionalChargeName" & tmp_j & "_Routing")) & ", '" & Request.Form("A"&(i+1)&"_Tarifa") & "', '" & Request.Form("A"&(i+1)&"_Regimen") & "', '" & Request.Form("A"&(i+1)&"_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"			
			end if
		end if

        'End If
	next
	if DocTyp = 1 then 'Import
		for i=0 to 5
			if CheckNum(Request.Form("VO"&(i+1))) > 0 then
				if CheckNum(Request.Form("INVO"&(i+1))) = 0 then
					CrtTime = CrtTime+1
                    'response.write i & "c<br>"
                QuerySelect = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, Pos, ServiceID, ServiceName, PrepaidCollect, CalcInBL, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo) values(" & AWBID  & ", '" & Request.Form("CO"&(i+1)) & "', " & CheckNum(Request.Form("O"&(i+1))) & ", " & CheckNum(Request.Form("VO"&(i+1)))  & ", " & CheckNum(Request.Form("TCO"&(i+1)))  & ", 2, " & DocTyp & ", '" & Request.Form("NO"&(i+1)) & "', '" & CrtDate & "', " & CrtTime & ", " & Session("OperatorID") & ", " & i+1 & ", " & Request.Form("SVIO"&(i+1)) & ", '" & Request.Form("SVNO"&(i+1)) & "', " & CheckNum(Request.Form("TPO"&(i+1)))  & ", " & CheckNum(Request.Form("CCBLO"&(i+1))) & ", " & CheckNum(Request.Form("OtherChargeName" & (i+1) & "_Routing")) & ", '" & Request.Form("O"&(i+1)&"_Tarifa") & "', '" & Request.Form("O"&(i+1)&"_Regimen") & "', '" & Request.Form("O"&(i+1)&"_TarifaTipo") & "')"
				Conn.Execute(QuerySelect)
				'response.write QuerySelect & "<br>"			
				end if
			end if
		next
	end if
End Sub


Function SaveInterChargeItems (Conn, AWBID, DocTyp, Countries)
Dim rs, ItemCurrs, ItemIDs, ItemVals, ItemLocs, ItemNames, ItemOVals, ItemPPCCs, ItemServIDs, ItemServNames, ItemInvoices
Dim ItemCalcInBls, CantItems, CreatedDate, CreatedTime, ItemInterCompanyIDs, ItemNames_Routing
Dim InterCompanyID, Result, IntercompanyChecked

    FormatTime CreatedDate, CreatedTime
	
	ItemCurrs = Split(Request.Form("ItemCurrs"), "|")
	ItemIDs = Split(Request.Form("ItemIDs"), "|")
	ItemServIDs = Split(Request.Form("ItemServIDs"), "|")
	ItemServNames = Split(Request.Form("ItemServNames"), "|")
	
    ItemVals = Split( Replace(Request.Form("ItemVals") , ",", ".") , "|")    'puede traer comas

	ItemLocs = Split(Request.Form("ItemLocs"), "|")
	ItemNames = Split(Request.Form("ItemNames"), "|")
	ItemOVals = Split(Request.Form("ItemOVals"), "|")
	ItemPPCCs = Split(Request.Form("ItemPPCCs"), "|")
	CantItems = CheckNum(Request.Form("CantItems"))
	ItemInvoices = Split(Request.Form("ItemInvoices"), "|")
	ItemCalcInBLs = Split(Request.Form("ItemCalcInBls"), "|")
	ItemInterCompanyIDs = Split(Request.Form("ItemInterCompanyIDs"), "|")
    ItemNames_Routing = Split(Request.Form("ItemNames_Routing"), "|")    
    
    'response.write ("Paso por aqui 1<br>")
    
    'Revisando que todos los IntercompanyID sean iguales
    if CantItems>=0 then
        IntercompanyChecked=0
        
        'Revisando que todos los IntercompanyID sean iguales, si son diferentes no guarda los rubros Intercompany
        for i=0 to CantItems
            
            if ItemInterCompanyIDs(i) <> "" then

            if ItemInterCompanyIDs(i) <> 0 then
                if InterCompanyID=0 Then
                    InterCompanyID = ItemInterCompanyIDs(i)
                    IntercompanyChecked=1
                else
		            if InterCompanyID <> ItemInterCompanyIDs(i) then
			             'Existe algun codigo de Intercompany diferente, lo que no permite registrar los datos
                         IntercompanyChecked=2
		            end if
                end if
            end if

            end if
	    next

        dim query_tmp 
        select Case IntercompanyChecked
        Case 1
            'Registrando los cobros Intercompany
            query_tmp = "update ChargeItems set Expired=1 where AWBID=" & AWBID & " and DocTyp=" & DocTyp & " and InvoiceID=0 and InterProviderType=5 and InterChargeType=2"
            'response.write (query_tmp & "<br>")
            Conn.Execute(query_tmp)
	        for i=0 to CantItems
		        if ItemVals(i) > 0 then
			        if ItemInvoices(i)=0 then 'Guardando solo los nuevos rubros que no se han facturado en el BAW							        
                        query_tmp = "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, ItemName, CreatedDate, CreatedTime, UserID, OverSold, PrepaidCollect, ServiceID, ServiceName, CalcInBL, InterChargeType, InterCompanyID, InterProviderType, Doctyp, ItemName_Routing) values (" & AWBID  & ", '" & ItemCurrs(i) & "', " & ItemIDs(i) & ", " & ItemVals(i)  & ", " & ItemLocs(i)  & ", 1, '" & ItemNames(i) & "', '" & CreatedDate & "', " & CreatedTime & ", " & Session("OperatorID") & ", " & ItemOVals(i) & ", " & ItemPPCCs(i) & ", " & ItemServIDs(i) & ", '" & Trim(ItemServNames(i)) & "', " & ItemCalcInBLs(i) & ", 2, " & CheckNum(ItemInterCompanyIDs(i)) & ", 5, " & DocTyp & "," & CheckNum(ItemNames_Routing(i)) & ")"
                        'response.write (query_tmp & "<br>")
                        Conn.Execute(query_tmp)
				        CreatedTime = CreatedTime+1
			        end if
		        end if
	        next

            'Llamando al BAW para que genere los documentos intercompany correspondientes al aereo, SysID=2
    	    if Countries<>"" then
			    'Ultimo parametro es 0 en la funcion significa ir a provisionar costos
			    Result = SetBAWIntercompany(AWBID, 2, DocTyp+1, Countries, Session("Login"), InterCompanyID)
		    end if
        Case 2
            Result = "Existen Codigos de Intercompany Distintos, en esta operacion puede registrar datos solo a un Intercompany"
        end Select

            

    else
        Result = "No hay datos para registrar Intercompany"
    end if
    SaveInterChargeItems = Result
End Function






Sub SaveCostItems (Conn, BLID, Action)
Dim ItemNames, ItemIDs, ItemThirdParties, ItemCurrs, ItemCosts, ItemSTypes, ItemSIDs, ItemSNames, ItemDistribs, ItemSPOs, ItemSPRs, ItemPos, ItemServIDs, ItemServNames, ItemProvitions, ItemSAffected, ItemNeutrales
Dim CantHouses, CantItems, CreatedDate, CreatedTime, CostID, HouseCost, DocType, Countries, ItemNames_, ItemIDs_, ItemServIDs_, ItemServNames_, ItemSPRsDate

	FormatTime CreatedDate, CreatedTime

    ItemNames_ = Split(Request.Form("ItemNames_"), "|") 
    ItemIDs_ = Split(Request.Form("ItemIDs_"), "|")
    ItemServIDs_ = Split(Request.Form("ItemServIDs_"), "|")
    ItemServNames_ = Split(Request.Form("ItemServNames_"), "|")
	ItemSPRsDate = Split(Request.Form("ItemSPRsDate"), "|")

	ItemServIDs = Split(Request.Form("ItemServIDs"), "|")
	ItemServNames = Split(Request.Form("ItemServNames"), "|")
	ItemNames = Split(Request.Form("ItemNames"), "|")
	ItemIDs = Split(Request.Form("ItemIDs"), "|")
	ItemThirdParties = Split(Request.Form("ItemThirdParties"), "|")
	ItemCurrs = Split(Request.Form("ItemCurrs"), "|")
	ItemCosts = Split(Request.Form("ItemCosts"), "|")
	ItemSTypes = Split(Request.Form("ItemSTypes"), "|")
	ItemSIDs = Split(Request.Form("ItemSIDs"), "|")
	ItemSNames = Split(Request.Form("ItemSNames"), "|")
	ItemDistribs = Split(Request.Form("ItemDistribs"), "|")
	ItemSAffected = Split(Request.Form("ItemSAffected"), "|")
	ItemSPOs = Split(Request.Form("ItemSPOs"), "|")
	ItemSPRs = Split(Request.Form("ItemSPRs"), "|")
	CantItems = CheckNum(Request.Form("CantItems"))
	DocType = CheckNum(Request.Form("DocType"))
	CantHouses = CheckNum(Request.Form("CantHouses")) 
	ItemPos = Split(Request.Form("ItemPos"), "|")
	ItemProvitions = Split(Request.Form("ItemProvitions"), "|")
	ItemNeutrales = Split(Request.Form("ItemNeutrales"), "|")
    
	if Action=2 then
        Conn.Execute("update Costs set Expired=1 where BLID=" & BLID & " and DocType=" & DocType & " and ProvisionID=0 and cxp_exactus_id=0")
		'Conn.Execute("update Costs set Expired=1 where BLID=" & BLID & " and DocType=" & DocType & " and ProvisionID=0")
	end if
	
	if DocType=1 then 'Export
		set rs = Conn.Execute("select Countries from Awb where AWBID=" & BLID)
	else 'Import
		set rs = Conn.Execute("select Countries from Awbi where AWBID=" & BLID)
	end if
	if Not rs.EOF then
		Countries = rs(0)
	end if
	CloseOBJ rs
	
    Dim TestArray1

	for i=0 to CantItems
		if ItemCosts(i) > 0 then
			if ItemProvitions(i)=0 then 'Guardando solo los nuevos rubros que no se han provisionado en el BAW	
               		

                On Error Resume Next

                TestArray1 = Split(ItemSPRsDate(i),"/")
                ItemSPRsDate(i) = TestArray1(2) & "-" & TestArray1(1) & "-" & TestArray1(0)


				'Guardando los costos de cada Master

                QuerySelect = "INSERT INTO Costs (BLID, Currency, ItemID, Cost, ItemName, CreatedDate, CreatedTime, UserID, SupplierType, SupplierID, SupplierName, Distribution, DocType, PurchaseOrder, ServiceID, ServiceName, Reference, Countries, ThirdParties, IsAffected, SupplierNeutral, " & _ 
                
                "SubTipoID, SubTipoName, TipoDocID, TipoDocName, ReferenceDate) VALUES (" & BLID  & ", '" & ItemCurrs(i) & "', " & ItemIDs(i) & ", " & ItemCosts(i) & ", '" & ItemNames(i) & "', '" & CreatedDate & "', " & CreatedTime & ", " & Session("OperatorID") & ", " & ItemSTypes(i) & ", " & ItemSIDs(i) & ", '" & ItemSNames(i) & "', " & ItemDistribs(i) & ", " & DocType & ", '" & Trim(ItemSPOs(i)) & "', " & ItemServIDs(i) & ", '" & Trim(ItemServNames(i)) & "', '" & Trim(ItemSPRs(i)) & "', '" & Countries & "', " & ItemThirdParties(i) & ", " & ItemSAffected(i) & ", '" & ItemNeutrales(i) & "', " & _ 
                
                "'" & ItemIDs_(i) & "', '" & ItemNames_(i) & "', '" & ItemServIDs_(i) & "', '" & Trim(ItemServNames_(i)) & "', '" & ItemSPRsDate(i) & "')"
                
				'response.write(QuerySelect & "<br>")					
				Conn.Execute(QuerySelect)
				
                If Err.Number<>0 then                                '
	                response.write Err.Number & " - " & Err.Description & "<br>"
                end if
                                
				'Obteniendo el CostID para guardar sus respectivos costos a cada House
				Set rs = Conn.Execute("select CostID from Costs where BLID=" & BLID & " and ItemID=" & ItemIDs(i) & " and CreatedDate='" & CreatedDate & "' and CreatedTime=" & CreatedTime)
				if Not rs.EOF then
					CostID = rs(0)
				end if
				CloseOBJ rs
				'Guardando los costos distribuidos a cada House si el tipo de Distribucion no es Directo (4=Sin Distribucion)
				if ItemDistribs(i)<>4 then
					for j=0 to CantHouses
						HouseCost = CheckNum(Request("H"&j&"C"&ItemPos(i)))
						if HouseCost <> 0 then
							Conn.Execute("insert into CostsDetail(CostID, SBLID, Cost) values (" & CostID & ", " & CheckNum(Request("SBLID"&j)) & ", " & HouseCost & ")")
							'response.write ("insert into CostsDetail(CostID, SBLID, Cost) values (" & CostID & ", " & CheckNum(Request("SBLID"&j)) & ", " & HouseCost & ")<br>")
						end if
					next
				end if
				CreatedTime = CreatedTime+1
			end if
		end if
	next
End Sub



Function GetEstatusCombex(GUIA, TIPO)

    Dim oXmlHTTP, SOAPRequest, responseText, xmlhttp, xmlResponse, xnodelist, objItem, ResultCode
    Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
    oXmlHTTP.Open "POST", "http://" & soapServerCombex & soapPathCombex, False	    
    oXmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8" 
    
    SOAPRequest = _
      "<?xml version=""1.0"" encoding=""utf-8""?>" &_
      "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">" &_
        "<soap12:Body>" &_
            "<Guia xmlns=""COMBEX_SERVICE"">" & _
                "<Guia>" & GUIA & "</Guia>" & _
                "<Tipo>" & TIPO & "</Tipo>" & _
            "</Guia>" & _
        "</soap12:Body>" &_
      "</soap12:Envelope>"


    'response.write SOAPRequest & "<br>"

    oXmlHTTP.send SOAPRequest    

    responseText = ""
    ResultCode = oXmlHTTP.Status

    If ResultCode = 200 Then ' Response from server was success
		responseText = oXmlHTTP.responseText
	End If

	If Len(responseText) <> 0 Then  

        'response.ContentType="text/xml"
        'response.write responseText

		Set xmlResponse = CreateObject("Microsoft.XMLDOM")  
		xmlResponse.async = false  
        xmlResponse.setProperty "ServerHTTPRequest", True
		xmlResponse.loadXml responseText  
        Set xnodelist = xmlResponse.SelectNodes("/soap:Envelope/soap:Body/" & strWebMethodCombex & "Response/" & strWebMethodCombex & "Result/GUIA_COMBEX/GUIA")        

        ResultCode = ""
        For Each objItem In xnodelist
            ResultCode = ResultCode & objItem.SelectSingleNode("GUIA").text
            ResultCode = ResultCode & "|" & objItem.SelectSingleNode("FECHA_DE_INGRESO").text
            ResultCode = ResultCode & "|" & objItem.SelectSingleNode("FECHA_UBICACION").text
            ResultCode = ResultCode & "|" & objItem.SelectSingleNode("INI_REV_SAT").text
            ResultCode = ResultCode & "|" & objItem.SelectSingleNode("FECHA_FACTURADA").text
            ResultCode = ResultCode & "|" & objItem.SelectSingleNode("EGRESO_CARGA").text
        Next

		Set xmlResponse = nothing
		Set xnodelist = nothing

	End if 
        
    GetEstatusCombex = ResultCode

End Function




Function SetBAWIntercompany (BLID, SysID, OP, CT, USU, IntercompanyID)
	Dim xmlhttp, strSoap, responseText, StatusCode, ResultCode, xmlResponse, xnodelist, strWebMethod, objItem, CT_BAW
			
    'Traduciendo pais ISO en ID de pais del BAW
	CT_BAW = SetCountryBAW(CT)
    
    set xmlhttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	
	Select Case OP
	Case 1 'Exportacion
		OP = 4
	Case Else 'Importacion
		OP = 3
	End Select

	strWebMethod = "Contabilizar_Automaticamente"
	strSoap = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
	"<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & _
	" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">" & _
	"<soap12:Body>" & _        	
		"<Contabilizar_Automaticamente xmlns=""http://" & soapServer & soapPathIntercompany & "/"">" & _
		"<_empresaORIGENID>" & CT_BAW & "</_empresaORIGENID>" & _
        "<_intercompanyDESTINOID>" & IntercompanyID & "</_intercompanyDESTINOID>" & _
		"<_sisID>" & SysID & "</_sisID>" & _
		"<_ttoID>" & OP & "</_ttoID>" & _
		"<_blID>" & BLID & "</_blID>" & _
		"<_usuID>" & USU & "</_usuID>" & _
		"</Contabilizar_Automaticamente>" & _
	"</soap12:Body>" & _
	"</soap12:Envelope>"

	'Send Soap Request 
    'response.write strSoap & "<br>"

	xmlhttp.Open "POST", "http://" & soapServer & soapPathIntercompany, False	' False = Do not respond immediately
	xmlhttp.setRequestHeader "Man", "POST " & soapPathIntercompany & " HTTP/1.1"
	xmlhttp.setRequestHeader "Host", soapServer
	xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	xmlhttp.setRequestHeader "SOAPAction", "http://" & soapServer & soapPathIntercompany & "/" & strWebMethod
	xmlhttp.send(strSoap)
	
	If xmlhttp.Status = 200 Then			' Response from server was success
		responseText = xmlhttp.responseText
	Else									' Response from server failed
		StatusCode = xmlhttp.Status
	End If
	set xmlhttp = nothing
	
	'-------------INTERPRETAR EL XML RESPUESTA
	
	If Len(responseText) <> 0 Then  
		Set xmlResponse = CreateObject("MSXML2.DOMDocument")  
		xmlResponse.async = false  
		xmlResponse.loadXml responseText  
		Set xnodelist = xmlResponse.documentElement.selectNodes("/soap:Envelope/soap:Body/Contabilizar_AutomaticamenteResponse")
		For Each objItem In xnodelist   
			ResultCode =  objItem.selectSingleNode("Contabilizar_AutomaticamenteResult").Text
		Next
		set xmlResponse = nothing
	Else
		ResultCode = StatusCode
	End if
    
    SetBAWIntercompany = ResultCode
End Function

Function SetBAWProvition_Test (BLID, SysID, OP, CT, USU, Action)
    SetBAWProvition = 0
End Function

Function SetBAWProvition (BLID, SysID, OP, CT, USU, Action)
	Dim xmlhttp, strSoap, responseText, StatusCode, ResultCode, xmlResponse, xnodelist, strWebMethod, objItem, CT_BAW
			
    'Traduciendo pais ISO en ID de pais del BAW
	CT_BAW = SetCountryBAW(CT)
    
    set xmlhttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	
	Select Case OP
	Case 1 'Exportacion
		OP = 4
	Case Else 'Importacion
		OP = 3
	End Select

	Select Case Action
	Case 0
		strWebMethod = "Provisionar_Costos"
		strSoap = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & _
		" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
		"<soap:Body>" & _        	
		  "<Provisionar_Costos xmlns=""http://" & soapServer & soapPath & "/"">" & _
		  "<_blID>" & BLID & "</_blID>" & _
		  "<_sisID>" & SysID & "</_sisID>" & _
		  "<_Tipo_Operacion>" & OP & "</_Tipo_Operacion>" & _
		  "<_paisID>" & CT_BAW & "</_paisID>" & _
		  "<_usuID>" & USU & "</_usuID>" & _
		  "</Provisionar_Costos>" & _
		"</soap:Body>" & _
		"</soap:Envelope>"
	Case 1
		strWebMethod = "Get_Alerta"
		strSoap = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & _
		" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
		"<soap:Body>" & _        	
		  "<Get_Alerta xmlns=""http://" & soapServer & soapPath & "/"">" & _
		  "<ID>" & BLID & "</ID>" & _
		  "</Get_Alerta>" & _
		"</soap:Body>" & _
		"</soap:Envelope>"
	Case 2
		strWebMethod = "Validar_Cobros_Pendientes"
		strSoap = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & _
		" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
		"<soap:Body>" & _        	
		  "<Validar_Cobros_Pendientes xmlns=""http://" & soapServer & soapPath & "/"">" & _
		  "<v_BLID>" & BLID & "</v_BLID>" & _
		  "<v_Tipo_Operacion>" & OP & "</v_Tipo_Operacion>" & _
		  "<_paisID>" & CT_BAW & "</_paisID>" & _
		  "</Validar_Cobros_Pendientes>" & _
		"</soap:Body>" & _
		"</soap:Envelope>"
	End Select

	'Send Soap Request 

    'response.write strSoap & "<br>"

	xmlhttp.Open "POST", "http://" & soapServer & soapPath, False	' False = Do not respond immediately
	xmlhttp.setRequestHeader "Man", "POST " & soapPath & " HTTP/1.1"
	xmlhttp.setRequestHeader "Host", soapServer
	xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	xmlhttp.setRequestHeader "SOAPAction", "http://" & soapServer & soapPath & "/" & strWebMethod
	xmlhttp.send(strSoap)
	
	If xmlhttp.Status = 200 Then			' Response from server was success
		responseText = xmlhttp.responseText
	Else									' Response from server failed
		StatusCode = xmlhttp.Status
	End If
	set xmlhttp = nothing
	
	'-------------INTERPRETAR EL XML RESPUESTA
	
	If Len(responseText) <> 0 Then  
		Set xmlResponse = CreateObject("MSXML2.DOMDocument")  
		xmlResponse.async = false  
		xmlResponse.loadXml responseText  
		Select Case Action
		Case 0 'Provisionar_Costos
			Set xnodelist = xmlResponse.documentElement.selectNodes("/soap:Envelope/soap:Body/Provisionar_CostosResponse")
			For Each objItem In xnodelist   
				ResultCode =  objItem.selectSingleNode("Provisionar_CostosResult").Text
			Next
		Case 1 'Get_Alerta
			Set xnodelist = xmlResponse.documentElement.selectNodes("/soap:Envelope/soap:Body/Get_AlertaResponse")
			For Each objItem In xnodelist   
				ResultCode =  objItem.selectSingleNode("Get_AlertaResult").Text
			Next			
		Case 2 'Validar_Cobros_Pendientes
			Set xnodelist = xmlResponse.documentElement.selectNodes("/soap:Envelope/soap:Body/Validar_Cobros_PendientesResponse")
			For Each objItem In xnodelist   
				ResultCode =  objItem.selectSingleNode("Validar_Cobros_PendientesResult").Text
			Next
		End Select
		set xmlResponse = nothing
	Else
		ResultCode = StatusCode
	End if
	SetBAWProvition = ResultCode
End Function

Function SetCountryBAW(Country)
	Select Case Country
	Case "GT"
		SetCountryBAW = 1
	Case "SV"
		SetCountryBAW = 2
	Case "HN"
		SetCountryBAW = 3
	Case "NI"
		SetCountryBAW = 4
	Case "CR"
		SetCountryBAW = 5
	Case "PA"
		SetCountryBAW = 6
	Case "BZ"
		SetCountryBAW = 7
	Case "N1"
		SetCountryBAW = 11
    Case "GTLTF"
		SetCountryBAW = 15
    Case "SVLTF"
		SetCountryBAW = 26
    Case "HNLTF"
		SetCountryBAW = 23
    Case "NILTF"
		SetCountryBAW = 24
    Case "CRLTF"
		SetCountryBAW = 21
    Case "PALTF"
		SetCountryBAW = 25
    Case "BZLTF"
		SetCountryBAW = 22    

    Case "GTTLA"
		SetCountryBAW = 32
	Case "SVTLA"
		SetCountryBAW = 38
	Case "HNTLA"
		SetCountryBAW = 35
	Case "NITLA"
		SetCountryBAW = 36
	Case "CRTLA"
		SetCountryBAW = 33
	Case "PATLA"
		SetCountryBAW = 34
	Case "BZTLA"
		SetCountryBAW = 37

    Case "GTLGX"
		SetCountryBAW = 39
	Case "CRLGX"
		SetCountryBAW = 40
	Case "PALGX"
		SetCountryBAW = 41

    Case Else
		SetCountryBAW = 0
	End Select
End Function

Function FormatHour(TheTime)
Dim LenTime
	LenTime = Len(TheTime)
	select Case LenTime
	Case 5
		FormatHour = Left(TheTime,1) & ":" & Mid(TheTime,2,2) & ":" & Right(TheTime,2)
	Case 6
		FormatHour = Left(TheTime,2) & ":" & Mid(TheTime,3,2) & ":" & Right(TheTime,2)
	End Select
End Function


Function CheckCreditClient(ClientID, CountryBAW)
Dim ActualDate, TiempoCredito, MontoCredito, Conn, rs, ShowAlertaCredito
Dim LimitDate, SQLMonto, SQLTiempo, SaldoTotal, SaldoTiempo

    ActualDate = Date 'Esta fecha se usa para comparar creditos vencidos en tiempo
    TiempoCredito = 5 'Tiempo de credito estandar para clientes que no tienen credito (contado)
    MontoCredito = 500 'Monto de credito estandar para clientes que no tienen credito (contado)
    ShowAlertaCredito = 1 'para indicar si muestra o no muestra el credito en el arribo

    openConn2 Conn
	'response.write "select a.ccb_tiempo_autorizado, a.ccb_monto_autorizado, b.incluir_saldo from ((credito_cliente_baw a left join clientes b on a.ccb_id_cliente=b.id_cliente) left join contactos c on b.id_cliente=c.id_cliente) where a.ccb_id_status=5 and ccb_tiempo_autorizado<>0 and a.ccb_monto_autorizado<>0 and a.ccb_pai_id=" & CountryBAW & " and a.ccb_id_cliente=" & ClientID & "<br>"
    Set rs = Conn.Execute("select a.ccb_tiempo_autorizado, a.ccb_monto_autorizado, b.incluir_saldo from ((credito_cliente_baw a left join clientes b on a.ccb_id_cliente=b.id_cliente) left join contactos c on b.id_cliente=c.id_cliente) where a.ccb_id_status=5 and ccb_tiempo_autorizado<>0 and a.ccb_monto_autorizado<>0 and a.ccb_pai_id=" & CountryBAW & " and a.ccb_id_cliente=" & ClientID)
	if Not rs.EOF then
		TiempoCredito = CheckNum(rs(0))
        MontoCredito = CheckNum(rs(1))
        ShowAlertaCredito = CheckNum(rs(2))
	end if
	CloseOBJs rs, Conn

    if ShowAlertaCredito = 1 then
    openConnBAW Conn
       'Restando a la fecha actual los dias de credito para obtener la fecha limite y poder buscar los docs debajo de esa fecha (vencidos)
        LimitDate = replace(ConvertDate(DateAdd("d", -TiempoCredito, ActualDate),2),"/","-",1,-1)
            
	    'Obteniendo Saldo Total de Facturas y NDs de fiscal convertido a dolares y de financiera que ya vienen en dolares
	    SQLMonto = "select coalesce(sum(total)-sum(abono),0) from (" & _
		    "select tfa_cli_id as cliID, tfa_total/tca_tcambio as total, coalesce((select sum(tfr_abono/tca_tcambio) from tbl_factura_abono where tfa_id=tfr_tfa_id and tfr_tfa_sysref_id=1 group by tfr_tfa_id),0) as abono " & _
		    "from tbl_facturacion, tbl_tipo_cambio where date(tca_fecha)=date(tfa_fecha_emision) and tca_pai_id=tfa_pai_id and tfa_conta_id=1 and tfa_ted_id not in (3,4) " & _
		    "and tfa_pai_id=" & CountryBAW & " and tfa_cli_id=" & ClientID & _
		    " union all " & _
		    "select tnd_cli_id as cliID, tnd_total/tca_tcambio as total, coalesce((select sum(tfr_abono/tca_tcambio) from tbl_factura_abono where tnd_id=tfr_tfa_id and tfr_tfa_sysref_id=4 group by tfr_tfa_id),0) as abono " & _
		    "from tbl_nota_debito, tbl_tipo_cambio where date(tca_fecha)=date(tnd_fecha_emision) and tca_pai_id=tnd_pai_id and tnd_tcon_id=1 and tnd_ted_id not in (3,4) and tnd_tpi_id=3 " & _
		    "and tnd_pai_id=" & CountryBAW & " and tnd_cli_id=" & ClientID & _
		    " union all " & _
		    "select tfa_cli_id as cliID, tfa_total as total, coalesce((select sum(tfr_abono) from tbl_factura_abono where tfa_id=tfr_tfa_id and tfr_tfa_sysref_id=1 group by tfr_tfa_id),0) as abono " & _
		    "from tbl_facturacion where tfa_conta_id=2 and tfa_ted_id not in (3,4) " & _
		    "and tfa_pai_id=" & CountryBAW & " and tfa_cli_id=" & ClientID & _
		    " union all " & _
		    "select tnd_cli_id as cliID, tnd_total as total, coalesce((select sum(tfr_abono) from tbl_factura_abono where tnd_id=tfr_tfa_id and tfr_tfa_sysref_id=4 group by tfr_tfa_id),0) as abono " & _
		    "from tbl_nota_debito where tnd_tcon_id=2 and tnd_ted_id not in (3,4) and tnd_tpi_id=3 " & _
		    "and tnd_pai_id=" & CountryBAW & " and tnd_cli_id=" & ClientID & _
		    ") as saldos"

        'Obteniendo Saldo Vencido en Tiempo de Facturas y NDs de fiscal convertido a dolares y de financiera que ya vienen en dolares
	    SQLTiempo = "select coalesce(sum(total)-sum(abono),0) from (" & _
		    "select tfa_cli_id as cliID, tfa_total/tca_tcambio as total, coalesce((select sum(tfr_abono/tca_tcambio) from tbl_factura_abono where tfa_id=tfr_tfa_id and tfr_tfa_sysref_id=1 group by tfr_tfa_id),0) as abono " & _
		    "from tbl_facturacion, tbl_tipo_cambio where date(tca_fecha)=date(tfa_fecha_emision) and tca_pai_id=tfa_pai_id and tfa_conta_id=1 and tfa_ted_id not in (3,4) " & _
		    "and tfa_pai_id=" & CountryBAW & " and tfa_cli_id=" & ClientID & " and tfa_fecha_emision<'" & LimitDate & "'" & _
		    " union all " & _
		    "select tnd_cli_id as cliID, tnd_total/tca_tcambio as total, coalesce((select sum(tfr_abono/tca_tcambio) from tbl_factura_abono where tnd_id=tfr_tfa_id and tfr_tfa_sysref_id=4 group by tfr_tfa_id),0) as abono " & _
		    "from tbl_nota_debito, tbl_tipo_cambio where date(tca_fecha)=date(tnd_fecha_emision) and tca_pai_id=tnd_pai_id and tnd_tcon_id=1 and tnd_ted_id not in (3,4) and tnd_tpi_id=3 " & _
		    "and tnd_pai_id=" & CountryBAW & " and tnd_cli_id=" & ClientID & " and tnd_fecha_emision<'" & LimitDate & "'" & _
		    " union all " & _
		    "select tfa_cli_id as cliID, tfa_total as total, coalesce((select sum(tfr_abono) from tbl_factura_abono where tfa_id=tfr_tfa_id and tfr_tfa_sysref_id=1 group by tfr_tfa_id),0) as abono " & _
		    "from tbl_facturacion where tfa_conta_id=2 and tfa_ted_id not in (3,4) " & _
		    "and tfa_pai_id=" & CountryBAW & " and tfa_cli_id=" & ClientID & " and tfa_fecha_emision<'" & LimitDate & "'" & _
		    " union all " & _
		    "select tnd_cli_id as cliID, tnd_total as total, coalesce((select sum(tfr_abono) from tbl_factura_abono where tnd_id=tfr_tfa_id and tfr_tfa_sysref_id=4 group by tfr_tfa_id),0) as abono " & _
		    "from tbl_nota_debito where tnd_tcon_id=2 and tnd_ted_id not in (3,4) and tnd_tpi_id=3 " & _
		    "and tnd_pai_id=" & CountryBAW & " and tnd_cli_id=" & ClientID & " and tnd_fecha_emision<'" & LimitDate & "'" & _
		    ") as saldos"

        'Obteniendo Saldo Total
        'response.write SQLMonto & "<br>"
        Set rs = Conn.Execute(SQLMonto)
	    if Not rs.EOF then
		    SaldoTotal = Round(rs(0),2) 'Para Monitoreo, Monto pendiente
		end if
	    CloseOBJ rs	
            
        'Obteniendo Saldo vencido en Tiempo
        Set rs = Conn.Execute(SQLTiempo)
	    if Not rs.EOF then
		    SaldoTiempo = Round(rs(0),2) 'Para Monitoreo, Monto pendiente vencido en tiempo
	    end if
	    CloseOBJs rs, Conn

        if SaldoTotal > MontoCredito then
            CheckCreditClient = "<font color=red><b>Estimado cliente, nuestro registro contable muestra un saldo vencido de: USD " & SaldoTotal & _
                                "<br>Sirvase contactar a nuestro depto de Creditos y Cobros, gracias</b></font>" 
        else
            if SaldoTiempo <> 0 then
                CheckCreditClient = "<font color=red><b>Estimado cliente, nuestro registro contable muestra un saldo de: USD " & SaldoTotal & _
                "<br>de los cuales esta vencido USD " & SaldoTiempo & _
                "<br>Sirvase contactar a nuestro depto de Creditos y Cobros, gracias</b></font>"
            end if
        end if
    end if
End Function


Function Base64Encode2(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode2 = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Base64Decode2(ByVal vCode)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.text = vCode
    Base64Decode2 = Stream_BinaryToString(oNode.nodeTypedValue)
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Private Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = adTypeText
  BinaryStream.CharSet = "us-ascii"
  BinaryStream.Open
  BinaryStream.WriteText Text
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary
  BinaryStream.Position = 0
  Stream_StringToBinary = BinaryStream.Read
  Set BinaryStream = Nothing
End Function

Private Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = adTypeBinary
  BinaryStream.Open
  BinaryStream.Write Binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText
  BinaryStream.CharSet = "us-ascii"
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function


'Function Base64Encode(inData)
'  'rfc1521
'  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
'  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
'  Dim cOut, sOut, I
'  'For each group of 3 bytes
'  For I = 1 To Len(inData) Step 3
'    Dim nGroup, pOut, sGroup
'    'Create one long from this 3 bytes.
'    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _
'      &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
'    'Oct splits the long To 8 groups with 3 bits
'    nGroup = Oct(nGroup)
'    'Add leading zeros
'    nGroup = String(8 - Len(nGroup), "0") & nGroup
'    'Convert To base64
'    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
'      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
'      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
'      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
'    'Add the part To OutPut string
'    sOut = sOut + pOut
'    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
'    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
'  Next
'  Select Case Len(inData) Mod 3
'    Case 1: '8 bit final
'      sOut = Left(sOut, Len(sOut) - 2) + "=="
'    Case 2: '16 bit final
'      sOut = Left(sOut, Len(sOut) - 1) + "="
'  End Select
'  Base64Encode = sOut
'End Function

Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function



Function WsGetLogo(pais_iso, sistema,  doc_id,  titulo,  edicion)
        Dim SOAPRequest, responseText
        SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
          "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
            "<soap:Body>" &_
                "<GetLogoData xmlns=""http://tempuri.org/"">" & _
                    "<pais_iso>" & pais_iso & "</pais_iso>" & _
                    "<sistema>" & sistema & "</sistema>" & _
                    "<doc_id>" & doc_id & "</doc_id>" & _
                    "<titulo>" & titulo & "</titulo>" & _
                    "<edicion>" & edicion & "</edicion>" & _
                "</GetLogoData>" & _
            "</soap:Body>" &_
          "</soap:Envelope>"
          WsGetLogo = WsGetParams(SOAPRequest, "GetLogoData", 1)
End Function

Function WsSendMails(pais_iso, to_,  subject,  body,  fromName,  sistema,  user, ip)
    Dim SOAPRequest, responseText
    SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
      "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
        "<soap:Body>" &_
            "<SendMail xmlns=""http://tempuri.org/"">" & _
                "<pais_iso>" & pais_iso & "</pais_iso>" & _
                "<to>" & to_ & "</to>" & _
                "<subject>" & subject & "</subject>" & _
                "<body>" & body & "</body>" & _
                "<fromName>" & fromName & "</fromName>" & _
                "<sistema>" & sistema & "</sistema>" & _
                "<user>" & user & "</user>" & _
                "<ip>" & ip & "</ip>" & _
            "</SendMail>" & _
        "</soap:Body>" &_
      "</soap:Envelope>"
      WsSendMails = WsGetParams(SOAPRequest, "SendMail", 1)
End Function




Function WsModeDev(sistema, countries, metodo, ws21)

    Dim SOAPRequest

    SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
      "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
            "<" & metodo & " xmlns=""http://tempuri.org/"">" & _
                "<Sistema>" & sistema & "</Sistema>" & _
                "<Countries>" & countries & "</Countries>" & _
            "</" & metodo & ">" & _
        "</soap:Body>" & _
      "</soap:Envelope>"

     WsModeDev = WsGetParams(SOAPRequest, metodo, ws21)

End Function

Function WsNotification(tracking_id, sub_product, bl_id, impex, status_id, produccion, user, ip, ws21)
    
    Dim SOAPRequest

    'response.write "(" & tracking_id & ")(" & sub_product & ")(" & bl_id & ")(" & impex & ")(" & status_id & ")(" & user & ")(" & ip & ")<br>"

    SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
      "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
        "<soap:Body>" &_
            "<Notification xmlns=""http://tempuri.org/"">" & _
                "<tracking_id>" & tracking_id & "</tracking_id>" & _
                "<product>" & "1" & "</product>" & _
                "<sub_product>" & sub_product & "</sub_product>" & _
                "<impex>" & impex & "</impex>" & _
                "<bl_id>" & bl_id & "</bl_id>" & _
                "<status_id>" & status_id & "</status_id>" & _
                "<produccion>" & produccion & "</produccion>" & _
                "<user>" & user & "</user>" & _
                "<ip>" & ip & "</ip>" & _
            "</Notification>" & _
        "</soap:Body>" &_
      "</soap:Envelope>"

      WsNotification = WsGetParams(SOAPRequest, "Notification", ws21)
End Function






Function WsExactusSetPedidos(bl_id, impex, bodega, actividad, condicionpago, observaciones, user, ip, ws21, abierto, cliente_id, rubro_id, dua, pedido_erp)

    Dim SOAPRequest, responseText      
    
    SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
      "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
        "<soap:Body>" & _
            "<ExactusSetPedidos xmlns=""http://tempuri.org/"">" & _
                "<producto>" & "1" & "</producto>" & _
                "<bodega>" & bodega & "</bodega>" & _
                "<actividad>" & actividad & "</actividad>" & _
                "<condicionpago>" & condicionpago & "</condicionpago>" & _
                "<impex>" & impex & "</impex>" & _
                "<bl_id>" & bl_id & "</bl_id>" & _                
                "<observaciones>" & observaciones & "</observaciones>" & _                
                "<user>" & user & "</user>" & _
                "<ip>" & ip & "</ip>" & _
                "<Countries>" & Session("OperatorCountry") & "</Countries>" & _
                "<abierto>" & abierto & "</abierto>" & _
                "<cliente_id>" & cliente_id & "</cliente_id>" & _            
                "<rubro_id>" & rubro_id & "</rubro_id>" & _            
                "<dua>" & dua & "</dua>" & _            
                "<pedido_erp>" & pedido_erp & "</pedido_erp>" & _            
            "</ExactusSetPedidos>" & _
        "</soap:Body>" &_
      "</soap:Envelope>"

      WsExactusSetPedidos = WsGetParams(SOAPRequest, "ExactusSetPedidos", ws21)

End Function


Function WsExactusCatalogos(NombreCatalogo, NombreEsquema, ws21)

    Dim SOAPRequest, responseText

    SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
      "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
        "<soap:Body>" & _
            "<ExactusCatalogos xmlns=""http://tempuri.org/"">" & _
                "<NombreCatalogo>" & NombreCatalogo & "</NombreCatalogo>" & _
                "<NombreEsquema>" & NombreEsquema & "</NombreEsquema>" & _
            "</ExactusCatalogos>" & _
        "</soap:Body>" &_
      "</soap:Envelope>"

    SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?><soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""><soap:Body><ExactusCatalogos xmlns=""http://tempuri.org/""><NombreCatalogo>" & NombreCatalogo & "</NombreCatalogo></ExactusCatalogos></soap:Body></soap:Envelope>"

    WsExactusCatalogos = WsGetParams(SOAPRequest, "ExactusCatalogos", ws21)

End Function



Function WsSendMailAttachOne(pais_iso, to1, subject, body, fromName, sistema, user, ip, cc, bc, filename, file64, empresa)

    Dim SOAPRequest       
    
    SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
      "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
        "<soap:Body>" & _
            "<SendMailAttachOne xmlns=""http://tempuri.org/"">" & _
                "<pais_iso>" & pais_iso & "</pais_iso>" & _
                "<to>" & to1 & "</to>" & _
                "<subject>" & subject & "</subject>" & _
                "<body>" & body & "</body>" & _
                "<fromName>" & fromName & "</fromName>" & _
                "<sistema>" & sistema & "</sistema>" & _                
                "<user>" & user & "</user>" & _
                "<ip>" & ip & "</ip>" & _
                "<cc>" & cc & "</cc>" & _                
                "<bc>" & bc & "</bc>" & _                
                "<filename>" & filename & "</filename>" & _                
                "<file64>" & file64 & "</file64>" & _                
                "<empresa>" & empresa & "</empresa>" & _                
            "</SendMailAttachOne>" & _
        "</soap:Body>" &_
      "</soap:Envelope>"

      WsSendMailAttachOne = WsGetParams(SOAPRequest, "SendMailAttachOne", "1")

End Function



Function WsExactus_TIPO_DOC_CP(NombreCatalogo, usuario)

    Dim SOAPRequest

    SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
      "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
        "<soap:Body>" & _
            "<TIPO_DOC_CP xmlns=""http://tempuri.org/"">" & _
                "<NombreCatalogo>" & NombreCatalogo & "</NombreCatalogo>" & _
                "<usuario>" & usuario & "</usuario>" & _
            "</TIPO_DOC_CP>" & _
        "</soap:Body>" &_
      "</soap:Envelope>"

    WsExactus_TIPO_DOC_CP = WsGetParams(SOAPRequest, "TIPO_DOC_CP", "1")

End Function



Function WsEvaluaPedidos(HBLNumber, ObjectID, sistema, CountryExactus, pedido2str)

    Dim SOAPRequest       
    
    SOAPRequest = "<?xml version=""1.0"" encoding=""utf-8""?>" &_
      "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
        "<soap:Body>" & _
            "<EvaluaPedidos xmlns=""http://tempuri.org/"">" & _
                "<HBLNumber>" & HBLNumber & "</HBLNumber>" & _
                "<ObjectID>" & ObjectID & "</ObjectID>" & _
                "<sistema>" & sistema & "</sistema>" & _
                "<CountryExactus>" & CountryExactus & "</CountryExactus>" & _
                "<pedido2str>" & pedido2str & "</pedido2str>" & _
            "</EvaluaPedidos>" & _
        "</soap:Body>" &_
      "</soap:Envelope>"

      WsEvaluaPedidos = WsGetParams(SOAPRequest, "EvaluaPedidos", "1")

End Function





Function WsGetParams(SOAPRequest, metodo, ws21)

    Dim oXmlHTTP, responseText, ResultCode, soapServerMail      

    On Error Resume Next

    'if CheckNum(ws21) = 0 then  
        'soapServerMail = "10.10.1.21:9093" 'pruebas_ 
        'soapServerMail = "localhost:81" 'IIS 
        soapServerMail = "localhost:4343" 'localhost 
    'else    
        'soapServerMail = "10.10.1.21:7480"  'produccion
        'soapServerMail = "10.10.1.21:9093" 'pruebas_ 
    'end if

    Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
    oXmlHTTP.Open "POST", "http://" & soapServerMail & "/SendParametros.asmx", False	    
    oXmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8" 
    oXmlHTTP.send SOAPRequest    
    responseText = ""
    ResultCode = oXmlHTTP.Status
    If ResultCode = 200 Then ' Response from server was success
		responseText = oXmlHTTP.responseText
	End If

    Dim xmlResponse, xnodelist, objItem, Result(23)
    Result(1) = ""
    Result(21) = ""
     
	If Len(responseText) <> 0 Then		
        Set xmlResponse = CreateObject("Microsoft.XMLDOM")  
		xmlResponse.async = false  
        xmlResponse.setProperty "ServerHTTPRequest", True
		xmlResponse.loadXml responseText  
        Set xnodelist = xmlResponse.SelectNodes("/soap:Envelope/soap:Body/" & metodo & "Response/" & metodo & "Result")                
        For Each objItem In xnodelist

            if metodo = "SendMail" or metodo = "Notification" or metodo = "SendMailAttachOne" or metodo = "TIPO_DOC_CP" then  

                Result(0) = objItem.SelectSingleNode("stat").text

                if metodo = "Notification" then
                    Result(1) = objItem.SelectSingleNode("msg").text
                else 
                    if objItem.SelectSingleNode("stat").text = "1" then
                        Result(1) = "<font color=green>" & objItem.SelectSingleNode("msg").text & "</font><br>"    
                    else
                        Result(1) = "<font color=red>" & objItem.SelectSingleNode("msg").text & "</font><br>"    
                    end if
                end if
	    	
            end if

           if metodo = "ExactusCatalogos" then  
                Result(0) = 1
                Result(1) = objItem.text
           end if

           if metodo = "ExactusSetPedidos" then  
                Result(0) = objItem.SelectSingleNode("stat").text
                Result(1) = objItem.SelectSingleNode("msg").text
                Result(2) = objItem.SelectSingleNode("pedido_erp").text
                Result(3) = objItem.SelectSingleNode("log").text
                Result(4) = objItem.SelectSingleNode("error").text
           end if

           if metodo = "EvaluaPedidos" then  
                Result(0) = objItem.SelectSingleNode("stat").text
                Result(1) = objItem.SelectSingleNode("msg").text
                Result(2) = objItem.SelectSingleNode("pedido_erp").text
                Result(3) = objItem.SelectSingleNode("tipo_conta").text
                Result(4) = objItem.SelectSingleNode("esquema").text
           end if

           if metodo = "ModeDev" or metodo = "ModeDev2" then  
                Result(0) = objItem.SelectSingleNode("stat").text
                Result(1) = objItem.SelectSingleNode("msg").text
            end if

            if metodo = "Notification" then  
                Result(2) = objItem.SelectSingleNode("sent_si").text
                Result(3) = objItem.SelectSingleNode("sent_no").text
                Result(4) = objItem.SelectSingleNode("tracking_id").text
                Result(5) = objItem.SelectSingleNode("product").text
                Result(6) = objItem.SelectSingleNode("sub_product").text
                Result(7) = objItem.SelectSingleNode("impex").text
                Result(8) = objItem.SelectSingleNode("bl_id").text
                Result(9) = objItem.SelectSingleNode("status_id").text
                Result(10) = objItem.SelectSingleNode("produccion").text
                Result(11) = objItem.SelectSingleNode("user").text
                Result(12) = objItem.SelectSingleNode("ip").text
                Result(13) = objItem.SelectSingleNode("Countries").text
                Result(14) = objItem.SelectSingleNode("CountriesDest").text
	    	end if


            if metodo = "GetLogoData" then            
                'response.write "(" & responseText & ")"
                Result(0) = objItem.SelectSingleNode("country").text                
                If Not IsNull(objItem.SelectSingleNode("observaciones").text) Then                 
                    Result(1) = objItem.SelectSingleNode("observaciones").text                    
                End if
                Result(2) = objItem.SelectSingleNode("edicion").text
                Result(3) = objItem.SelectSingleNode("titulo").text
                Result(4) = objItem.SelectSingleNode("nombre_empresa").text
                Result(5) = objItem.SelectSingleNode("nit").text
                Result(6) = objItem.SelectSingleNode("direccion").text
                Result(7) = objItem.SelectSingleNode("telefonos").text
                'Result(8) = objItem.SelectSingleNode("trackactivo").text
                'Result(9) = objItem.SelectSingleNode("trackpuerto").text
                'Result(10) = objItem.SelectSingleNode("trackmailserver").text
                'Result(11) = objItem.SelectSingleNode("trackauth").text
                'Result(12) = objItem.SelectSingleNode("trackfromaddress").text
                'Result(13) = objItem.SelectSingleNode("trackpassword").text
                Result(14) = objItem.SelectSingleNode("home_page").text
                Result(15) = objItem.SelectSingleNode("firma").text
                'Result(16) = objItem.SelectSingleNode("fact_elect_codigo").text
                'Result(17) = objItem.SelectSingleNode("fact_elect_user").text
                'Result(18) = objItem.SelectSingleNode("fact_elect_pass").text
                Result(19) = objItem.SelectSingleNode("logo").text 'binario
                Result(20) = "<img src='data:image/jpeg;base64," & objItem.SelectSingleNode("logo2").text & "'>"
                If Not IsNull(objItem.SelectSingleNode("error").text) Then                 
                    Result(21) = objItem.SelectSingleNode("error").text                    
                End if
                Result(22) = objItem.SelectSingleNode("descripcion").text
            end if           
        Next       
		Set xmlResponse = nothing
		Set xnodelist = nothing
    End if 

    If Err.Number <> 0 Then
        Result(21) = Err.description          

        response.write "<br>WebService " & metodo & " " & soapServerMail & " Error : " & Err.Number & " - " & Err.description & "<br>"  
        
    end if
    
    WsGetParams = Result

End Function


Function SendMail(Message, ToAddress, Subject, Country)
    Dim body, result 
    body = Base64Encode(Message)
    result = WsSendMails(Country, ToAddress,  Subject,  body,  "",  "AEREO", Session("Login"), Request.ServerVariables("REMOTE_ADDR"))
    response.write result(1)
End Function




Function SendMail2(Message, ToAddress, Subject, aTableValues5)

    On Error Resume Next

    Dim iConf, Mailer, iFromAddress, iPassword, iPuerto, iMailServer, iAuth, iActivo 
    
    iPuerto = Iif(isNull(aTableValues5(13,0)),0,aTableValues5(13,0))
    iMailServer = Iif(isNull(aTableValues5(14,0)),"",aTableValues5(14,0))
    iAuth = Iif(isNull(aTableValues5(15,0)),0,aTableValues5(15,0))
    iFromAddress = Iif(isNull(aTableValues5(16,0)),"",aTableValues5(16,0))
    iPassword = Iif(isNull(aTableValues5(17,0)),"",aTableValues5(17,0))
    iActivo = Iif(isNull(aTableValues5(12,0)),0,aTableValues5(12,0))
    
    'response.write "iFromAddress :" & iFromAddress & "; iPassword :" & iPassword & "; Port :" & iPuerto & "; iMailServer :" & iMailServer & "; iAuth :" & iAuth & "; iActivo :" & iActivo & ";<br>"
            
    SendMail = -3
            
    if iActivo = 1 and iMailServer <> "" and iFromAddress <> "" then 'trackactivo      
    
        if iAuth = 1 then 
            iAuth = cdoBasic 
        end if

        SendMail = -2   
    
        Set iConf = CreateObject("CDO.Configuration")
        With iConf.Fields  
            .Item(cdoSendUsingMethod)       = cdoSendUsingPort
            .Item(cdoSMTPServer)            = iMailServer
		    .Item(cdoSMTPServerPort)        = iPuerto
		    .Item(cdoSMTPConnectionTimeout) = 10
		    .Item(cdoSMTPAuthenticate)      = iAuth
		    .Item(cdoSendUserName)          = iFromAddress
		    .Item(cdoSendPassword)          = iPassword
            .Update  
        End With 

	    Set Mailer = CreateObject("CDO.Message")
	        Mailer.Configuration = iConf
		    Mailer.From = iFromAddress
    	    Mailer.To = ToAddress
	        Mailer.Subject = Subject
    	    Mailer.HTMLBody = Message
        
            SendMail = Mailer.Send
    
    else
            'response.write "<font color=red> Servidor smtp sin configuracion.</font><br>"
    
    end if

    If Err.Number <> 0 Then
        response.write "SendMail " & " To : " & ToAddress &  "; Error : " & Err.description & "<br>" '" & Err.Number & " 

        response.write "MailServer : " & iMailServer & "; Auth : " & iAuth & "; Activo : " & iActivo & "; From : " & iFromAddress & "; Port : " & iPuerto & "<br><br>" 'Pass : " & iPassword & "; 

        SendMail = -1

        Err.Number = 0 
    end if

    Set Mailer = Nothing
	Set iConf = Nothing            
    
End Function



'Sub SendMail_2(Message, eMails, Subject, FromAddress)
'Dim Mailer
'Dim iConf
'    Set iConf = CreateObject("CDO.Configuration") 
'    With iConf.Fields  
'        .Item(cdoSendUsingMethod) = cdoSendUsingPort  
'        .Item(cdoSMTPServer) = "200.119.132.163"
'        '.Item(cdoSMTPServer) = "mail.aimargroup.com"
'		'.Item(cdoSMTPServerPort)        = 25
'		'.Item(cdoSMTPConnectionTimeout) = 10
'		'.Item(cdoSMTPAuthenticate)      = cdoBasic
'		'.Item(cdoSendUserName)          = "username"
'		'.Item(cdoSendPassword)          = "password"
'        .Update  
'    End With  
'	Set Mailer = CreateObject("CDO.Message")
'	    Mailer.Configuration = iConf
'		Mailer.From = FromAddress
'   	Mailer.To = eMails
'	    Mailer.Subject = Subject
'    	Mailer.HTMLBody = Message
'	    Mailer.Send
'	Set Mailer = Nothing
'	Set iConf = Nothing

'Dim Mailer
'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'		Mailer.Charset = 2 'Utilizamos UNICODE
'		Mailer.ContentType = "text/html"
'		Mailer.FromName = FromName
'		Mailer.FromAddress = FromAddress 
'		Mailer.Subject = Subject
'		Mailer.BodyText = Message
'		Mailer.RemoteHost = IP_SMTP'"mail-fwd.oemgrp.com"
'		Mailer.AddRecipient "", eMails
'		Mailer.SendMail
'Set Mailer = Nothing

'Set Mailer = Server.CreateObject("CDONTS.NewMail")
'Set Mailer = Server.CreateObject("CDO.Message")
'    Mailer.From = FromName
'    Mailer.To = eMails
'    Mailer.Subject = Subject
'    Mailer.HTMLBody = Message
'    Mailer.Send
'Set Mailer = Nothing
'End Sub

'Function SendMails(Message, eMails, Subject, FromAddress)    
'    Dim Mailer
'    Dim iConf
'    SendMails = -1   
'    On Error Resume Next       
'    Set iConf = CreateObject("CDO.Configuration")
'    With iConf.Fields  
'        .Item(cdoSendUsingMethod)       = cdoSendUsingPort
'        .Item(cdoSMTPServer)            = "mail.aimargroup.com"
'		.Item(cdoSMTPServerPort)        = 25
'		.Item(cdoSMTPConnectionTimeout) = 10
'		.Item(cdoSMTPAuthenticate)      = cdoBasic
'		.Item(cdoSendUserName)          = FromAddress
'		.Item(cdoSendPassword)          = "Lq14@6A8"
'       .Update  
'    End With 
'
'	Set Mailer = CreateObject("CDO.Message")
'	    Mailer.Configuration = iConf
'		Mailer.From = FromAddress
'    	Mailer.To = eMails
'	    Mailer.Subject = Subject
'    	Mailer.HTMLBody = Message       
'       SendMails = Mailer.Send
'    If Err.Number <> 0 Then
'        response.write  Err.Number & " " & Err.description
'    'else
'    '    SendMails = 0
'    end if
'    Set Mailer = Nothing
'	Set iConf = Nothing               
'End Function


Sub SendMails2(Message, eMails, Subject, FromAddress, result)    

    Dim Mailer
    Dim iConf

    On Error Resume Next    

    Set iConf = CreateObject("CDO.Configuration")
    With iConf.Fields  
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Item(cdoSMTPServer)            = "mail.aimargroup.com"
		.Item(cdoSMTPServerPort)        = 25
		.Item(cdoSMTPConnectionTimeout) = 10
		.Item(cdoSMTPAuthenticate)      = cdoBasic
		.Item(cdoSendUserName)          = FromAddress
		.Item(cdoSendPassword)          = "Lq14@6A8"
        .Update  
    End With     

	Set Mailer = CreateObject("CDO.Message")
	    Mailer.Configuration = iConf
		Mailer.From = FromAddress
    	Mailer.To = eMails
	    Mailer.Subject = Subject
    	Mailer.HTMLBody = Message
	    result = Mailer.Send

	Set Mailer = Nothing
	Set iConf = Nothing

End Sub


Function CheckHAWBNumber(Conn, AwbType, HAWBNumber) 
    Dim rs, AwbTable

    CheckHAWBNumber = 0
    if AwbType =  1 then
        AwbTable = "Awb"
    else
        AwbTable = "Awbi"
    end if

    Set rs = Conn.Execute("select HAWBNumber from " & AwbTable & " where HAWBNumber='" & HAWBNumber & "'")
    if Not rs.EOF then
        CheckHAWBNumber = 1
    end if
    CloseOBJ rs    
End Function

Function CheckAWBNumber(Conn, AWBNumber) 
    Dim rs
    CheckAWBNumber = 0    
    Set rs = Conn.Execute("SELECT GuideID FROM Guides WHERE GuideNumber = '" & AWBNumber & "'")
    if Not rs.EOF then
        CheckAWBNumber = 1
    end if
    CloseOBJ rs   
End Function

Function NextAWBNumber(Conn, AwbType, CarrierID, TipoMaster)
    
    NextAWBNumber = ""


    if CarrierID > -1 then
    

    Select Case TipoMaster

    Case "Nuevo" ' (GuideStatus 0 sin awb / 1 sin awb) (GuideType 0 no; iniciada 1 iniciada) (GuideActive 1:on 0:off)
        'response.write "SELECT GuideNumber FROM Guides WHERE GuideStatus='0' AND GuideActive = '1' AND GuideCarrierID = '" & CarrierID & "'  ORDER BY GuideNumber LIMIT 0,1"

        SQLQuery = "SELECT GuideNumber FROM Guides WHERE GuideType='0' AND GuideStatus='0' AND GuideActive='1' AND GuideCarrierID = '" & CarrierID & "'  ORDER BY GuideNumber LIMIT 0,1"

        'response.write SQLQuery & "<br>"

        Set rs = Conn.Execute(SQLQuery) 
	    If Not rs.EOF Then
		    NextAWBNumber = rs(0)            
	    end if
	    CloseOBJ rs		
        
        dim conteo
        conteo = 0                
        Set rs = Conn.Execute("SELECT COUNT(GuideNumber) FROM Guides WHERE GuideType='0' AND GuideStatus='0' AND GuideActive='1' AND GuideCarrierID = '" & CarrierID & "'")
		If Not rs.EOF Then
			conteo = cInt(rs(0))
		end if
        CloseOBJ rs

        if conteo < 6 then            
            dim mensaje, result
            result = -1            
            
            dim countries
            countries = ""
            carrierName = ""                
	        Set rs = Conn.Execute("select CarrierID, Name, Countries from Carriers where CarrierID = '" & CarrierID & "'")
	        If Not rs.EOF Then
   		        carrierName = rs(1)
                countries = rs(2)
            End If
	        CloseOBJ rs		

            mensaje = "Estimado usuario " & Session("OperatorName") & ":<br>"            
            if conteo = 0 then
                mensaje = mensaje & "No hay guias para la aerolinea " & carrierName & "<br>"
            else
                mensaje = mensaje & "Le quedan " & conteo & " guias para la aerolinea " & carrierName & "<br>"
            end if            
            'response.write( "<" & "script" & ">" & "alert('" & "AVISO: Quedan " & conteo & " guias" & "');" & "<" & "/script" & ">" )
            
            response.write( "<" & "script" & ">" & "alert('" & Replace(mensaje,"<br>","\r\n") & "');" & "<" & "/script" & ">" )
            
            mensaje = mensaje & "Esta es una auto notificacion del sistema aereo<br>"
            
            result = SendMail(mensaje, Session("OperatorEmail"), "Aereo Guias " & carrierName, countries)
            
        end if
                
    Case "Ultimo"        
        dim sql, AwbTable
        if AwbType = 1 then
            AwbTable = "Awb"
        else
            AwbTable = "Awbi"
        end if
        'sql = "Select a.AWBNumber from Awb a where a.AWBID = SELECT MAX(b.AWBID) from Awb b, Carriers c where b.Countries=c.Countries and c.CarrierID=" & CarrierID  & ""
        'sql = "SELECT b.AWBNumber from Awb b, Carriers c where b.Countries=c.Countries and c.CarrierID=" & CarrierID  & " ORDER BY b.AWBID DESC LIMIT 1"
        'sql = "SELECT AWBNumber FROM " & AwbTable & " WHERE CarrierID = " & CarrierID  & " ORDER BY AwbID DESC LIMIT 1"
        sql = "SELECT AWBNumber FROM " & AwbTable & " WHERE AWBID = (SELECT MAX(AWBID) FROM Awb WHERE AwbNumber <> '' AND CarrierID=" & CarrierID  & ")"
        'response.write(sql)
		Set rs = Conn.Execute(sql)
		If Not rs.EOF Then
			NextAWBNumber = rs(0)
		end if
		CloseOBJ rs	

    Case "Espacio"
        'response.write("(" & trim(AWBNumber) & ")")
        if " " & trim(AWBNumber) = AWBNumber then
            NextAWBNumber = trim(AWBNumber)
        else
            NextAWBNumber = " " & trim(AWBNumber)
        end if
             
    End Select

    else

    'response.write( "<" & "script" & ">" & "alert('Seleccione Transportista');" & "<" & "/script" & ">" )
  		
    end if

End Function


Function NextHAWBNumber(HAWBNumber, Conn, AwbType, Countries, TipoMaster, TipoHouse, AWBNumber)
    
    'response.write("<br>(TipoMaster=" & TipoMaster & ")(TipoHouse=" & TipoHouse & ")(AWBNumber=" & AWBNumber & "-" & AWBNumberAnt & ") HAWBNumber=(" & HAWBNumber & "-" & HAWBNumberAnt & ")")
    
    dim AwbTable, sql

    if AwbType = 1 then
        AwbTable = "Awb"
    else
        AwbTable = "Awbi"
    end if

    'if TipoMaster = "Nuevo" then
    '    NextHAWBNumber = ""
    '    exit function
    'end if
    
    Select Case TipoHouse

    Case "Directo" 
        
        sql = "SELECT AWBNumber FROM " & AwbTable & " WHERE HAWBNumber = '" & AWBNumber & "'"
        Set rs = Conn.Execute(sql)
        if Not rs.EOF then 
            'response.write "1(" & HAWBNumber & ")"
            response.write "Ya existe registro directo"
            NextHAWBNumber = HAWBNumber
        else                
            'response.write "2(" & AWBNumber & ")"
            if AWBNumber = "" Then                
                'if AWBNumber <> AWBNumberAnt and AWBNumberAnt <> "" then 'cuando ya se guardo y ha cambiado de linea aerea son distintos debe tomar igual al actual            
                NextHAWBNumber = AsignarHAWBNumber(Conn, Countries, AwbTable)
            else        
                NextHAWBNumber = AWBNumber
            end if            

        end if
        CloseOBJ rs
        
        
    Case "Asignar" 
        
        'If HAWBNumber <> HAWBNumberAnt OR HAWBNumber = "" Then        
            NextHAWBNumber =  AsignarHAWBNumber(Conn, Countries, AwbTable)
        'Else
            'NextHAWBNumber = HAWBNumber
        'End If

    Case "" , "Manual" 

        NextHAWBNumber = HAWBNumber
    
    End Select
    
    'response.write("<br>(NextHAWBNumber=" & NextHAWBNumber & ")")

End Function


Function AsignarHAWBNumber(Conn, Countries, AwbTable)
        dim sql, pref, pais
        AsignarHAWBNumber = ""        
        pref = "AIM"    
        'sql = "SELECT CONCAT('" & pref & "',LPAD(IFNULL(MAX(CAST(SUBSTRING(HAWBNumber,INSTR(HAWBNumber, '" & pref & "')+3, INSTR(HAWBNumber,'" & Countries & "') - (INSTR(HAWBNumber, '" & pref & "')+3)) AS UNSIGNED)),0)+1," & largo & ",'0') ,'" & Countries & "') as Num FROM " & AwbTable & " WHERE HAWBNumber LIKE '" & pref & "%' AND HAWBNumber LIKE '%" & Countries & "'"
        if Countries = "GT" then
            pais = "GUA"            
        else
            pais = Countries
        end if        
        sql = "SELECT CONCAT('" & pref & "',IFNULL(MAX(CAST(REPLACE(SUBSTRING(HAWBNumber,(INSTR(HAWBNumber, '" & pref & "')+3),INSTR(HAWBNumber,'" & pais & "')-(INSTR(HAWBNumber,'" & pref & "')+3)),'-','') AS UNSIGNED)),0)+1,'" & pais & "') as Num FROM " & AwbTable & " WHERE HAWBNumber LIKE '" & pref & "%' AND HAWBNumber LIKE '%" & pais & "'"
        'response.write(sql)
        Set rs = Conn.Execute(sql)
        if Not rs.EOF then
            AsignarHAWBNumber = rs(0)            
        end if
        CloseOBJ rs
End Function




Function PerfilOpcion()     

    Dim Conn, rs, SQLQuery, strString, i, iLab, iKey, a
    
    iLab = Split("Ins,Upd,Aut,Pdf,Exc,Adm,Del,Log",",")
    Set iArr2 = Server.CreateObject("Scripting.Dictionary")    

    SQLQuery = "SELECT um_mn_id, um_permisos FROM contactos_users_menu WHERE um_fields = 'AEREO' AND id_usuario = '" & Session("OperatorID") & "' AND um_st = 'Activo'"     
    'response.write SQLQuery & "<br>"
    OpenConn2 Conn        
    Set rs = Conn.Execute(SQLQuery)    
    If Not rs.EOF Then      
		do while Not rs.EOF
            strString = rs("um_permisos")
            For i = 1 To Len(strString)                                            
                iKey = rs("um_mn_id") & iLab(i-1)
                iArr2.Add iKey,Mid(strString,i,1)

                'Response.Write iKey & "-" & Mid(strString,i,1) & " " 
            Next             
            iKey = rs("um_mn_id") & "Log"
            iArr2.Add iKey, IIf(inStr(1,strString,"1",1) > 0,"1","0") 'Log se asigna con cualquier permiso
        	rs.MoveNext
		loop
    end if
    CloseOBJs rs, Conn

    'a=iArr2.Items
    'for i=0 to iArr2.Count-1
        'iKey = iLab(i)
        'Response.Write(a(i))
        'Response.Write("<br>")
    'next

    'For i = 0 To UBound(iLab)
    'for i=0 to iArr2.Count-1
        'iKey = 'rs("um_mn_id") & iLab(i)
        'iKey = iLab(i)
        'response.write "(" & iKey & ")" '(" & iArr2.Item(iKey) & ")"
    'Next 

    'response.write "(" & iArr2.Item("23Ins") & ")"
    

End Function





Function PerfilOpcion2(GID, OperatorID) 

    SQLQuery = "SELECT " & _  
        "cast(substring(um_permisos,1,1) as boolean) as Ins, " & _
        "cast(substring(um_permisos,2,1) as boolean) as Upd, " & _
        "cast(substring(um_permisos,3,1) as boolean) as Aut, " & _
        "cast(substring(um_permisos,4,1) as boolean) as Pdf, " & _
        "cast(substring(um_permisos,5,1) as boolean) as Exc, " & _
        "cast(substring(um_permisos,6,1) as boolean) as Adm, " & _
        "cast(substring(um_permisos,7,1) as boolean) as Del, 1 as no " & _            
    "FROM contactos_users_menu WHERE um_fields = 'AEREO' AND um_mn_id = '" & GID & "' AND id_usuario = '" & OperatorID & "' AND um_st = 'Activo' " & _
    "UNION " & _
    "SELECT false as Ins, false as Upd, false as Aut, false as Pdf, false as Exc, false as Adm, false as Del, 2 as no " & _
    "ORDER BY no LIMIT 1"
    
    SQLQuery = "SELECT 2 as no, '0000000' as um_permisos UNION SELECT 1 as no, um_permisos FROM contactos_users_menu WHERE um_fields = 'AEREO' AND um_mn_id = '" & GID & "' AND id_usuario = '" & OperatorID & "' AND um_st = 'Activo' ORDER BY no LIMIT 1"
    'response.write SQLQuery & "<br>"

    Dim Conn, rs, SQLQuery
    OpenConn2 Conn
        
    Dim strString, i, iLab
    iLab = Split("Ins,Upd,Aut,Pdf,Exc,Adm,Del",",")

    Set iArr=Server.CreateObject("Scripting.Dictionary")    
    Set rs = Conn.Execute(SQLQuery)    
    If Not rs.EOF Then      
        'response.write rs("um_permisos") & "<br>"
        strString = rs("um_permisos")
        For i = 1 To Len(strString)                        
            'iArr.Add iLab(i-1),IIf(Mid(strString,i,1) = "1",true,false)
            iArr.Add iLab(i-1),Mid(strString,i,1)
        Next 

        iArr.Add "Log",IIf(inStr(1,strString,"1",1) > 0,"1","0") 'Log se asigna con cualquier permiso

    end if
    CloseOBJs rs, Conn

    'For iI = 0 To UBound(iLab)
    '    response.write "(" & iLab(iI) & ")(" & iArr.Item(iLab(iI)) & ")"
    'Next 

End Function




Function InsertGuia(Conn, rs, replica, AwbType, AWBNumber, HAWBNumber2, Country2, Piezas2, Peso2, Transportista2, AirportDepID2, AirportDesID2, iAirportFromCode, iAirportToCode, data1, data2, data3, TipoCarga)

	QuerySelect = "INSERT INTO " & IIf(AwbType = 1,"Awb","Awbi") & " (AwbNumber, HAwbNumber, CreatedDate, CreatedTime, Countries, TotNoOfPieces, NoOfPieces, TotWeight, Weights, ChargeableWeights, CarrierRates, CarrierID, AirportDepID, AirportDesID, AWBDate, AgentContactSignature, AirportToCode1, AirportToCode2, RequestedRouting, flg_totals" & IIf(AwbType = 1,", replica","") & ", WeightsSymbol) VALUES ('" & AWBNumber & "', '" & HAWBNumber2 & "', CURRENT_DATE(), REPLACE(TIME(now()),':',''), '" & Country2 & "', '" & Piezas2 & "', '" & Piezas2 & "', '" & Peso2 & "', '" & Peso2 & "', '" & Peso2 & "', '0', '" & Transportista2 & "', '" & AirportDepID2 & "', '" & AirportDesID2 & "', CURRENT_DATE(), '" & Session("OperatorName") & "', '" & iAirportFromCode & "', '" & iAirportToCode & "', '" & iAirportFromCode & "/" & iAirportToCode & "','1'" & data1 & ", 'KG');" 
	'response.write "(awb_frame1)<br>" & QuerySelect & "<br>"
    Piezas2 = ""
    Peso2 = ""
    Conn.Execute(QuerySelect)

	InsertGuia = 0
	QuerySelect = "SELECT LAST_INSERT_ID()" 
	'response.write "(awb_frame2)<br>" & QuerySelect & "<br>"
	Set rs = Conn.Execute(QuerySelect)                                            
	if Not rs.EOF then                                                
		InsertGuia = CheckNum(rs(0))      
	end if

	if InsertGuia > 0 then

		'QuerySelect = "INSERT INTO Awb_IE_Expansion (aiee_AwbID_fk, aiee_ImpExp, aiee_TipoAwb, aiee_replica, aiee_master_hija) VALUES (" & InsertGuia & "," & AwbType & ", '" & replica & "', '" &  Iif(replica = "Consolidado" or replica = "Master-Hija" or replica = "Master-Master-Hija", "Consolidado", "Directo") & "', 'Hija')" 

		QuerySelect = "INSERT INTO Awb_IE_Expansion (aiee_AwbID_fk, aiee_ImpExp, aiee_TipoAwb, aiee_replica, aiee_master_hija, aiee_TipoCarga) VALUES (" & InsertGuia & "," & AwbType & ", '" & replica & "', '" &  data2 & "', '" & data3 & "', '" & TipoCarga & "')" 
		'response.write "(awb_frame3)<br>" & QuerySelect & "<br>"
		Conn.Execute(QuerySelect)  

	end if

End Function	



	
Sub ReplicarHeaderRubros(Conn, rs, esMaster, AwbType, AWBNumber, HAwbNumber, ObjectID, ObjectIDtmp, ClientCollectID_tmp, esMHHija, peso)

    On Error Resume Next

        Dim c, d

        ClientCollectID_tmp = ""

	    'siempre debe buscar el id de la master-hija
        QuerySelect = "SELECT MIN(c.AwbID) as AwbID, c.HAwbNumber, c.ConsignerData, c.ShipperData, c.AgentData FROM " & IIf(AwbType = 1,"Awb","Awbi") & " c WHERE c.AwbNumber = '" & AWBNumber & "' AND c.HAwbNumber <> '' AND c.Expired = '0'"           
        'response.write QuerySelect & "<br>"
        Set rs = Conn.Execute(QuerySelect)
        If Not rs.EOF Then

            ObjectIDtmp = rs("AwbID") 'guia master-hija

            if esMaster = True then         'si el registro actual es Master

                'si registro actual es la  master, actualiza la master-hija guia y rubros
                if IsNull(rs("ConsignerData")) and IsNull(rs("ShipperData")) and IsNull(rs("AgentData")) then 
                    ClientCollectID_tmp = "0"   'actualiza guia y rubros
                else
                    ClientCollectID_tmp = "1"   'actualiza solo rubros
                end if

            end if

            'si el registro actual es la master-hija, actualizara rubros a las hijas
            if esMaster = False and rs("HAwbNumber") = HAwbNumber then 
                ClientCollectID_tmp = "3"   'actualiza solo rubros
            end if

            'si el registro actual es MH-Hija debe replicarse con la M-H
            if esMaster = False and esMHHija = True and rs("HAwbNumber") <> HAwbNumber then 
            
                'estos datos son la la M-H por lo cual ya vienen con valores
                'si esMHHija = True es suficiente para actualizar header
                'if IsNull(rs("ConsignerData")) and IsNull(rs("ShipperData")) and IsNull(rs("AgentData")) then 
                    ClientCollectID_tmp = "4"   'actualiza guia y rubros
                'else
                '    ClientCollectID_tmp = "5"   'actualiza solo rubros
                'end if

            end if


        End If
        CloseOBJ rs

    
        'response.write "(" & ClientCollectID_tmp & ")(" & ObjectID & ")(" & ObjectIDtmp & ")<br>"

    
        if ObjectIDtmp <> "" then '////////////////////////////////////// MODULO ACTUALIZAR RUBROS

            Select Case ClientCollectID_tmp
            Case "0", "1"    'si registro actual es la  master, actualiza la master-hija guia y rubros

                d = ReplicaRubros (Conn, 0, ObjectID, AwbType, ObjectIDtmp, esMaster, esMHHija, peso)
              
                if d <> 4 then
                    response.write "(d=" & d & ") Problemas al replicar rubros, por favor intente nuevamente.<br>"
                end if

            Case "4", "5"    'solo debe tomar el id actual de la MH-Hija

                c = ObjectID
                ObjectID = ObjectIDtmp  'guia master-hija
                ObjectIDtmp = c         'guia MH-Hija actual

                d = ReplicaRubros (Conn, 0, ObjectID, AwbType, ObjectIDtmp, esMaster, esMHHija, peso)
        
                if d <> 4 then
                    response.write "(d=" & d & ") Problemas al replicar rubros, por favor intente nuevamente.<br>"
                end if

            Case "3"    'si el registro actual es la master-hija
                
                'debe obtener ids de las hijas y replicar data                        
                QuerySelect = "SELECT c.AwbID, c.HAwbNumber, c.ConsignerData, c.ShipperData, c.AgentData FROM " & IIf(AwbType = 1,"Awb","Awbi") & " c WHERE c.AwbNumber = '" & AWBNumber & "' AND c.HAwbNumber <> '' AND c.Expired = '0' AND c.AwbID <> " & ObjectIDtmp & " "
                'response.write QuerySelect & "<br>"
                Set rs = Conn.Execute(QuerySelect)

                ObjectIDtmp = ""
                c = 0

                If Not rs.EOF Then      
		            do while Not rs.EOF

                        'response.write "(" & IsNull(rs("ConsignerData")) & ")(" & IsNull(rs("ShipperData")) & ")(" & IsNull(rs("AgentData")) & ")<br>" 

                        if IsNull(rs("ConsignerData")) and IsNull(rs("ShipperData")) and IsNull(rs("AgentData")) then 
                            if ObjectIDtmp <> "" then 
                                ObjectIDtmp = ObjectIDtmp & ","
                            end if
                            ObjectIDtmp = ObjectIDtmp & rs("AwbID") 
        
                            ClientCollectID_tmp = "2"   'actualiza guia y rubros

                        end if

                        d = ReplicaRubros (Conn, c, ObjectID, AwbType, rs("AwbID"), esMaster, esMHHija, peso)

                if d <> 4 then
                    response.write "(d=" & d & ") Problemas al replicar rubros, por favor intente nuevamente.<br>"
                end if

                        c = c + 1

        	            rs.MoveNext
		            loop
                end if

                CloseOBJ rs

            End Select	

        end if 

        'response.write "(" & ClientCollectID_tmp & ")(" & ObjectID & ")(" & ObjectIDtmp & ")<br>"

        QuerySelect = ""

        if ObjectID <> ObjectIDtmp and ObjectIDtmp <> "" then '/////////////////// MODULO ACTUALIZAR GUIA


            Dim Mat0Carrier, Mat1Agente, Mat2Otros


            if ClientCollectID_tmp <> "" then 

                '"a.HAwbNumber='" & Request("HAWBNumber2") & "' " & _ 
                '"a.CreatedDate=b.CreatedDate, " & _ 
                '"a.CreatedTime=b.CreatedTime, " & _ 
                '"a.AwbNumber=b.AwbNumber, " & _ 
                '"a.Expired=b.Expired, " & _ 
                '"a.flg_master=b.flg_master, " & _ 
                '"a.flg_totals=b.flg_totals, " & _ 
                '"a.TotWeight=b.TotWeight, " & _ 
                '"a.ConsignerData=b.ConsignerData, " & _ 
                '"a.ConsignerID=b.ConsignerID, " & _ 
                '"a.TotNoOfPieces=b.TotNoOfPieces, " & _  
                '"a.NoOfPieces=b.NoOfPieces, " & _ 
                '"a.Weights=b.Weights, " & _ 
                '"a.ChargeableWeights=b.ChargeableWeights, " & _ 



                QuerySelect = ""

                'if esMHHija = True then 'si el registro actual es hija
                if ClientCollectID_tmp = "4" then
                    QuerySelect = QuerySelect & "a.CarrierRates=b.CarrierRates, " 'este dato solo se replica a las hijas, no a la master-hija
                end if

                'si el registro actual es la master-hija, actualizara tarifa ó se esta generando la nueva M-M-H hija
                'if ClientCollectID_tmp = "2" or ClientCollectID_tmp = "3" or ClientCollectID_tmp = "4" or ClientCollectID_tmp = "5"  then 
                '   QuerySelect = QuerySelect & "a.CarrierRates=b.CarrierRates, " 'este dato solo se replica a las hijas, no a la master-hija
                'end if

                'falta incluir los rubros otros de import

                
                '//////////////////////////ACTUALIZA VALORES DE RUBROS Y OTROS

                if esMaster = True and esMHHija = False then 'debe ser la M-H

                    if Cdbl(peso) > 0 then 

                        'QuerySelect = QuerySelect & RecalculoPorPeso(Conn, rs, ObjectID, Cdbl(peso), AWBType)
                    
                        QuerySelect = QuerySelect & "a.ChargeableWeights=" & CDbl(peso) & ", a.TotWeightChargeable=" & CDbl(peso) & ", a.Carriersubtot=b.CarrierRates*" & CDbl(peso) & ", "     

                    end if

                    QuerySelect = QuerySelect & _ 
                    "a.CarrierRates=0, " & _
                    "a.TotCarrierRate=0, " & _
                    "a.Carriersubtot=0, " 

                    '"a.CarrierRates=CASE WHEN a.CarrierRates > 0 THEN a.CarrierRates ELSE 0 END, " & _
                    '"a.TotCarrierRate=CASE WHEN a.TotCarrierRate > 0 THEN a.TotCarrierRate ELSE 0 END, " & _
                    '"a.Carriersubtot=CASE WHEN a.Carriersubtot > 0 THEN a.Carriersubtot ELSE 0 END, " & _ 

                    'todos los AdditionalChargeVal1 .. 15 no se deben tocar en este bloque
                    QuerySelect = QuerySelect & _ 
                    "a.AdditionalChargeVal1=CASE WHEN CAST(a.AdditionalChargeVal1 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal1 END, " & _ 
                    "a.AdditionalChargeVal2=CASE WHEN CAST(a.AdditionalChargeVal2 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal2 END, " & _ 
                    "a.AdditionalChargeVal3=CASE WHEN CAST(a.AdditionalChargeVal3 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal3 END, " & _ 
                    "a.AdditionalChargeVal4=CASE WHEN CAST(a.AdditionalChargeVal4 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal4 END, " & _ 
                    "a.AdditionalChargeVal5=CASE WHEN CAST(a.AdditionalChargeVal5 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal5 END, " & _ 
                    "a.AdditionalChargeVal6=CASE WHEN CAST(a.AdditionalChargeVal6 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal6 END, " & _ 
                    "a.AdditionalChargeVal7=CASE WHEN CAST(a.AdditionalChargeVal7 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal7 END, " & _ 
                    "a.AdditionalChargeVal8=CASE WHEN CAST(a.AdditionalChargeVal8 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal8 END, " & _ 
                    "a.AdditionalChargeVal9=CASE WHEN CAST(a.AdditionalChargeVal9 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal9 END, " & _ 
                    "a.AdditionalChargeVal10=CASE WHEN CAST(a.AdditionalChargeVal10 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal10 END, " & _ 
                    "a.AdditionalChargeVal11=CASE WHEN CAST(a.AdditionalChargeVal11 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal11 END, " & _ 
                    "a.AdditionalChargeVal12=CASE WHEN CAST(a.AdditionalChargeVal12 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal12 END, " & _ 
                    "a.AdditionalChargeVal13=CASE WHEN CAST(a.AdditionalChargeVal13 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal13 END, " & _ 
                    "a.AdditionalChargeVal14=CASE WHEN CAST(a.AdditionalChargeVal14 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal14 END, " & _ 
                    "a.AdditionalChargeVal15=CASE WHEN CAST(a.AdditionalChargeVal15 AS Decimal) > 0 THEN '' ELSE a.AdditionalChargeVal15 END, " 


                    if AwbType = 2 then 'si es IMPORT agrega estos campos

                        QuerySelect = QuerySelect & _
                        "a.OtherChargeVal1=CASE WHEN CAST(a.OtherChargeVal1 AS Decimal) > 0 THEN '' ELSE a.OtherChargeVal1 END, " & _ 
                        "a.OtherChargeVal2=CASE WHEN CAST(a.OtherChargeVal2 AS Decimal) > 0 THEN '' ELSE a.OtherChargeVal2 END, " & _ 
                        "a.OtherChargeVal3=CASE WHEN CAST(a.OtherChargeVal3 AS Decimal) > 0 THEN '' ELSE a.OtherChargeVal3 END, " & _ 
                        "a.OtherChargeVal4=CASE WHEN CAST(a.OtherChargeVal4 AS Decimal) > 0 THEN '' ELSE a.OtherChargeVal4 END, " & _ 
                        "a.OtherChargeVal5=CASE WHEN CAST(a.OtherChargeVal5 AS Decimal) > 0 THEN '' ELSE a.OtherChargeVal5 END, " & _ 
                        "a.OtherChargeVal6=CASE WHEN CAST(a.OtherChargeVal6 AS Decimal) > 0 THEN '' ELSE a.OtherChargeVal6 END, " 

                    end if

                else 

                    if Cdbl(peso) > 0 then 

                        QuerySelect = QuerySelect & RecalculoPorPeso(Conn, rs, ObjectID, Cdbl(peso), AWBType)
        
                    else

                        QuerySelect = QuerySelect & _
                        "a.CarrierRates=b.CarrierRates, " & _ 
                        "a.TotCarrierRate=b.TotCarrierRate, " & _                    
                        "a.Carriersubtot=b.Carriersubtot, " & _ 
                        "a.AdditionalChargeVal1=b.AdditionalChargeVal1, " & _ 
                        "a.AdditionalChargeVal2=b.AdditionalChargeVal2, " & _ 
                        "a.AdditionalChargeVal3=b.AdditionalChargeVal3, " & _ 
                        "a.AdditionalChargeVal4=b.AdditionalChargeVal4, " & _ 
                        "a.AdditionalChargeVal5=b.AdditionalChargeVal5, " & _ 
                        "a.AdditionalChargeVal6=b.AdditionalChargeVal6, " & _ 
                        "a.AdditionalChargeVal7=b.AdditionalChargeVal7, " & _ 
                        "a.AdditionalChargeVal8=b.AdditionalChargeVal8, " & _ 
                        "a.AdditionalChargeVal9=b.AdditionalChargeVal9, " & _ 
                        "a.AdditionalChargeVal10=b.AdditionalChargeVal10, " & _ 
                        "a.AdditionalChargeVal11=b.AdditionalChargeVal11, " & _ 
                        "a.AdditionalChargeVal12=b.AdditionalChargeVal12, " & _ 
                        "a.AdditionalChargeVal13=b.AdditionalChargeVal13, " & _ 
                        "a.AdditionalChargeVal14=b.AdditionalChargeVal14, " & _ 
                        "a.AdditionalChargeVal15=b.AdditionalChargeVal15, " 



                        if AwbType = 2 then 'si es IMPORT agrega estos campos

                            QuerySelect = QuerySelect & _
                            "a.OtherChargeVal1=b.OtherChargeVal1, " & _ 
                            "a.OtherChargeVal2=b.OtherChargeVal2, " & _ 
                            "a.OtherChargeVal3=b.OtherChargeVal3, " & _ 
                            "a.OtherChargeVal4=b.OtherChargeVal4, " & _ 
                            "a.OtherChargeVal5=b.OtherChargeVal5, " & _ 
                            "a.OtherChargeVal6=b.OtherChargeVal6, " 

                        end if

                    end if

                end if

                '//////////////////////////ACTUALIZA NOMBRES DE RUBROS Y OTROS
                QuerySelect = QuerySelect & _        
                "a.WeightsSymbol=b.WeightsSymbol, " & _ 
                "a.Commodities=b.Commodities, " & _ 
                "a.NatureQtyGoods=b.NatureQtyGoods, " & _ 
                "a.AdditionalChargeName1_Routing=b.AdditionalChargeName1_Routing, " & _ 
                "a.AdditionalChargeName2_Routing=b.AdditionalChargeName2_Routing, " & _ 
                "a.AdditionalChargeName3_Routing=b.AdditionalChargeName3_Routing, " & _ 
                "a.AdditionalChargeName4_Routing=b.AdditionalChargeName4_Routing, " & _ 
                "a.AdditionalChargeName5_Routing=b.AdditionalChargeName5_Routing, " & _ 
                "a.AdditionalChargeName6_Routing=b.AdditionalChargeName6_Routing, " & _ 
                "a.AdditionalChargeName7_Routing=b.AdditionalChargeName7_Routing, " & _ 
                "a.AdditionalChargeName8_Routing=b.AdditionalChargeName8_Routing, " & _ 
                "a.AdditionalChargeName9_Routing=b.AdditionalChargeName9_Routing, " & _ 
                "a.AdditionalChargeName10_Routing=b.AdditionalChargeName10_Routing, " & _ 
                "a.AdditionalChargeName11_Routing=b.AdditionalChargeName11_Routing, " & _ 
                "a.AdditionalChargeName12_Routing=b.AdditionalChargeName12_Routing, " & _ 
                "a.AdditionalChargeName13_Routing=b.AdditionalChargeName13_Routing, " & _ 
                "a.AdditionalChargeName14_Routing=b.AdditionalChargeName14_Routing, " & _ 
                "a.AdditionalChargeName15_Routing=b.AdditionalChargeName15_Routing, " & _                                
                "a.AdditionalChargeName1=b.AdditionalChargeName1, " & _ 
                "a.AdditionalChargeName2=b.AdditionalChargeName2, " & _ 
                "a.AdditionalChargeName3=b.AdditionalChargeName3, " & _ 
                "a.AdditionalChargeName4=b.AdditionalChargeName4, " & _ 
                "a.AdditionalChargeName5=b.AdditionalChargeName5, " & _ 
                "a.AdditionalChargeName6=b.AdditionalChargeName6, " & _ 
                "a.AdditionalChargeName7=b.AdditionalChargeName7, " & _ 
                "a.AdditionalChargeName8=b.AdditionalChargeName8, " & _ 
                "a.AdditionalChargeName9=b.AdditionalChargeName9, " & _ 
                "a.AdditionalChargeName10=b.AdditionalChargeName10, " & _ 
                "a.AdditionalChargeName11=b.AdditionalChargeName11, " & _ 
                "a.AdditionalChargeName12=b.AdditionalChargeName12, " & _ 
                "a.AdditionalChargeName13=b.AdditionalChargeName13, " & _ 
                "a.AdditionalChargeName14=b.AdditionalChargeName14, " & _ 
                "a.AdditionalChargeName15=b.AdditionalChargeName15, " 



                if AwbType = 2 then 'si es IMPORT agrega estos campos

                    QuerySelect = QuerySelect & _
                    "a.OtherChargeName1=b.OtherChargeName1, " & _ 
                    "a.OtherChargeName2=b.OtherChargeName2, " & _ 
                    "a.OtherChargeName3=b.OtherChargeName3, " & _ 
                    "a.OtherChargeName4=b.OtherChargeName4, " & _ 
                    "a.OtherChargeName5=b.OtherChargeName5, " & _ 
                    "a.OtherChargeName6=b.OtherChargeName6, " 

                end if

                if ClientCollectID_tmp = "0" or ClientCollectID_tmp = "2" or ClientCollectID_tmp = "4" then
                                  
                    if AwbType = 1 then 'si es export agrega estos campos

                        QuerySelect = QuerySelect & _
                        "a.CustomFee_Routing=b.CustomFee_Routing, " & _ 
                        "a.TerminalFee_Routing=b.TerminalFee_Routing, " & _ 
                        "a.TAX_Routing=b.TAX_Routing, " & _ 
                        "a.replica=b.replica, " 

                    end if
                                               
                    QuerySelect = QuerySelect & "a.AccountShipperNo=b.AccountShipperNo, " & _ 
                    "a.ShipperData=b.ShipperData, " & _ 
                    "a.AccountConsignerNo=b.AccountConsignerNo, " & _ 
                    "a.AgentData=b.AgentData, " & _ 
                    "a.AccountInformation=b.AccountInformation, " & _ 
                    "a.IATANo=b.IATANo, " & _ 
                    "a.AccountAgentNo=b.AccountAgentNo, " & _ 
                    "a.AirportDepID=b.AirportDepID, " & _ 
                    "a.RequestedRouting=b.RequestedRouting, " & _ 
                    "a.AirportToCode1=b.AirportToCode1, " & _ 
                    "a.CarrierID=b.CarrierID, " & _ 
                    "a.AirportToCode2=b.AirportToCode2, " & _ 
                    "a.AirportToCode3=b.AirportToCode3, " & _ 
                    "a.CarrierCode2=b.CarrierCode2, " & _ 
                    "a.CarrierCode3=b.CarrierCode3, " & _ 
                    "a.CurrencyID=b.CurrencyID, " & _ 
                    "a.ChargeType=b.ChargeType, " & _ 
                    "a.ValChargeType=b.ValChargeType, " & _ 
                    "a.OtherChargeType=b.OtherChargeType, " & _ 
                    "a.DeclaredValue=b.DeclaredValue, " & _ 
                    "a.AduanaValue=b.AduanaValue, " & _ 
                    "a.AirportDesID=b.AirportDesID, " & _ 
                    "a.FlightDate1=b.FlightDate1, " & _ 
                    "a.FlightDate2=b.FlightDate2, " & _ 
                    "a.SecuredValue=b.SecuredValue, " & _ 
                    "a.HandlingInformation=b.HandlingInformation, " & _ 
                    "a.Observations=b.Observations, " & _ 
                    "a.TotChargeWeightPrepaid=b.TotChargeWeightPrepaid, " & _ 
                    "a.TotChargeWeightCollect=b.TotChargeWeightCollect, " & _ 
                    "a.TotChargeValuePrepaid=b.TotChargeValuePrepaid, " & _ 
                    "a.TotChargeValueCollect=b.TotChargeValueCollect, " & _ 
                    "a.TotChargeTaxPrepaid=b.TotChargeTaxPrepaid, " & _ 
                    "a.TotChargeTaxCollect=b.TotChargeTaxCollect, " & _ 
                    "a.TotPrepaid=b.TotPrepaid, " & _ 
                    "a.TotCollect=b.TotCollect, " & _ 
                    "a.Invoice=b.Invoice, " & _ 
                    "a.ExportLic=b.ExportLic, " & _ 
                    "a.AgentContactSignature=b.AgentContactSignature, " & _ 
                    "a.CommoditiesTypes=b.CommoditiesTypes, " & _ 
                    "a.TotWeightChargeable=b.TotWeightChargeable, " & _ 
                    "a.Instructions=b.Instructions, " & _ 
                    "a.Agentsignature=b.Agentsignature, " & _ 
                    "a.Countries=b.Countries, " & _ 
                    "a.ReservationDate=b.ReservationDate, " & _ 
                    "a.DeliveryDate=b.DeliveryDate, " & _ 
                    "a.DepartureDate=b.DepartureDate, " & _ 
                    "a.Comment=b.Comment, " & _ 
                    "a.Comment2=b.Comment2, " & _ 
                    "a.ArrivalDate=b.ArrivalDate, " & _ 
                    "a.HDepartureDate=b.HDepartureDate, " & _ 
                    "a.Cont=b.Cont, " & _ 
                    "a.Destinity=b.Destinity, " & _ 
                    "a.TotalToPay=b.TotalToPay, " & _ 
                    "a.Concept=b.Concept, " & _ 
                    "a.FiscalFactory=b.FiscalFactory, " & _ 
                    "a.ArrivalAttn=b.ArrivalAttn, " & _ 
                    "a.ArrivalFlight=b.ArrivalFlight, " & _ 
                    "a.Comment3=b.Comment3, " & _ 
                    "a.DisplayNumber=b.DisplayNumber, " & _ 
                    "a.WType=b.WType, " & _ 
                    "a.ShipperID=b.ShipperID, " & _ 
                    "a.AgentID=b.AgentID, " & _ 
                    "a.SalespersonID=b.SalespersonID, " & _ 
                    "a.ShipperAddrID=b.ShipperAddrID, " & _ 
                    "a.ConsignerAddrID=b.ConsignerAddrID, " & _ 
                    "a.AgentAddrID=b.AgentAddrID, " & _ 
                    "a.Voyage=b.Voyage, " & _                
                        "a.TerminalFee=b.TerminalFee, " & _ 
                        "a.CustomFee=b.CustomFee, " & _ 
                        "a.FuelSurcharge=b.FuelSurcharge, " & _ 
                        "a.SecurityFee=b.SecurityFee, " & _ 
                        "a.PBA=b.PBA, " & _ 
                        "a.TAX=b.TAX, " & _ 
                        "a.PickUp=b.PickUp, " & _ 
                        "a.Intermodal=b.Intermodal, " & _ 
                        "a.SedFilingFee=b.SedFilingFee, " & _ 
                    "a.RoutingID=b.RoutingID, " & _ 
                    "a.Routing=b.Routing, " & _ 
                    "a.ManifestNumber=b.ManifestNumber, " & _ 
                    "a.CalcAdminFee=b.CalcAdminFee, " & _ 
                    "a.AWBDate=b.AWBDate, " & _ 
                    "a.CTX=b.CTX, " & _ 
                    "a.TCTX=b.TCTX, " & _ 
                    "a.TPTX=b.TPTX, " & _ 
                    "a.UserID=b.UserID, " & _ 
                    "a.Closed=b.Closed, " & _ 
                    "a.MAWBID=b.MAWBID, " & _ 
                    "a.rep_exp=b.rep_exp, " & _ 
                    "a.ConsignerColoader=b.ConsignerColoader, " & _ 
                    "a.ShipperColoader=b.ShipperColoader, " & _ 
                    "a.AgentNeutral=b.AgentNeutral, " & _ 
                    "a.ManifestNo=b.ManifestNo, " & _ 
                    "a.MonitorArrivalDate=b.MonitorArrivalDate, " & _ 
                    "a.ClientCollectID=b.ClientCollectID, " & _ 
                    "a.ClientsCollect=b.ClientsCollect, " & _ 
                    "a.id_coloader=b.id_coloader, " & _ 
                    "a.TotCarrierRate_Routing=b.TotCarrierRate_Routing, " & _ 
                    "a.FuelSurcharge_Routing=b.FuelSurcharge_Routing, " & _ 
                    "a.SecurityFee_Routing=b.SecurityFee_Routing, " & _ 
                    "a.PickUp_Routing=b.PickUp_Routing, " & _ 
                    "a.SedFilingFee_Routing=b.SedFilingFee_Routing, " & _ 
                    "a.Intermodal_Routing=b.Intermodal_Routing, " & _ 
                    "a.PBA_Routing=b.PBA_Routing, " & _ 
                    "a.AnotherChargesAgentPrepaid=b.AnotherChargesAgentPrepaid, " & _ 
                    "a.AnotherChargesAgentCollect=b.AnotherChargesAgentCollect, " & _ 
                    "a.AnotherChargesCarrierPrepaid=b.AnotherChargesCarrierPrepaid, " & _ 
                    "a.AnotherChargesCarrierCollect=b.AnotherChargesCarrierCollect, " & _ 
                    "a.id_cliente_order=b.id_cliente_order, " & _ 
                    "a.id_cliente_orderData=b.id_cliente_orderData, " & _ 
                    "a.file=' ', " & _ 
                    "a.FreezeCosts=b.FreezeCosts, " 

                end if


                QuerySelect = "UPDATE " & IIf(AwbType = 1,"Awb","Awbi") & " a " & _ 
                "INNER JOIN " & IIf(AwbType = 1,"Awb","Awbi") & " b on b.AwbID = " & ObjectID & " " & _ 
                "SET " & QuerySelect & " a.flg_master=b.flg_master, a.flg_totals=b.flg_totals " & _ 
                "WHERE a.AwbID IN (" & ObjectIDtmp & ")"                
                Conn.Execute(QuerySelect)

            end if

        end if 'if ObjectID <> ObjectIDtmp then 
    

        'response.write QuerySelect & "<br>"

    If Err.Number <> 0 Then

        response.write QuerySelect & "<br>"

        response.write "Problemas al replicar guia, por favor intente nuevamente.<br>"

        response.write "ReplicarHeaderRubros : " & Err.Number & " - " & Err.Description & "<br>"

    End If

End Sub




Function ReplicaRubros(Conn, c, ObjectID, AwbType, AwbID, esMaster, esHija, peso) 


    On Error Resume Next

        QuerySelect = ""
        ReplicaRubros = 0
        Dim Qry2

        'replica rubros de la master-hija a las hijas      En ChargeItems AwbType = 1 Export  ; AwbType = 2 Import          
        QuerySelect = "CREATE TEMPORARY TABLE temp_" & c & " ENGINE=MEMORY AS ( SELECT * FROM ChargeItems WHERE AWBID = " & ObjectID & " AND Expired = 0 AND DocTyp = " & IIf(AwbType = 1,0,1) & " AND TarifaTipo <> 'FLAT' );" 'FLAT NO SE REPLICA
        Qry2 = Qry2 & QuerySelect & "<br>"
        'response.write QuerySelect & "<br>"
        Conn.Execute(QuerySelect)
        ReplicaRubros = ReplicaRubros + 1

        'ObjectIDtmp pueden venir varios ids, se debera hacer un loop
        QuerySelect = "UPDATE temp_" & c & " SET AWBID = " & AwbID & ", " & IIF(esMaster = True and esHija = False,"Value = 0, TarifaPricing = '0', ","") & "CreatedDate = CURDATE(), CreatedTime = DATE_FORMAT(NOW(), '%H%i%S'), ChargeID = null" 

        if Cdbl(peso) > 0 then 'agrega el recalculo para el peso de las hijas incluso si este cambia
            QuerySelect = QuerySelect & ", Value = cast(TarifaPricing as decimal(11,2))*" & Cdbl(peso)
        end if
        
        QuerySelect = QuerySelect & ";"

        Qry2 = Qry2 & QuerySelect & "<br>"
        'response.write QuerySelect & "<br>"
        Conn.Execute(QuerySelect)
        ReplicaRubros = ReplicaRubros + 1

        QuerySelect = "UPDATE ChargeItems SET Expired = 1 WHERE AWBID = " & AwbID & " AND Expired = 0 AND DocTyp = " & IIf(AwbType = 1,0,1) & " AND ItemID IN (SELECT ItemID FROM temp_" & c & ") " & IIF(esMaster = True and esHija = False,"AND Value = 0","") & ";"
        Qry2 = Qry2 & QuerySelect & "<br>"
        'response.write QuerySelect & "<br>"
        Conn.Execute(QuerySelect)
        ReplicaRubros = ReplicaRubros + 1

        QuerySelect = "INSERT INTO ChargeItems SELECT * FROM temp_" & c & " WHERE ItemID NOT IN (SELECT ItemID FROM ChargeItems WHERE AWBID = " & AwbID & " AND Expired = 0 AND DocTyp = " & IIf(AwbType = 1,0,1) & ");"                 
        Qry2 = Qry2 & QuerySelect & "<br>"
        'response.write QuerySelect & "<br><br>"
        Conn.Execute(QuerySelect)
        ReplicaRubros = ReplicaRubros + 1

        'QuerySelect = "SELECT * FROM ChargeItems WHERE AWBID = " & ObjectIDtmp & " AND Expired = 0 AND DocTyp = " & IIf(AwbType = 1,0,1) & ";"
        'response.write QuerySelect & "<br>"
        
    If Err.Number <> 0 Then

        response.write Qry2 & "<br>"

        response.write "(d=" & d & ") Problemas al replicar rubros, por favor intente nuevamente.<br>"

        response.write "ReplicaRubros : " & Err.Number & " - " & Err.Description & "<br>"

    End If

End Function


Function ValidaResultadoTarifa(campo, TarifaPricing, peso)

    ValidaResultadoTarifa = ""

    On Error Resume Next
            
            if CDbl(TarifaPricing) > 0 then

                ValidaResultadoTarifa = "a." & campo & "=" & CDbl(TarifaPricing)*CDbl(peso) & ", "

            end if

    If Err.Number <> 0 Then

        response.write "<br>ValidaResultadoTarifa Error """ & campo & """ : " & Err.Number & " - " & Err.description & "<br>"  

    end if

End Function




Function RecalculoPorPeso(Conn, rs, ObjectIDtmp, peso, AWBType) 

    RecalculoPorPeso = ""

    On Error Resume Next

        RecalculoPorPeso = "SELECT ItemID, ItemName, AgentTyp, Pos, TarifaPricing, Value FROM ChargeItems WHERE AWBID IN (" & ObjectIDtmp & ") AND Expired = '0' AND DocTyp = '" & IIf(AWBType = 1,"0","1") & "' ORDER BY AgentTyp, Pos"
        'response.write RecalculoPorPeso & "<br>"
        Set rs = Conn.Execute(RecalculoPorPeso)
	    if Not rs.EOF then      

            RecalculoPorPeso = "a.ChargeableWeights=" & CDbl(peso) & ", a.TotWeightChargeable=" & CDbl(peso) & ", "
    
            Do While Not rs.EOF

                Select Case rs("ItemID")
                    Case 11 'Air Freight                                        
            
                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("Carriersubtot", rs("TarifaPricing"), peso) 

                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("TotCarrierRate", rs("TarifaPricing"), peso) 
                                        
                    Case 12 'Fuel Surcharge
                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("FuelSurcharge", rs("TarifaPricing"), peso) 
                                        
                    Case 13 'Security Charge
                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("SecurityFee", rs("TarifaPricing"), peso) 
                                        
                    Case 14 'Custom Fee
                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("CustomFee", rs("TarifaPricing"), peso) 
                                        
                    Case 15 'Terminal Fee
                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("TerminalFee", rs("TarifaPricing"), peso) 
                                        
                    Case 31 'Pick Up
                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("PickUp", rs("TarifaPricing"), peso) 
                                        
                    Case 38 'Sed (Sed Filling Fee)
                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("SedFilingFee", rs("TarifaPricing"), peso) 
                                        
                    Case 115 'Intermodal
                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("Intermodal", rs("TarifaPricing"), peso) 
                                        
                    Case 116 'PBA
                        RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("PBA", rs("TarifaPricing"), peso) 
                                                                              
                    Case Else

                        Select Case rs("AgentTyp")
                        Case "0"    'Carrier 

                            Select Case rs("Pos")
                            Case "1"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal3", rs("TarifaPricing"), peso) 
                            Case "2" 
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal4", rs("TarifaPricing"), peso) 
                            Case "3"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal5", rs("TarifaPricing"), peso) 
                            Case "4"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal8", rs("TarifaPricing"), peso) 
                            End Select

                        Case "1"    'Agente

                            Select Case rs("Pos")
                            Case "1"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal1", rs("TarifaPricing"), peso) 
                            Case "2" 
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal2", rs("TarifaPricing"), peso) 
                            Case "3"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal6", rs("TarifaPricing"), peso) 
                            Case "4"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal7", rs("TarifaPricing"), peso) 
                            Case "5"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal9", rs("TarifaPricing"), peso) 
                            Case "6"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal10", rs("TarifaPricing"), peso) 
                            Case "7"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal11", rs("TarifaPricing"), peso) 
                            Case "8"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal12", rs("TarifaPricing"), peso) 
                            Case "9"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal13", rs("TarifaPricing"), peso) 
                            Case "10"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal14", rs("TarifaPricing"), peso) 
                            Case "11"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("AdditionalChargeVal15", rs("TarifaPricing"), peso) 
                            End Select

                        Case "2"    'Otros

                            Select Case rs("Pos")
                            Case "1"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("OtherChargeVal1", rs("TarifaPricing"), peso) 
                            Case "2" 
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("OtherChargeVal2", rs("TarifaPricing"), peso) 
                            Case "3"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("OtherChargeVal3", rs("TarifaPricing"), peso) 
                            Case "4"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("OtherChargeVal4", rs("TarifaPricing"), peso) 
                            Case "5"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("OtherChargeVal5", rs("TarifaPricing"), peso) 
                            Case "6"
                                    RecalculoPorPeso = RecalculoPorPeso & ValidaResultadoTarifa("OtherChargeVal6", rs("TarifaPricing"), peso) 
                            End Select

                        End Select
                End Select

                rs.MoveNext
	        Loop

            CloseOBJ rs
	                    
	    end if

    If Err.Number <> 0 Then

        response.write "<br>RecalculoPorPeso Error : " & Err.Number & " - " & Err.description & "<br>"  

    end if

End Function


Function TarifarioPricing(AwbType, Countries, ObjectID, ServiceID, ItemID, No, ItemTarifa, ItemTarifaHidden, ItemMonto, peso) 

    Dim QuerySelect, rs, Conn, i, TipoGuia, Msg, Screen, Seguir, tip_TarifaRango, aiee_TipoAwb, TotWeight, tip_pais, tip_tipo, tip_tipo_tarifa, tip_movimiento, tip_salida, tip_destino, tip_carrier, tip_consignee, tip_shipper, tip_coloader, tip_moneda, tip_monto, tip_minimo, tip_flat, tip_rango, TipoCarga, Tarifa, aList0Values, aList1Values, CountList0Values, CountList1Values, aList2Values, CountList2Values, Country ', InternationalLocal, IntercompanyFilter, txtbusqueda
    
    TarifarioPricing = ""

    On Error Resume Next

        Tarifa = 0
        TotWeight = 0
        aiee_TipoAwb = ""
        Screen = ""

        Seguir = False

        '////////////////////////////////////////////////////////////   SI PAIS ESTA ENTRE LISTAS PRICING ///////////////////////////////////////////
        If InStr(1, Session("Pricing"), Countries) > 0 Then

            OpenConn Conn

            tip_tipo = ""
            tip_movimiento = "EXPORT" 'Iif(AwbType = "1", "EXPORT", "IMPORT")
            tip_salida = 0
            tip_destino = 0
            tip_carrier = 0
            tip_consignee = 0
            tip_shipper = 0
            tip_coloader = 0
            tip_TarifaRango = 0
            tip_moneda = ""      
            TipoGuia = ""
            Msg = ""
            tip_tipo_tarifa = ""
            
            '                       0           1               2           3           4           5                   6                                   7                       8                           9       10          11              12          13              14                  15                              17      18                      19                          20          21                                          22                                          23                      24                                              25
            QuerySelect = "SELECT AwbID, a.CreatedDate, a.CreatedTime, AwbNumber, HAwbNumber, a.Countries, COALESCE(ConsignerData,'') as ConsignerData, TotNoOfPieces, COALESCE(TotWeight,0)+0 AS TotWeight, Routing, a.CarrierID, AirportDepID, AirportDesID, SalespersonID, b.Name, COALESCE(aiee_TipoAwb,'') as aiee_TipoAwb, Routing, ShipperID, COALESCE(ConsignerID,0) as ConsignerID, AgentID, id_coloader, COALESCE(aiee_replica,'') as aiee_replica, COALESCE(aiee_master_hija,'') as aiee_master_hija, COALESCE(ChargeableWeights,0)+0 AS ChargeableWeights, aiee_TipoCarga FROM " & IIf(AwbType = "1", "Awb", "Awbi") & " a " & _ 
            "INNER JOIN Carriers b ON a.CarrierID = b.CarrierID " & _ 
            "LEFT JOIN Awb_IE_Expansion c ON c.aiee_AwbID_fk = a.AwbID AND aiee_ImpExp = " & AwbType & " " & _
            "WHERE AwbID = " & ObjectID 
            'Screen = Screen & QuerySelect & "<br>"
            Set rs = Conn.Execute(QuerySelect)
            If Not rs.EOF Then 
                
                'lo usaran los paises que no tiene configuracion nueva
                aiee_TipoAwb = rs("aiee_replica")   'Consolidado / Directo
    
                aiee_TipoAwb = rs("aiee_TipoAwb")   'Master-Hija / Hija-Directa / Master-Master-Hija 

                TipoCarga = rs("aiee_TipoCarga")

                if aiee_TipoAwb <> "" then

                    'if CDbl(rs("TotWeight")) >= CDbl(rs("ChargeableWeights")) then 2022-05-13 reunion se define que debe ser ChargeableWeights queda a discrecion del usuario
                    '    TotWeight = CDbl(rs("TotWeight"))
                    'else
                        TotWeight = CDbl(rs("ChargeableWeights"))
                    'end if

                    tip_salida = rs("AirportDepID") 
                    tip_destino = rs("AirportDesID")
                    tip_pais = rs("Countries")
                    tip_carrier = rs("CarrierID")
                    tip_consignee = rs("ConsignerID")
                    tip_shipper = rs("ShipperID")
                    tip_coloader = rs("id_coloader")

                    'Screen = Screen &  "(" & tip_carrier & ")(" & tip_consignee & ")(" & tip_coloader & ")<br>"

                    'CONSOLIDADO
                    if rs("AwbNumber") <> "" and rs("HAwbNumber") = "" then 'Consolidado y Master
                        tip_tipo = "COSTO"
                        TipoGuia = "MASTER "
                    end if

                    if rs("AwbNumber") <> rs("HAwbNumber") and rs("HAwbNumber") <> "" then 'Consolidado y Hija
                        tip_tipo = "VENTA"
                        TipoGuia = "HIJA " & No & " " 
                    end if

                    'DIRECTO
                    if rs("AwbNumber") = rs("HAwbNumber") and rs("HAwbNumber") <> "" then   'Directo
                        tip_tipo = "VENTA"
                        TipoGuia = "DIRECTO " 
                    end if
                
                    'Screen = Screen &  "(" & TipoGuia & ")(" & rs("aiee_master_hija") & ")<br>"

                end if

            end if 

            CloseOBJ rs

        CloseOBJ Conn

        if CDbl(peso) > 0 then
            TotWeight = peso
        end if



        Seguir = False
        'segun platica 2022-03-10 Operativo / IT
        'if aiee_TipoAwb = "Master-Master-Hija" and tip_tipo = "VENTA" and No = 1 then 'dividir las hijas, la 1. si busca tarifas, la 2. no requiere
        '    Seguir = True
        'else

        '    if aiee_TipoAwb = "Master-Master-Hija" and tip_tipo = "VENTA" and No = 2 then 'dividir las hijas, la 1. si busca tarifas, la 2. no requiere
        '        Seguir = False
        '    else
                if aiee_TipoAwb = "Master-Hija" or aiee_TipoAwb = "Hija-Directa" or aiee_TipoAwb = "Master-Master-Hija" then        
                    Seguir = True
                end if
        '    end if
        
        'end if

        else
            Screen = Screen &  Countries & " no se han creado listas en Pricing<br>"
        end if 'sigue si el pais esta configurado listas



        Msg = ""
        
        Dim c, max, rangos, data
        Dim tpl_transporte_fk
        Dim tpe_tipo_persona_fk
        Dim tpe_id_persona_fk
                                
        if tip_tipo = "VENTA" then

            tpe_tipo_persona_fk = 1 'cliente / shipper

            tpe_id_persona_fk = Iif(aiee_TipoAwb = "Master-Master-Hija", tip_shipper, tip_consignee)

        end if


        if tip_tipo = "COSTO" then

            tpe_tipo_persona_fk = 4 'carrier

            tpe_id_persona_fk = tip_carrier

        end if

        tpl_transporte_fk = 1 'AEREO


        tip_monto = 0
        tip_minimo = 0
        tip_flat = 0
        tip_rango = 0

        '////////////////////////////////TOTAL WEIGHT TEST/////////////////////////////////7777
        TotWeight = CDbl(TotWeight)
        'TotWeight = CDbl(31)

        'Screen = Screen &  "(" & aiee_TipoAwb & ")(" & TotWeight & ")(" & tip_movimiento & ")(" & tip_tipo & ")(" & tip_salida & ")(" & tip_destino & ")<br>"


        if Seguir = True then

            OpenConn3 Conn

	        if ItemID > 0 then
		        'Obteniendo tarifa PRICING

                    'tip_tipo = "COSTO" 


                    'QuerySelect = "SELECT " & _ 
                    '"CASE WHEN tpg_tipo = 'MIN' THEN " & _ 
                    '"	CASE WHEN " & TotWeight & " <= tpg_valor_dec THEN tpat_tarifa " & _ 
                    '"	WHEN " & TotWeight & " > tpg_valor_dec THEN tpat_tarifa " & _ 
                    '"	END  " & _ 
                    '"WHEN tpg_tipo = 'FLAT' THEN tpat_tarifa " & _ 
                    '"ELSE tpat_tarifa END as tpat_tarifa, " & _ 
                    '"tpg_valor_dec, tpl_moneda_fk, tpl_pk, tpg_tipo " & _
                    '"FROM ""ti_pricing_articulo"" a " & _
                    '"INNER JOIN ""ti_pricing_ruta"" b ON b.""tpr_pk"" = a.""tpa_tpr_fk"" AND b.""tpr_tpp_origen_fk"" = " & tip_salida & " AND b.""tpr_tpp_destino_fk"" = " & tip_destino & " " & _
                    '"INNER JOIN ""ti_pricing_list"" c ON c.""tpl_pk"" = b.""tpr_tpl_fk"" AND c.""tpl_pais_fk"" = '" & tip_pais & "' AND c.""tpl_tipo_carga"" = '" & TipoCarga & "' AND c.""tpl_tipo"" = '" & tip_tipo & "' AND c.""tpl_movimiento"" = '" & tip_movimiento & "' AND c.""tpl_transporte_fk"" = '" & tpl_transporte_fk & "' AND CURRENT_DATE >= c.""tpl_fecha_inicio"" AND CURRENT_DATE <= c.""tpl_fecha_vencimiento"" AND c.""tpl_tps_fk"" = '1' " & _
                    '"INNER JOIN ti_pricing_articulo_rango ON tpag_tpa_fk = a.tpa_pk AND tpag_tps_fk = 1 " & _
                    '"INNER JOIN ti_pricing_rango ON tpg_pk = tpag_tpg_fk AND tpg_tps_fk = 1 " & _
                    '"INNER JOIN ti_pricing_articulo_tarifa ON tpat_tps_fk = 1 AND tpag_tpg_fk = tpat_tpg_fk AND tpag_tpa_fk = tpat_tpa_fk " & _
                    '"INNER JOIN ""ti_pricing_entidad"" ON ""tpe_tpl_fk"" = c.""tpl_pk"" AND ""tpe_tps_fk"" = '1' AND ""tpe_tipo_persona_fk"" = '" & tpe_tipo_persona_fk & "' AND ""tpe_id_persona_fk"" = '" & tpe_id_persona_fk & "' " & _
                    '"WHERE  a.""tpa_servicio_fk"" = " & ServiceID & " AND a.""tpa_rubro_fk"" = " & ItemID & " AND " & _
                    '"CASE WHEN tpg_tipo = 'MIN' THEN " & _ 
                    '"	CASE WHEN " & TotWeight & " <= tpg_valor_dec THEN true  " & _ 
                    '"	WHEN " & TotWeight & " > tpg_valor_dec THEN true " & _ 
                    '"	END  " & _ 
                    '"WHEN tpg_tipo = 'FLAT' THEN true " & _ 
                    '"ELSE " & _ 
                    '"	" & TotWeight & " <= tpg_valor_dec " & _ 
                    '"END " & _ 
                    '"ORDER BY tpat_tarifa DESC  " & _
                    '"LIMIT 1"


                    '2022-03-28
QuerySelect = "select * from ( " & _  
"SELECT  " & _  
"CASE WHEN tpg_tipo = 'RANGO' THEN tpg_valor_dec ELSE -1 END as ""VAL_RANGO"", " & _  
"CASE WHEN tpg_tipo = 'FLAT' THEN tpg_valor_dec ELSE -1 END as ""VAL_FLAT"",  " & _  
"CASE WHEN tpg_tipo = 'MIN' THEN tpg_valor_dec ELSE -1 END as ""VAL_MIN"",  " & _  
"CASE WHEN tpg_tipo = 'RANGO' THEN tpat_tarifa ELSE -1 END as ""TAR_RANGO"", " & _  
"CASE WHEN tpg_tipo = 'FLAT' THEN tpat_tarifa ELSE -1 END as ""TAR_FLAT"",  " & _  
"CASE WHEN tpg_tipo = 'MIN' THEN tpat_tarifa ELSE -1 END as ""TAR_MIN"",  " & _  
" " & TotWeight & " * tpat_tarifa as ""MONTO"", tpl_moneda_fk, tpl_pk, tpg_tipo " & _  
"FROM ""ti_pricing_articulo"" a " & _
                    "INNER JOIN ""ti_pricing_ruta"" b ON b.""tpr_pk"" = a.""tpa_tpr_fk"" AND b.""tpr_tpp_origen_fk"" = " & tip_salida & " AND b.""tpr_tpp_destino_fk"" = " & tip_destino & " " & _
                    "INNER JOIN ""ti_pricing_list"" c ON c.""tpl_pk"" = b.""tpr_tpl_fk"" AND c.""tpl_pais_fk"" = '" & tip_pais & "' AND c.""tpl_tipo_carga"" = '" & TipoCarga & "' AND c.""tpl_tipo"" = '" & tip_tipo & "' AND c.""tpl_movimiento"" = '" & tip_movimiento & "' AND c.""tpl_transporte_fk"" = '" & tpl_transporte_fk & "' AND CURRENT_DATE >= c.""tpl_fecha_inicio"" AND CURRENT_DATE <= c.""tpl_fecha_vencimiento"" AND c.""tpl_tps_fk"" = '1' " & _
                    "INNER JOIN ti_pricing_articulo_rango ON tpag_tpa_fk = a.tpa_pk AND tpag_tps_fk = 1 " & _
                    "INNER JOIN ti_pricing_rango ON tpg_pk = tpag_tpg_fk AND tpg_tps_fk = 1 " & _
                    "INNER JOIN ti_pricing_articulo_tarifa ON tpat_tps_fk = 1 AND tpag_tpg_fk = tpat_tpg_fk AND tpag_tpa_fk = tpat_tpa_fk " & _
                    "INNER JOIN ""ti_pricing_entidad"" ON ""tpe_tpl_fk"" = c.""tpl_pk"" AND ""tpe_tps_fk"" = '1' AND ""tpe_tipo_persona_fk"" = '" & tpe_tipo_persona_fk & "' AND ""tpe_id_persona_fk"" = '" & tpe_id_persona_fk & "' " & _
                    "WHERE  a.""tpa_servicio_fk"" = " & ServiceID & " AND a.""tpa_rubro_fk"" = " & ItemID & " " & _
") x  ORDER BY ""VAL_RANGO"" ASC, ""VAL_FLAT"" ASC, ""VAL_MIN"" ASC"

                    'Screen = Screen &  "<b>Esta guia es " & TotWeight & " " & aiee_TipoAwb & " " & tip_pais & " " & tip_tipo & " </b><br>"

                    'Screen = Screen &  QuerySelect & "<br>"

                    if TotWeight > 0 then

                        Seguir = False

		                Set rs = Conn.Execute(QuerySelect)
                        If Not rs.EOF Then
			                
                            Do While Not rs.EOF

                                select case rs("tpg_tipo")
                                case "MIN"

                                    if rs("TAR_MIN") > 0 then
                                        tip_minimo = CDbl(rs("TAR_MIN")) & "|" & rs("tpl_moneda_fk") & "|" & rs("tpg_tipo")
                                    end if
                                
                                case "FLAT"

                                    if rs("TAR_FLAT") > 0 then 
                                        tip_flat = CDbl(rs("TAR_FLAT")) & "|" & rs("tpl_moneda_fk") & "|" & rs("tpg_tipo") 
                                    end if

                                case "RANGO"

                                    tip_rango = CDbl(rs("VAL_RANGO"))

                                    'Screen = Screen &  "1(TotWeight=" & TotWeight & ")(tip_rango=" & tip_rango & ")(tip_monto=" & tip_monto & ")<br>"  
                            
                                    if TotWeight <= tip_rango and tip_monto = "0" then

                                        tip_monto = CDbl(rs("MONTO")) & "|" & rs("tpl_moneda_fk") & "|" & rs("tpg_tipo") & " " & tip_rango & " " & rs("TAR_RANGO") & "*" & TotWeight

                                        'Screen = Screen &  "2(TotWeight=" & TotWeight & ")(tip_rango=" & tip_rango & ")(tip_monto=" & tip_monto & ")<br>"                                
                                    
                                        tip_TarifaRango = rs("TAR_RANGO")

                                    end if
                                    
                                end select
								                              
                                Seguir = True
								
                                rs.MoveNext
                            Loop	

                            Tarifa = 0

                            if Seguir = True then

                                'Screen = Screen &  "1(tip_monto=" & tip_monto & ")(tip_flat=" & tip_flat & ")(tip_minimo=" & tip_minimo & ")<br>"    

                                '/////////////////////RANGO///////////////////////////
                                Msg = Split(tip_monto,"|")
                                if ubound(Msg) > 0 then
                                    Tarifa = CDbl(Msg(0))
                                    tip_monto = CDbl(Msg(0))
                                    tip_moneda = Msg(1)
                                    tip_tipo_tarifa = Msg(2)
                                end if

                                '/////////////////////FLAT///////////////////////////
                                Msg = Split(tip_flat,"|")
                                if ubound(Msg) > 0 then
                                    tip_flat = CDbl(Msg(0))
                                    if tip_flat > 0 then                         
                                        Tarifa = tip_flat                            
                                        tip_moneda = Msg(1)
                                        tip_tipo_tarifa = Msg(2)
                                    end if
                                end if

                                '/////////////////////MINIMO///////////////////////////
                                Msg = Split(tip_minimo,"|")
                                if ubound(Msg) > 0 then
                                    tip_minimo = CDbl(Msg(0))
                                    if Tarifa < tip_minimo then                         
                                        Tarifa = tip_minimo                            
                                        tip_moneda = Msg(1)
                                        tip_tipo_tarifa = Msg(2)
                                    end if
                                end if


                                '/////////////////////CONFLICTO 1///////////////////////////
                                if (tip_monto > 0 or tip_minimo > 0) and tip_flat > 0 then 'conflicto tiene tarifa flat / rangos / minimo
                                    Screen = Screen &  "Hay conflicto en el tarifario Rangos / Minimo " & tip_minimo & " / Flat " & tip_flat & "<br>"
                                end if


                                '/////////////////////CONFLICTO 2 - NO HAY RANGO///////////////////////////
                                if tip_monto = 0 and tip_minimo = 0 and tip_flat = 0 then 
                                    max = 13 
                                    rangos = 1
                                end if

                                'Screen = Screen &  "2(tip_monto=" & tip_monto & ")(tip_flat=" & tip_flat & ")(tip_minimo=" & tip_minimo & ")<br>"    
                                    
                            end if
                                   
		                End If
		                CloseOBJ rs


                        Msg = "" 'tiene que inicializar para eliminar el array


                        if Seguir = False then

                        'este query funciona para ambos tipos COSTO VENTA
QuerySelect = "SELECT tpl_pk, tpr_pk, a.""tpa_servicio_fk"", a.""tpa_rubro_fk"", b.""tpr_tpp_origen_fk"", b.""tpr_tpp_destino_fk"", " & _

" c.""tpl_pais_fk"", c.""tpl_tipo_carga"", c.""tpl_tipo"", c.""tpl_movimiento"", c.""tpl_transporte_fk"", " & _

" CASE WHEN CURRENT_DATE >= c.""tpl_fecha_inicio"" AND CURRENT_DATE <= c.""tpl_fecha_vencimiento"" THEN 1 ELSE 0 END as fechas, " & _

" COALESCE(""tpe_tipo_persona_fk"",0) as tpe_tipo_persona_fk, COALESCE(""tpe_id_persona_fk"",0) as tpe_id_persona_fk " & _

" , (SELECT count(*) FROM ""ti_pricing_articulo_rango"" d WHERE d.""tpag_tpa_fk"" = a.""tpa_pk"" AND d.""tpag_tps_fk"" = 1) as rangos " & _ 

"FROM ""ti_pricing_articulo"" a " & _
"LEFT JOIN ""ti_pricing_ruta"" b            ON b.""tpr_pk"" = a.""tpa_tpr_fk"" AND b.""tpr_tps_fk"" = 1 " & _ 
"LEFT JOIN ""ti_pricing_list"" c            ON c.""tpl_pk"" = b.""tpr_tpl_fk"" AND c.""tpl_tps_fk"" = 1 " & _
"LEFT JOIN ""ti_pricing_entidad"" e         ON e.""tpe_tpl_fk"" = c.""tpl_pk"" AND e.""tpe_tps_fk"" = 1 " & _
"ORDER BY tpl_pk, tpr_pk, a.""tpa_servicio_fk"", a.""tpa_rubro_fk"" "

                            'Screen = Screen &  QuerySelect & "<br>"
		                    Set rs = Conn.Execute(QuerySelect)
                            If Not rs.EOF Then

                                Dim rstDuplicate 

                                Set rstDuplicate = rs.Clone 

                                max = 0                           

	                            Do While Not rs.EOF

                                    c = 0

                                    if rs("tpa_rubro_fk") = ItemID or rs("tpa_servicio_fk") = ServiceID or rs("tpe_tipo_persona_fk") = tpe_tipo_persona_fk or rs("tpe_id_persona_fk") = tpe_id_persona_fk or rs("fechas") = 1 or rs("tpl_transporte_fk") = tpl_transporte_fk or rs("tpl_movimiento") = tip_movimiento or rs("tpl_tipo") = tip_tipo or rs("tpl_tipo_carga") = TipoCarga or rs("tpl_pais_fk") = tip_pais or rs("tpr_tpp_destino_fk") = tip_destino or rs("tpr_tpp_origen_fk") = tip_salida or CheckNum(rs("rangos")) > 0 then


                                        if rs("tpa_rubro_fk") = ItemID then
                                            c = c + 1
                                        end if


                                        if rs("tpa_servicio_fk") = ServiceID then
                                            c = c + 1
                                        end if


                                            if rs("tpe_tipo_persona_fk") = tpe_tipo_persona_fk then 'cliente / shipper / carrier
                                                c = c + 1
                                            end if


                                            if rs("tpe_id_persona_fk") = tpe_id_persona_fk then
                                                c = c + 1
                                            end if


                                        if rs("fechas") = 1 then
                                            c = c + 1
                                        end if

                                        if rs("tpl_transporte_fk") = 1 then
                                            c = c + 1
                                        end if

                                        if rs("tpl_movimiento") = tip_movimiento then 
                                            c = c + 1
                                        end if


                                        if rs("tpl_tipo") = tip_tipo then
                                            c = c + 1
                                        end if


                                        if rs("tpl_tipo_carga") = TipoCarga then
                                            c = c + 1
                                        end if


                                        if rs("tpl_pais_fk") = tip_pais then
                                            c = c + 1
                                        end if

                                        if rs("tpr_tpp_destino_fk") = tip_destino then
                                            c = c + 1
                                        end if

                                        if rs("tpr_tpp_origen_fk") = tip_salida then
                                            c = c + 1
                                        end if


                                        if CheckNum(rs("rangos")) > 0 then
                                            c = c + 1
                                        end if

                                    end if

                                    if c > max then
                                        max = c
                                        rangos = CheckNum(rs("rangos")) 
                                        data = rs("tpl_pk") & "-" & rs("tpr_pk") & "-" & rs("tpa_servicio_fk") & "-" & rs("tpa_rubro_fk") & "-" & rs("tpe_id_persona_fk")  
                                    end if

                                    rs.MoveNext
	                            Loop
                                CloseOBJ rs

                                c = 0

                                Msg = ""

                                'Screen = Screen &  "Max " & max & " Rangos " & rangos & " " & data & "<br>"


                                if max > 0 then


                                'Set rs = Conn.Execute(QuerySelect)

                                Do While Not rstDuplicate.EOF

                                    if data = rstDuplicate("tpl_pk") & "-" & rstDuplicate("tpr_pk") & "-" & rstDuplicate("tpa_servicio_fk") & "-" & rstDuplicate("tpa_rubro_fk") & "-" & rstDuplicate("tpe_id_persona_fk") then 

                                        c = 0

                                        Msg = ""

                                        if rstDuplicate("tpa_rubro_fk") = ItemID then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Rubro " & ItemID & "<br>"
                                        end if


                                        if rstDuplicate("tpa_servicio_fk") = ServiceID then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Servicio " & ServiceID & "<br>"
                                        end if



                                            if rstDuplicate("tpe_tipo_persona_fk") = tpe_tipo_persona_fk and rstDuplicate("tpe_id_persona_fk") = tpe_id_persona_fk then

                                            else 

		                                        if tip_tipo = "VENTA" then
			                                        Msg = Msg & "Tipo Entidad " & Iif(aiee_TipoAwb = "Master-Master-Hija", "Shipper", "Cliente") & "<br>"
		                                        end if

		                                        if tip_tipo = "COSTO" then
			                                        Msg = Msg & "Tipo Entidad " & "Aerolinea" & "<br>"
		                                        end if

                                            end if


                                            if rstDuplicate("tpe_tipo_persona_fk") = tpe_tipo_persona_fk then 'cliente / shipper / carrier
                                                c = c + 1
                                            end if


                                            if rstDuplicate("tpe_id_persona_fk") = tpe_id_persona_fk then
                                                c = c + 1
			                                else
				                                Msg = Msg & "Entidad ID " & tpe_id_persona_fk & "<br>"
			                                end if


                                        

                                        

                                        if rstDuplicate("fechas") = 1 then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Fecha Vencida "
                                        end if

                                        if rstDuplicate("tpl_transporte_fk") = 1 then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Transporte " & "Aereo" & "<br>"
                                        end if

                                        if rstDuplicate("tpl_movimiento") = tip_movimiento then 
                                            c = c + 1
                                        else
                                            Msg = Msg & "Movimiento " & tip_movimiento & "<br>"
                                        end if


                                        if rstDuplicate("tpl_tipo") = tip_tipo then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Tipo Lista " & tip_tipo & "<br>"
                                        end if


                                        if rstDuplicate("tpl_tipo_carga") = TipoCarga then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Tipo Carga " & TipoCarga & "<br>"
                                        end if


                                        if rstDuplicate("tpl_pais_fk") = tip_pais then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Pais " & tip_pais & "<br>"
                                        end if

                                        if rstDuplicate("tpr_tpp_destino_fk") = tip_destino then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Destino " & tip_destino & "<br>"
                                        end if

                                        if rstDuplicate("tpr_tpp_origen_fk") = tip_salida then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Salida " & tip_salida & "<br>"
                                        end if


                                        if CheckNum(rstDuplicate("rangos")) > 0 then
                                            c = c + 1
                                        else
                                            Msg = Msg & "Rangos " & CheckNum(rstDuplicate("rangos")) & "<br>"
                                        end if

                                        'Screen = Screen &  "Max " & c & " Rangos " & CheckNum(rstDuplicate("rangos")) & " " & rstDuplicate("tpl_pk") & "-" & rstDuplicate("tpr_pk") & "-" & rstDuplicate("tpa_servicio_fk") & "-" & rstDuplicate("tpa_rubro_fk") & "<BR>"

                                    end if

                                    rstDuplicate.MoveNext
	                            Loop
                                CloseOBJ rstDuplicate

                                end if
          
                                Seguir = True
                 
		                    End If
		                    

                            'Screen = Screen &  "Tarifa (PRICING) Sin Valor<br>"
                            'TotWeight = 0
                        end if

                    else
                        Screen = Screen &  "Para capturar la Tarifa (PRICING), debe almacenar el peso de la carga<br>"
                    end if
                    

	        end if 'hay servicio


            Screen = Screen & " Mode: " & aiee_TipoAwb & " : " & Iif(AwbType = "1", "EXPORT", "IMPORT") & " : " & TipoGuia & " : " & tip_tipo & "<br>" '&  max & " " & rangos
            
            if Msg <> "" then
                Screen = Screen &  "Datos que no coinciden en busqueda de Tarifarios: <br>" & Msg  & "<br>"           
            end if

            if max = 13 and rangos > 0 then 'todos los datos coinciden, seguro no hay tarifas
                Screen = Screen &  ServiceID & " " & ItemID & " Sin rango disponible para ChargeableWeights : " & TotWeight             
                'TotWeight = 0 esto bloquea la seleccion del rubro
            end if

            CloseOBJ Conn 'cierra ti_pricing

        end if 'sigue

    If Err.Number <> 0 Then
        Screen = Screen &  "<br>TarifarioPricing Error : " & Err.Number & " - " & Err.description & "<br>"  
    end if

                        '       0                           1                       2                       3                   4                               5                           6                       7                   8                   9               10  
    TarifarioPricing = CStr(aiee_TipoAwb) & "|" & CStr(TotWeight) & "|" &  CStr(TipoCarga) & "|" & CStr(Tarifa) & "|" & CStr(tip_tipo_tarifa) & "|" & CStr(tip_TarifaRango) & "|" & CStr(tip_moneda) & "|" & ItemTarifa & "|" & ItemTarifaHidden & "|" & ItemMonto & "|" & Screen

End Function


Function IFNULL(dato)        
    IFNULL = ""
    if not IsNull(dato) then
	    IFNULL = dato
	end if
End Function




Function WsAirportDisplay(iConn, iRs, movimiento, tipo, id, Country2, Routing2, Transportista2, ReqAirportDepID2)

    Dim Result(3)

    Result(0) = ""
    Result(1) = "*"
    Result(2) = ""

    On Error Resume Next

        if InStr(1, Session("Pricing"), "'" & Country2 & "'") > 0 then

            OpenConn3 iConn

            'if movimiento = "EXPORT" and tipo = "SALIDA" then
            '    o = "EXPORT"
            'end if

            'if movimiento = "EXPORT" and tipo = "DESTINO" then
            '    'o = "IMPORT"
            '    o = "EXPORT"
            'end if

            'if movimiento = "IMPORT" and tipo = "SALIDA" then
            '    o = "IMPORT"
            'end if

            'if movimiento = "IMPORT" and tipo = "DESTINO" then
            '    'o = "EXPORT"
            '    o = "IMPORT"
            'end if

            QuerySelect = "SELECT DISTINCT tpp_pk as id, UPPER(tpp_codigo) as cod, UPPER(tpp_nombre) as nom " 
            
            'QuerySelect = QuerySelect & " -- tpl_pk, tpl_pais_fk, tpl_tipo, tpl_regional, tpl_transporte_fk, tpl_moneda_fk, tpl_movimiento, tpl_fecha_inicio, tpl_fecha_vencimiento, tpl_tps_fk " 
            
            'QuerySelect = QuerySelect & " -- , tpr_pk, tpr_tpl_fk, tpr_tpp_origen_fk, tpr_tpp_destino_fk " 

            QuerySelect = QuerySelect & " FROM ti_pricing_list " 

            QuerySelect = QuerySelect & " INNER JOIN ti_pricing_ruta ON tpr_tps_fk = '1' AND tpr_tpl_fk = tpl_pk " 

            QuerySelect = QuerySelect & " INNER JOIN ti_pricing_puerto ON tpp_pk IN (tpr_tpp_origen_fk,tpr_tpp_destino_fk) AND tpp_transporte_fk = '1' AND tpp_tps_fk = '1' " 

            QuerySelect = QuerySelect & " AND tpp_pais_iso_fk = CASE WHEN '" & tipo & "' = 'SALIDA' THEN SUBSTR(tpl_pais_fk,1,2) ELSE tpp_pais_iso_fk END " 

            QuerySelect = QuerySelect & " AND tpp_pais_iso_fk <> CASE WHEN '" & tipo & "' = 'DESTINO' THEN SUBSTR(tpl_pais_fk,1,2) ELSE '' END " 

            QuerySelect = QuerySelect & " WHERE tpl_pais_fk = '" & Country2 & "' AND tpl_transporte_fk = '1' AND tpl_movimiento = '" & movimiento & "' AND tpl_tps_fk = '1' AND tpl_tipo = 'VENTA' "  

        else

            OpenConn iConn
			
			QuerySelect = "SELECT DISTINCT b.AirportID as id, UPPER(b.AirportCode) as cod, UPPER(b.Name) as nom FROM Airports b "

            if tipo = "SALIDA" then               
                QuerySelect = QuerySelect & ", CarrierDepartures a WHERE a.AirportID = b.AirportID AND b.Expired=0 AND a.CarrierID = " & Transportista2 & " order by b.Name"
            else
                QuerySelect =  QuerySelect & ", CarrierRates a WHERE b.Expired=0 AND b.AirportID = a.AirportDesID AND a.CarrierID = " & Transportista2 & " and a.AirportDepID = " & ReqAirportDepID2
                QuerySelect = QuerySelect & " and b.AirportID <> " & ReqAirportDepID2 & " order by b.Name"
            end if

			
            'if movimiento = "EXPORT" and tipo = "SALIDA" then
            '    if Routing2 = "NINGUNO" then 
            '        QuerySelect = QuerySelect & ", CarrierDepartures a WHERE a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID = " & Transportista2 & " order by b.Name"
            '    else
            '        QuerySelect = QuerySelect & ", CarrierDepartures a WHERE a.AirportID = b.AirportID and b.Expired=0 order by b.Name"
            '    end if
            'end if

            'if movimiento = "EXPORT" and tipo = "DESTINO" then
            '    QuerySelect =  QuerySelect & ", CarrierRates a where b.Expired=0 and b.AirportID = a.AirportDesID and a.CarrierID = " & Transportista2 & " and a.AirportDepID = " & ReqAirportDepID2 
            '    if Routing2 = "NINGUNO" then                 
            '        QuerySelect = QuerySelect & " and a.AirportID <> " & ReqAirportDepID2 & " order by b.Name"
            '    end if
            'end if

            'if movimiento = "IMPORT" and tipo = "SALIDA" then
            '    QuerySelect = QuerySelect & ", CarrierDepartures a WHERE a.AirportID = b.AirportID and b.Expired=0 and a.CarrierID = " & Transportista2 & " order by b.Name"
            'end if

            'if movimiento = "IMPORT" and tipo = "DESTINO" then
            '    QuerySelect = QuerySelect & ", CarrierRates a where b.Expired=0 and b.AirportID = a.AirportDesID and a.CarrierID = " & Transportista2 & " and a.AirportDepID = " & ReqAirportDepID2
            '    if Routing2 = "NINGUNO" then                 
            '        QuerySelect = QuerySelect & " and b.AirportID <> " & ReqAirportDepID2 & " order by b.Name"
            '    end if
            'end if

        end if

        'response.write "" & QuerySelect & "<br>"
     
        Set iRs = iConn.Execute(QuerySelect)														
		If Not iRs.EOF Then

            Result(0) = Result(0) & "<option value=-1>Seleccionar</option>"
		    
            do while Not iRs.EOF

				Result(0) = Result(0) & "<option value=" & iRs("id") & ">" & iRs("nom") & " - " & iRs("cod") & " - " & iRs("id") & "</option>"

                if CheckNum(iRs("id")) = CheckNum(id) then
                    Result(1) = iRs("cod")
                    Result(2) = iRs("nom") & " - " & iRs("cod") & " - " & iRs("id")
                end if
				iRs.MoveNext
			loop

		End If

		CloseOBJ iRs      

        'CloseOBJ iConn

        'response.write "(" & Result(0) & ")<br>"
        'response.write "(" & Result(1) & ")<br>"
        'response.write "(" & Result(2) & ")<br>"

    If Err.Number <> 0 Then

        response.write "<br>WsAirportDisplay Error : " & Err.Number & " - " & Err.description & "<br>"  
        
    end if
    
    WsAirportDisplay = Result

End Function




Function ValidaHomologacion(module, esquema, cat, ids)        

    Dim iconn, iRs, i, Result, iQuery

    iQuery = "SELECT eh_codigo, eh_erp_categoria, eh_erp_codigo, eh_erp_descripcion " & _
    "FROM exactus_homologaciones " & _ 
    "WHERE eh_codigo IN (" & ids & ") AND eh_erp_esquema = '" & esquema & "' AND eh_estado = '1' AND eh_categoria = '" & cat & "' AND eh_modulo = '" & module &  "' " & _ 
    "LIMIT 50"
    'Response.Write iQuery & "<br>"
    On Error Resume Next

        OpenConn2 iConn
        Set iRs = iConn.Execute(iQuery)														
	    If Not iRs.EOF Then

            'i = iRs.RecordCount + 1
            'ReDim Result(i)
            'i = 0

    	    Result = iRs.GetRows

            'do while Not iRs.EOF
			'    Result(i) = iRs("eh_codigo") & "|" & iRs("eh_erp_categoria") & " - " & iRs("eh_erp_codigo") & " - " & iRs("eh_erp_descripcion")
            '    Response.Write Result(i) & "<br>"
			'    iRs.MoveNext
            '    i = i + 1
		    'loop

        end if
	    CloseOBJs iRs, iConn    
    
    If Err.Number <> 0 Then
        response.write "<br>ValidaHomologacion Error : " & Err.Number & " - " & Err.description & "<br>"  
    end if

    ValidaHomologacion = Result

End Function






</script>
