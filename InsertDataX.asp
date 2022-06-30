<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="Utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim JavaMsg, SQLQuery, DisplayPost, CountTableValues, GroupID, SQLQuery2, Routing, RoutingID, Currencies, Email
Dim Conn, ConnMaster, rs, Action, rsFilter, ObjectID, CanDisplayInfo, aTableValues, ShipperAddrID, ConsignerAddrID, AgentAddrID
Dim TableName, ObjectName, QuerySelect, CreatedDate, CreatedTime, i, j, k, l, RangeID, ReservationDate, DeliveryDate, DepartureDate
Dim CarrierID, AirportID, TerminalFeeCS, TerminalFeePD, CustomFee, FuelSurcharge, SecurityFee, Comment, Comment2
Dim CountList1Values, CountList2Values, CountList3Values, CountList4Values, CountList5Values, CountList6Values, CountList7Values, CountList8Values, CountList9Values
Dim aList1Values, aList2Values, aList3Values, aList4Values, aList5Values, aList6Values, aList7Values, aList8Values, aList9Values
Dim Val, NameES, NameEN, TypeVal, CommodityCode, CurrencyCode, Xchange, IATANo, DefaultVal, PhoneID, AttnID, SalespersonID
Dim Address, AddressID, Address2, Phone1, Phone2, AccountNo, Attn, Expired, AirportCode, Name, BillName, CarrierCode
Dim OtherChargesPrintType, AWBNumber, HAWBNumber, AccountShipperNo, ShipperData, AccountConsignerNo, Consolidating
Dim ConsignerData, AgentData, AccountInformation, AccountAgentNo, AirportDepID, RequestedRouting, CalcAdminFee
Dim AirportToCode1, AirportToCode2, AirportToCode3, CarrierCode2, CarrierCode3, CurrencyID, ShipperID, ConsignerID, AgentID
Dim ChargeType, ValChargeType, OtherChargeType, DeclaredValue, AduanaValue, AirportDesID, FlightDate1, WhereSQL
Dim FlightDate2, SecuredValue, HandlingInformation, Observations, NoOfPieces, Weights, WeightsSymbol, Instructions
Dim Commodities, ChargeableWeights, CarrierRates, CarrierSubTot, NatureQtyGoods, TotNoOfPieces, TotWeight
Dim TotCarrierRate, TotChargeWeightPrepaid, TotChargeWeightCollect, TotChargeValuePrepaid, TotChargeValueCollect
Dim TotChargeTaxPrepaid, TotChargeTaxCollect, AnotherChargesAgentPrepaid, AnotherChargesAgentCollect, UserCreate, UserModify
Dim AnotherChargesCarrierPrepaid, AnotherChargesCarrierCollect, TotPrepaid, TotCollect, TerminalFee, ManifestNumber
Dim PBA, Tax, TaxRate, AdditionalChargeName1, AdditionalChargeVal1, AdditionalChargeName2, CarrierName, ComisionRate
Dim AdditionalChargeVal2, Invoice, ExportLic, AgentContactSignature, AWBDate, AgentSignature, CommoditiesTypes, TotWeightChargeable
Dim AdditionalChargeName3, AdditionalChargeVal3, AdditionalChargeName4, AdditionalChargeVal4, Countries, Region
Dim AdditionalChargeName5, AdditionalChargeVal5, AdditionalChargeName6, AdditionalChargeVal6, AwbType, Symbol, Closed
Dim ArrivalDate, HDepartureDate, Cont, Destinity, TotalToPay, Concept, FiscalFactory, ArrivalAttn, ArrivalFlight, Comment3
Dim DisplayNumber, AdditionalChargeName7, AdditionalChargeVal7, AdditionalChargeName8, AdditionalChargeVal8, WType
Dim OtherChargeName1, OtherChargeName2, OtherChargeName3, OtherChargeName4, OtherChargeName5, OtherChargeName6
Dim OtherChargeVal1, OtherChargeVal2, OtherChargeVal3, OtherChargeVal4, OtherChargeVal5, OtherChargeVal6, BusinessGID
Dim AdditionalChargeName9, AdditionalChargeVal9, AdditionalChargeName10, AdditionalChargeVal10, HTMLCode, SearchOption, Estate, AWBID
Dim AdditionalChargeName11, AdditionalChargeVal11, AdditionalChargeName12, AdditionalChargeVal12, isConsigneer, isShipper, CTX, TCTX, TPTX
Dim AdditionalChargeName13, AdditionalChargeVal13, AdditionalChargeName14, AdditionalChargeVal14, AdditionalChargeName15, AdditionalChargeVal15
Dim Arancel_GT, Arancel_SV, Arancel_HN, Arancel_NI, Arancel_CR, Arancel_PA, Arancel_BZ, Voyage, PickUp, Intermodal, SedFilingFee, CreatedIn
Dim ConsignerColoader, ShipperColoader, AgentNeutral, BAWResult
Dim ClientCollectID, ClientsCollect, ItemCurrs, ItemIDs, ItemVals, ItemLocs, ItemNames, ItemNames_Routing, ItemPrePaid, ItemOVals, ItemPPCCs, ItemServIDs, ItemServNames
Dim ItemInvoices, ItemCalcInBls, ItemIntercompanyIDs, CantItems
Dim rst, ResValidaGuia, guia99, MAWBID, id_coloader, ColoaderData, AwbTable, Seguro, ClientCollectID_tmp

Dim TotCarrierRate_Routing
Dim FuelSurcharge_Routing
Dim SecurityFee_Routing
Dim CustomFee_Routing
Dim TerminalFee_Routing
Dim PickUp_Routing
Dim SedFilingFee_Routing
Dim Intermodal_Routing
Dim PBA_Routing
Dim TAX_Routing

Dim AdditionalChargeName1_Routing
Dim AdditionalChargeName2_Routing
Dim AdditionalChargeName3_Routing
Dim AdditionalChargeName4_Routing
Dim AdditionalChargeName5_Routing
Dim AdditionalChargeName6_Routing
Dim AdditionalChargeName7_Routing
Dim AdditionalChargeName8_Routing
Dim AdditionalChargeName9_Routing
Dim AdditionalChargeName10_Routing
Dim AdditionalChargeName11_Routing
Dim AdditionalChargeName12_Routing
Dim AdditionalChargeName13_Routing
Dim AdditionalChargeName14_Routing
Dim AdditionalChargeName15_Routing

Dim OtherChargeName1_Routing
Dim OtherChargeName2_Routing
Dim OtherChargeName3_Routing
Dim OtherChargeName4_Routing
Dim OtherChargeName5_Routing
Dim OtherChargeName6_Routing  

Dim id_cliente_order, id_cliente_orderData, ReplicaAwbID

'todos estos campos nuevos en db para bloqueo de rubros 2016-03-31
'ReplicaAwbID este campo se agrego 2017-07-07 para validar cuando sea consolidado cambiar aerolinea





	GroupID = CheckNum(Request("GID")) 
	'Revisando que el Grupo sea: 
	
	ObjectID = CheckNum(Request("OID"))
	SQLQuery = ""
	CountTableValues = -1
	Action = CheckNum(Request("Action"))
	CreatedDate = CheckTxt(Request("CD"))
	CreatedTime = CheckNum(Request("CT"))
	AwbType = CheckNum(Request("AT"))
	SearchOption = CheckNum(Request("SO"))
	AddressID = 0
    
    if AwbType = 1 then
		AwbTable = "Awb"
	else
		AwbTable = "Awbi"
	end if

	FormatTime CreatedDate, CreatedTime
	
	If GroupID >= 1 And GroupID <= 22 Then
			GetTableData GroupID, TableName, ObjectName, QuerySelect, AwbType
			'Preparando el Query de Seleccion
            JavaMsg = ""			

			Select Case GroupID
			Case 7, 8, 10, 11
				openConn2 Conn 'Abriendo la conexion a BBDD Master
			Case Else
				WhereSQL = "CreatedDate='" & CreatedDate & "' and CreatedTime="
				openConn Conn 'Abriendo la conexion a BBDD
			End Select
			


			If Action >= 1 And Action <= 3 Then
                 'obteniendo los parametros para hacer las operaciones de Insert, Update o Delete
                 'Creando los filtros para cada opcion de almacenamiento
				Select Case GroupID
				'Case 1 'Awb                                 
				'	AWBNumber = request.Form("AWBNumber")
				'	HAWBNumber = request.Form("HAWBNumber")
				'	RsFilter = " AWBNumber='" & AWBNumber & "' and HAWBNumber='" & HAWBNumber & "'"

				Case 2 'Carriers
					Name = request.Form("Name")
					Countries = request.Form("Countries")
					RsFilter = " Name='" & Name & "' and Countries='" & Countries & "'"
				Case 7, 8, 10 'Carriers, Shippers, Consigners, Agents
					'Name = request.Form("Name")
					'Countries = request.Form("Countries")
					'RsFilter = " Name='" & Name & "' and Countries='" & Countries & "'"
					Name = request.Form("Name")
					'RsFilter = " nombre_cliente='" & Name & "'"
					RsFilter = "select * from " & TableName & " where " & ObjectName & "=" & ObjectID '"id_cliente=" & ObjectID
					WhereSQL = "fecha_creacion='" & CreatedDate & "' and hora_creacion="					
				Case 3 'Transportistas-Salida
				  	AirportID = CheckNum(request.Form("AirportID"))
					CarrierID = CheckNum(request.Form("CarrierID"))
				  	RsFilter = " AirportID=" & AirportID & " and CarrierID=" & CarrierID
				Case 5 'Transportistas-Rango
				  	RangeID = CheckNum(request.Form("RangeID"))
					CarrierID = CheckNum(request.Form("CarrierID"))
				  	RsFilter = " RangeID=" & RangeID & " and CarrierID=" & CarrierID
				Case 9 'Aeropuertos
					AirportCode = request.Form("AirportCode")
					RsFilter = " AirportCode='" & AirportCode & "'"
				Case 11 'Commodities
					'CommodityCode = request.Form("CommodityCode")
					'RsFilter = " CommodityCode='" & CommodityCode & "'"
					RsFilter = "select * from " & TableName & " where " & ObjectName & "=" & ObjectID
				Case 12 'Monedas
					CurrencyCode = request.Form("CurrencyCode")
					Countries = request.Form("Countries")
					RsFilter = " CurrencyCode='" & CurrencyCode & "' and Countries='" & Countries & "'"
				Case 13 'Rangos
					Val = request.Form("Val")
					RsFilter = " Val=" & Val
				Case 14 'Taxes
					Countries = request.Form("Countries")
					RsFilter = " Countries='" & Countries & "'"
				Case 21
					RsFilter = ObjectName & "=" & ObjectID                                         
                    guia99 = UCase(Request.Form("guia0"))                                            
				Case 22
					RsFilter = ObjectName & "=" & ObjectID                                                             
				Case Else
					RsFilter = ObjectName & "=" & ObjectID & " and CreatedTime=" & CreatedTime
				End Select

                ResValidaGuia = false
                
				Function ValidaAWBNumber(AwbType)
                    
                    if AwbType = 1 and Request.Form("Countries") = "GT" then 'export

                        AWBNumber = Request.Form("AWBNumber")
                        HAWBNumber = Request.Form("HAWBNumber")
                        
                        ValidaAWBNumber = true

                        if Request.Form("CarrierID") <> Request.Form("CarrierIDAnt") and Request.Form("Countries") = "GT" then
                            
                            if Request.Form("TipoMaster") = "Nuevo" then						                
                                AWBNumber = NextAWBNumber(Conn, AwbType, Request.Form("CarrierID"), Request.Form("TipoMaster"))
                            end if

                            if AWBNumber = "" then                                    
                                JavaMsg = "No encontro guia AWBNumber" 
                                ValidaAWBNumber = false
                            else
                                if Action = 1 then	                                                                     
                                    HAWBNumber = NextHAWBNumber(HAWBNumber, Conn, AwbType, Request.Form("Countries"), Request.Form("TipoMaster"), Request.Form("TipoHouse"), AWBNumber)
                                end if
                            end if

                        end if                        

                        MAWBID = 0

					    Set rst = Conn.Execute("SELECT AWBID FROM " & AwbTable & " WHERE AWBNumber='" & AWBNumber & "' AND HAWBNumber=''") 
					    if Not rst.EOF then
					        MAWBID = CheckNum(rst(0))                                        
					    end if
					    CloseOBJ rst

                    else
                        ValidaAWBNumber = true
                    end if

                End Function



                Function FinalizaGuia(Action, AWBNumber, ObjectID)

                    'response.write ( "(" & AWBNumber & ")(" & Request.Form("AWBNumber") & ")" ) 

                    '2016-03-30 solicitado por Carlos                    
                    'Este proceso queda para import export / insert update y a toda la region
					if trim(Request.Form("Routing")) <> "" then
                        Dim Conn2
                        OpenConn2 Conn2                        

                        SQLQuery = "UPDATE routings SET activo=false, bl_id = '" & ObjectID & "', no_bl = '" & Request.Form("AWBNumber") 
                        if Request.Form("HAWBNumber") <> "" then
                            SQLQuery = SQLQuery & "/" & trim(Request.Form("HAWBNumber")) 
                        end if
                        SQLQuery = SQLQuery & "' WHERE id_routing IN (" & CheckNum(Request.Form("RoutingID")) & "," & CheckNum(Request.Form("routing_seg")) & "," & CheckNum(Request.Form("routing_adu")) & "," & CheckNum(Request.Form("routing_ter")) & ")"
                        'response.write SQLQuery & "<br>"
                        Conn2.Execute(SQLQuery)
                        
                        CloseOBJ Conn2
                    end if
						          
                    if trim(Request.Form("Routing")) = "" then
                        'se cambio aca 2016-12-15, estaba en el if de abajo
                        if CheckNum(Request.Form("ShipperID")) <> 0 or CheckNum(Request.Form("ConsignerID")) <> 0 then
                            'Actualizando la BD Master indicando la fecha y tipo de servicio que realizo el cliente y el shipper
                            OpenConn2 ConnMaster		                
                            SQLQuery = "update clientes set id_estatus=1, ultima_fecha_descarga='" & ConvertDate(Now,2) & "', ultimo_tipo_movimiento=1 where id_cliente in (" & CheckNum(Request.Form("ShipperID")) & "," & CheckNum(Request.Form("ConsignerID")) & ")"
                            'response.write SQLQuery & "<br>"
	                        ConnMaster.Execute(SQLQuery)
                            CloseOBJ ConnMaster
                        end if
                    end if
                                                                    


                    if AwbType = 1 and Request.Form("Countries") = "GT" then 'export
                                          
                        'Dim Conn2		
                        'if trim(Request.Form("Routing")) <> "" then
                        '    OpenConn2 Conn2
                        '    'SQLQuery = "update routings set activo=false where routing='" & trim(Request.Form("Routing")) & "'"
                        '    'response.write SQLQuery & "<br>"
                        '    Conn2.Execute(SQLQuery)                            
                        '    CloseOBJ Conn2
                        'end if


                                            
                    
                        if Request.Form("CarrierID") <> Request.Form("CarrierIDAnt") and Request.Form("Countries") = "GT" then 

                            'SI TENIA ALGUN REGISTRO VINCULADO 
                            'SQLQuery = "SELECT GuideID FROM Guides WHERE GuideStatus='1' AND GuideCarrierID='" &  Request.Form("CarrierIDAnt") & "' AND GuideNumber='" & Request.Form("AWBNumberAnt") & "' ORDER BY GuideNumber DESC LIMIT 0,1"
                            'response.write SQLQuery & "<br>"

     
                            dim ContAwb
                            
                            if Request.Form("AWBNumberAnt") <> "" and Request.Form("CarrierIDAnt") <> "" then

                                'response.write "DATO NO NUEVO<br>"
                                ContAwb = 0
                                SQLQuery = "SELECT COUNT(AwbNumber) FROM " & AwbTable & " WHERE AwbNumber='" & Request.Form("AWBNumberAnt") & "' AND CarrierID='" &  Request.Form("CarrierIDAnt") & "' AND HAwbNumber <> '' "
                                'response.write SQLQuery & "<br>"
                                Set rst = Conn.Execute(SQLQuery)		                
                                If Not rst.EOF Then  
                                    ContAwb = CInt(rst(0))
                                end if
                                CloseOBJ rst

                                if ContAwb = 0 then 'si no hay datos en awb marca como no usado
                                    'response.write "Guia la marca como no usada<br>"
                                    SQLQuery = "UPDATE Guides SET GuideStatus='0', MasterDate=now(), MasterUser=" & Session("OperatorID") & " WHERE GuideStatus='1' AND GuideNumber='" & Request.Form("AWBNumberAnt") & "' AND GuideCarrierID='" &  Request.Form("CarrierIDAnt") & "'"
                                    'response.write SQLQuery & "<br>"
                                    Conn.Execute(SQLQuery)
		                        end if   
                                
                            end if


                            ContAwb = 0
                            'response.write "TIENE HOUSE?<br>"
                            SQLQuery = "SELECT COUNT(AwbNumber) FROM " & AwbTable & " WHERE AwbNumber='" & AWBNumber & "' AND CarrierID='" &  Request.Form("CarrierID") & "' AND HAwbNumber <> '' "
                            'response.write SQLQuery & "<br>"
                            Set rst = Conn.Execute(SQLQuery)		                
                            If Not rst.EOF Then  
                                ContAwb = CInt(rst(0))
                            end if
                            CloseOBJ rst   

                            if ContAwb > 0 then 'si hay datos en awb marca como usado
                                SQLQuery = "UPDATE Guides SET GuideStatus='1', MasterDate=now(), MasterAWB='" & ObjectID  & "', MasterUser=" & Session("OperatorID") & " WHERE GuideStatus='0' AND GuideCarrierID='" &  Request.Form("CarrierID") & "' AND GuideNumber='" & AWBNumber & "'"
                                'response.write "Guia la marca como usada<br>"
                                'response.write SQLQuery & "<br>"
                                Conn.Execute(SQLQuery)      
                                if  AWBNumber <> Request.Form("AWBNumber") then
                                    JavaMsg = "El AWBNumber cambio a " & AWBNumber
                                end if
                            end if

                        end if      
                                    
                    end if

                    FinalizaGuia = true

                End Function
                
				
               
                if GroupID = 1 then                    
                    ResValidaGuia = ValidaAWBNumber(AwbType)                    
                end if

                                 

				  Select Case GroupID
				  Case 7, 8, 10, 11
					  Set rs = Server.CreateObject("ADODB.Recordset")
				      rs.Open RsFilter, Conn, 2, 3, 1

                  
                  'Case 21
                  '    Set rs = Server.CreateObject("ADODB.Recordset")
                  '    SQLQuery = QuerySelect & " " & TableName & " WHERE " & RsFilter
				  '    rs.Open SQLQuery, Conn, 2, 3, 1
                  '
                  '    response.write( "(" & rs("MasterDate") & ")" )
                  '    response.write(SQLQuery)
                  
				  Case Else
					  Set rs = Server.CreateObject("ADODB.Recordset")
                      SQLQuery = "select * from " & TableName & " where " & RsFilter

                      'response.write(SQLQuery & "<br>" & Action)

				      rs.Open SQLQuery, Conn, 2, 3, 1
                      if not rs.EOF then
                          if GroupID = 1 then

                            'response.write(SQLQuery & "<br>")
                            
                            
                          end if
                      end if

                      'response.write(SQLQuery & "<br>")

					  'openTable Conn, TableName, rs 'Abriendo la base de Datos
					  'rs.Filter = RsFilter                      
				  End Select	


                dim ObjectID2 '2016-09-05

                ObjectID2 = ObjectID
                  
				  
				  Select Case Action
                  Case 1 'Insert
                        If rs.EOF Then 'Si no existe el atributo, puede ingresarlo
                            
							Select Case GroupID 'Actualizando datos de la tabla master (Clientes, Shippers, Exporters)
                            case 1                  
                                if ResValidaGuia = true then
                                    'response.write ( "(ResValidaGuia=" & ResValidaGuia & ")")
                                    SaveInfo Conn, rs, Action, GroupID, CreatedDate, CreatedTime, AwbType                                
                                    closeOBJ rs                                
                                end if
							Case 21								
								if CheckAWBNumber(Conn, guia99) = 0 then
									SaveInfo Conn, rs, Action, GroupID, CreatedDate, CreatedTime, AwbType
								else
									JavaMsg = "No. de AWBNumber " & guia99 & " ya grabado"
									Action = 0                                        
								end if 
                            case else
                                SaveInfo Conn, rs, Action, GroupID, CreatedDate, CreatedTime, AwbType                                
							End Select
                            closeOBJ rs

                            SQLQuery = "select " & ObjectName & " from " & TableName & " where " & WhereSQL & CreatedTime
                            'response.write "After Insert " & SQLQuery & "<br><br>"
                            Set rs = Conn.Execute(SQLQuery)                                
							If Not rs.EOF Then        
                                                                                                            
                                ObjectID = CheckNum(rs(0))
								Select Case GroupID 'Actualizando datos de la tabla master (Clientes, Shippers, Exporters)
                                Case 1                                							
                                    rst = FinalizaGuia(Action, AWBNumber, ObjectID)

								Case 7, 10									
									AddressID = SaveMaster (Conn, ObjectID)
								End Select
                            End If
                            CloseOBJ rs
                             
	                    Else
                            JavaMsg = "La informacion ya existe"
                        End If
                  Case 2 'Update
						CreatedTime = CreatedTime + 1 

						If Not rs.EOF Then 'Si existe el atributo, puede actualizarlo
							Select Case GroupID 'Actualizando datos de la tabla master (Clientes, Shippers, Exporters)
                            Case 1
                                    if ResValidaGuia = true then                                
                                        SaveInfo Conn, rs, Action, GroupID, CreatedDate, CreatedTime, AwbType
                                    end if
                            Case 21
                                if Request.Form("GuiaAnt") <> guia99 then
                                    if CheckAWBNumber(Conn, guia99) = 0 then
                                        SaveInfo Conn, rs, Action, GroupID, CreatedDate, CreatedTime, AwbType
                                    else
                                        JavaMsg = "No. de HAWBNumber " & guia99 & " ya grabado"
                                        Action = 0                                        
                                    end if 
                                else                                    
                                    SaveInfo Conn, rs, Action, GroupID, CreatedDate, CreatedTime, AwbType                                    
                                end if 
                                                               
							Case Else
                                SaveInfo Conn, rs, Action, GroupID, CreatedDate, CreatedTime, AwbType
							End Select
                            CloseOBJ rs

							Select Case GroupID 'Actualizando datos de la tabla master (Clientes, Shippers, Exporters)
                            Case 1
                                rst = FinalizaGuia(Action, AWBNumber, ObjectID)

							Case 7, 10									
								AddressID = SaveMaster (Conn, ObjectID)
							End Select

                        Else
                            JavaMsg = "La informacion no existe"
                        End If
	
                  Case 3 'Delete
                        If Not rs.EOF Then 'Si existe el atributo, puede borrarlo    

                              rs.Delete
                              
							  Select Case GroupID
							  Case 1 'Awb

							  		if AWBType = 1 then 'esto se pasa abajo 2016-09-05
										SQLQuery = "Delete from ChargeItems where AWBID=" & ObjectID & " and DocTyp=0"
									else
                                        SQLQuery = "Delete from ChargeItems where AWBID=" & ObjectID & " and DocTyp=1"
									end if

                                    'response.write SQLQuery & "<br>"
                                    Conn.Execute(SQLQuery)

							  		'if AWBType <> 1 then 									
									'	Conn.Execute("Delete from ChargeItems where AWBID=" & ObjectID & " and DocTyp=1")
									'end if


                                    if trim(Request.Form("Routing")) <> "" then

                                        'response.write ( "(" & Request.Form("Seguro") & ")(" & Request.Form("routing_seg") & ")<br>" )
                                        
                                        OpenConn2 Conn2

                                        IF trim(Request.Form("Seguro")) = "0" THEN
                                            SQLQuery = "UPDATE routings SET bl_id = 0, no_bl = '', activo=true WHERE id_routing='" & CheckNum(Request.Form("RoutingID")) & "'"
                                            'response.write SQLQuery & "<br>"
                                            Conn2.Execute(SQLQuery)    
                                        END IF

                                        if CheckNum(Request.Form("routing_seg")) > 0 then 
                                            SQLQuery = "UPDATE routings SET bl_id = 0, no_bl = '' WHERE id_routing='" & CheckNum(Request.Form("routing_seg")) & "'"
                                            'response.write SQLQuery & "<br>"
                                            Conn2.Execute(SQLQuery)    
                                        end if

                                        if CheckNum(Request.Form("routing_adu")) > 0 then 
                                            SQLQuery = "UPDATE routings SET bl_id = 0, no_bl = '' WHERE id_routing='" & CheckNum(Request.Form("routing_adu")) & "'"
                                            'response.write SQLQuery & "<br>"
                                            Conn2.Execute(SQLQuery)    
                                        end if

                                        if CheckNum(Request.Form("routing_ter")) > 0 then 
                                            SQLQuery = "UPDATE routings SET bl_id = 0, no_bl = '' WHERE id_routing='" & CheckNum(Request.Form("routing_ter")) & "'"
                                            'response.write SQLQuery & "<br>"
                                            Conn2.Execute(SQLQuery)    
                                        end if
                                                                    
                                        CloseOBJ Conn2
                                        
                                    end if


							  Case 2 'Carrier
							  		Conn.Execute("Delete from CarrierDepartures where CarrierID=" & ObjectID)
									Conn.Execute("Delete from CarrierRanges where CarrierID=" & ObjectID)
							  		Conn.Execute("Delete from CarrierRates where CarrierID=" & ObjectID)
									OpenConn2 ConnMaster
										ConnMaster.execute("Delete from carriers where carrier_id=" & ObjectID)
									CloseOBJ ConnMaster
							  Case 9 'Aeropuertos
							  		Conn.Execute("Delete from CarrierDepartures where AirportID=" & ObjectID)
							  		Conn.Execute("Delete from CarrierRates where AirportDepID=" & ObjectID & " or AirportDesID=" & ObjectID)
							  Case 13 'Rangos
							  		Conn.Execute("Delete from CarrierRanges where RangeID=" & ObjectID)
							  		Conn.Execute("Delete from CarrierRates where RangeID=" & ObjectID)
							  end Select
							  ObjectID = 0
                        Else
                            JavaMsg = "La informacion no existe"
                        End If
						Name = ""
						Countries = ""
						AirportID = 0
						CarrierID = 0
						RangeID = 0
						AirportCode = ""
						CommodityCode = ""
						CurrencyCode = ""
						Val = ""
                  End Select


                  if GroupID = 1 then
					if AWBType = 1 then 'esto se pasa abajo 2016-09-05
SQLQuery = "UPDATE Awb SET Weights = (SELECT round(SUM(Weights),2) FROM (SELECT * FROM Awb) x WHERE AwbNumber = '" & AWBNumber & "' AND HAwbNumber <> '') WHERE AwbNumber = '" & AWBNumber & "' AND HAwbNumber = ''"
					else
SQLQuery = "UPDATE Awbi SET Weights = (SELECT round(SUM(Weights),2) FROM (SELECT * FROM Awbi) x WHERE AwbNumber = '" & AWBNumber & "' AND HAwbNumber <> '') WHERE AwbNumber = '" & AWBNumber & "' AND HAwbNumber = ''"
					end if 
                    response.write SQLQuery & "<br>"
                    Conn.Execute(SQLQuery)   
                  end if


                  closeOBJ rs
        	End If
			
			if JavaMsg <> "" then
				Response.write "<SCRIPT>alert('" & JavaMsg & "');</SCRIPT>"
			end if

			If Action <> 4 then
				
				Select Case GroupID
				Case 7, 10
					if AddressID = 0 then
						AddressID=CheckNum(Request("AID"))
						if AddressID = 0 then
							AddressID = CheckNum(Request("AddressID"))
						end if
					end if
					SQLQuery = QuerySelect & " and c." & ObjectName & "=" & ObjectID & " and d.id_direccion=" & AddressID
				Case 8, 11
					SQLQuery = QuerySelect & TableName & " where " & ObjectName & "=" & ObjectID	
				Case 21
                    if Action = 1 or Action = 2 then
                        ObjectID = 0 '2015-03-23 despues de insert queda en modo insert
                    end if
					SQLQuery = QuerySelect & TableName & " where " & ObjectName & "=" & ObjectID	

				Case 22
					SQLQuery = QuerySelect & TableName & " where " & ObjectName & "=" & ObjectID	

				Case Else
					SQLQuery = QuerySelect & TableName & " where " & ObjectName & "=" & ObjectID & " and " & WhereSQL & CreatedTime
				End Select
				
                'response.write SQLQuery & "<br>"

				Set rs = Conn.Execute(SQLQuery)
                'response.write ("After Execute<br>")
				If Not rs.EOF Then
                    'response.write ("In..<br>")
					aTableValues = rs.GetRows
					CountTableValues = rs.RecordCount
				End If
                closeOBJ rs
				'closeOBJs rs, Conn
				closeOBJ Conn


                'dim AWBiID2                
                'if GroupID = 1 And AwbType = 1 and (Action = 1 or Action = 2 or Action = 3) then '2016-09-06
                '    AWBiID2 = -1
                'else
                '   closeOBJ Conn
                'end if

                'response.write ("After Evaluate<br>")

				Select Case GroupID
				Case 1 'Awb %>                    
					<%if AwbType = 1 then 'Export%>

                        <!--#include file=Awb.asp--> 

                        <%  
                        'dim tmpAction
                        'tmpAction = Action
                        'Action = 1                      
                        'AGREGADO EL 05-09-2016 replica de export a import
                        if (Action = 1 or Action = 2 or Action = 3) and Request.Form("replica") = "Consolidado" then
                        
                            openConn Conn

                            if ObjectID2 = 0 then
                                ObjectID2 = ObjectID
                            end if

                            'response.write ("(ObjectID2=" & ObjectID2 & ")<BR>")

                            dim countries2 'SE OBTIENE PAIS DESTINO PARA LA REPLICA EN IMPORT
                            SQLQuery = "SELECT Country FROM Airports WHERE AirportID = " & CheckNum(Request.Form("AirportDesID")) & " AND Expired = 0"
                            'response.write (SQLQuery & "<BR>")
                            Set rs = Conn.Execute(SQLQuery)
                            if Not rs.EOF then
                                countries2 = rs(0) 
                                
                                if InStr(Countries,"LTF") then 'DEBE CONCATENAR LTF, cuando lo traiga ?
                                    countries2 = Left(countries2,2) & Right(Countries,3)                                    
                                end if

                            end if

                            'PAIS ORIGEN DEBE SER DE LA REGION ?

                            'PAIS CREACION Countries
                                     
                            dim AWBiID2                   
                            AWBiID2 = InStr("*GT,SV,HN,NI,CR,PA,BZ,GTLTF,SVLTF,HNLTF,NILTF,CRLTF,PALTF,BZLTF,N1",countries2)    

                            'response.write ("(pais=" & AWBiID2  & ")(" & Request.Form("replica")  & ")<BR>")

                            if AWBiID2 > 0 and Request.Form("replica") = "Consolidado" then 'SI ES DE LA REGION Y SI ES CONSOLIDADO

                                AWBiID2 = 0 'inicializa nuevamente

                                dim InterCompanyID2, InvoiceID2

                                'response.write (Session("Countries") & "<BR>")
                                SQLQuery = "SELECT ifnull(sum(ChargeID),0), ifnull(sum(InterCompanyID),0), ifnull(sum(InvoiceID),0) FROM ChargeItems WHERE AWBID = " & ObjectID2 & " AND DocTyp = 0 AND Expired=0 AND CreatedDate > '2016-12-07'"                                
                                'response.write (SQLQuery & "<BR>")
                                Set rs = Conn.Execute(SQLQuery)
                                if Not rs.EOF then

                                    'response.write ( rs(0) & " " & rs(1) & " " & rs(2) & "<br>")
                                
                                    InterCompanyID2 = rs(1)
                                    InvoiceID2 = rs(2)

                                    if InterCompanyID2 = 0 and InvoiceID2 = 0 then 'NO ES InterCompanyID Y NO TIENE InvoiceID EL EXPORT 'rs(0) > 0 and 
                                         
                                        if Action = 1 then 'ES INSERT
                                        
                                            'response.write ( "1.(Action=" & Action & ")(ObjectID2=" & ObjectID2 & ")<BR>" )

SQLQuery = "INSERT INTO Awbi (CreatedDate, CreatedTime, " & _
"Expired, AwbNumber, AccountShipperNo, ShipperData, AccountConsignerNo, ConsignerData, AgentData, AccountInformation, IATANo, AccountAgentNo, AirportDepID, RequestedRouting, AirportToCode1, CarrierID, AirportToCode2, AirportToCode3, CarrierCode2, CarrierCode3, CurrencyID, ChargeType, ValChargeType, OtherChargeType, DeclaredValue, AduanaValue, AirportDesID, FlightDate1, FlightDate2, SecuredValue, HandlingInformation, Observations, NoOfPieces, Weights, WeightsSymbol, Commodities, ChargeableWeights, CarrierRates, Carriersubtot, NatureQtyGoods, TotNoOfPieces, TotWeight, TotCarrierRate, TotChargeWeightPrepaid, TotChargeWeightCollect, TotChargeValuePrepaid, TotChargeValueCollect, TotChargeTaxPrepaid, TotChargeTaxCollect, AnotherChargesAgentPrepaid, AnotherChargesAgentCollect, AnotherChargesCarrierPrepaid, AnotherChargesCarrierCollect, TotPrepaid, TotCollect, TerminalFee, CustomFee, FuelSurcharge, SecurityFee, PBA, TAX, AdditionalChargeName1, AdditionalChargeVal1, AdditionalChargeName2, AdditionalChargeVal2, Invoice, ExportLic, AgentContactSignature, CommoditiesTypes, TotWeightChargeable, Instructions, Agentsignature, AdditionalChargeName3, AdditionalChargeVal3, AdditionalChargeName4, AdditionalChargeVal4, " & _ 
"Countries, HAwbNumber, ReservationDate, DeliveryDate, DepartureDate, Comment, AdditionalChargeName5, AdditionalChargeVal5, AdditionalChargeName6, AdditionalChargeVal6, Comment2, ArrivalDate, HDepartureDate, Cont, Destinity, TotalToPay, Concept, FiscalFactory, ArrivalAttn, ArrivalFlight, Comment3, DisplayNumber, AdditionalChargeName7, AdditionalChargeVal7, AdditionalChargeName8, AdditionalChargeVal8, WType, " & _
"AdditionalChargeName9, AdditionalChargeVal9, AdditionalChargeName10, AdditionalChargeVal10, ShipperID, ConsignerID, AgentID, SalespersonID, ShipperAddrID, ConsignerAddrID, AgentAddrID, AdditionalChargeName11, AdditionalChargeVal11, AdditionalChargeName12, AdditionalChargeVal12, AdditionalChargeName13, AdditionalChargeVal13, AdditionalChargeName14, AdditionalChargeVal14, AdditionalChargeName15, AdditionalChargeVal15, Voyage, PickUp, Intermodal, SedFilingFee, RoutingID, Routing, ManifestNumber, " & _
"CalcAdminFee, AWBDate, CTX, TCTX, TPTX, UserID, Closed, MAWBID, " & _
"rep_exp, ConsignerColoader, ShipperColoader, AgentNeutral, ManifestNo, MonitorArrivalDate, ClientCollectID, ClientsCollect, id_coloader, TotCarrierRate_Routing, FuelSurcharge_Routing, SecurityFee_Routing, " & _
"PickUp_Routing, SedFilingFee_Routing, Intermodal_Routing, PBA_Routing, " & _
"AdditionalChargeName1_Routing, AdditionalChargeName2_Routing, AdditionalChargeName3_Routing, AdditionalChargeName4_Routing, AdditionalChargeName5_Routing, AdditionalChargeName6_Routing, AdditionalChargeName7_Routing, AdditionalChargeName8_Routing, AdditionalChargeName9_Routing, AdditionalChargeName10_Routing, AdditionalChargeName11_Routing, AdditionalChargeName12_Routing, AdditionalChargeName13_Routing, AdditionalChargeName14_Routing, AdditionalChargeName15_Routing, ReplicaAwbID, id_cliente_order, id_cliente_orderData " & _
") (SELECT NOW(), DATE_FORMAT(NOW(),'%H%i%s'), " & _
"Expired, AwbNumber, AccountShipperNo, ShipperData, AccountConsignerNo, ConsignerData, AgentData, AccountInformation, IATANo, AccountAgentNo, AirportDepID, RequestedRouting, AirportToCode1, CarrierID, AirportToCode2, AirportToCode3, CarrierCode2, CarrierCode3, CurrencyID, ChargeType, ValChargeType, OtherChargeType, DeclaredValue, AduanaValue, AirportDesID, FlightDate1, FlightDate2, SecuredValue, HandlingInformation, Observations, NoOfPieces, Weights, WeightsSymbol, Commodities, ChargeableWeights, CarrierRates, Carriersubtot, NatureQtyGoods, TotNoOfPieces, TotWeight, TotCarrierRate, TotChargeWeightPrepaid, TotChargeWeightCollect, TotChargeValuePrepaid, TotChargeValueCollect, TotChargeTaxPrepaid, TotChargeTaxCollect, AnotherChargesAgentPrepaid, AnotherChargesAgentCollect, AnotherChargesCarrierPrepaid, AnotherChargesCarrierCollect, TotPrepaid, TotCollect, TerminalFee, CustomFee, FuelSurcharge, SecurityFee, PBA, TAX, AdditionalChargeName1, AdditionalChargeVal1, AdditionalChargeName2, AdditionalChargeVal2, Invoice, ExportLic, AgentContactSignature, CommoditiesTypes, TotWeightChargeable, Instructions, Agentsignature, AdditionalChargeName3, AdditionalChargeVal3, AdditionalChargeName4, AdditionalChargeVal4, " & _ 
"'" & countries2 & "', HAwbNumber, ReservationDate, DeliveryDate, DepartureDate, Comment, AdditionalChargeName5, AdditionalChargeVal5, AdditionalChargeName6, AdditionalChargeVal6, Comment2, ArrivalDate, HDepartureDate, Cont, Destinity, TotalToPay, Concept, FiscalFactory, ArrivalAttn, ArrivalFlight, Comment3, DisplayNumber, AdditionalChargeName7, AdditionalChargeVal7, AdditionalChargeName8, AdditionalChargeVal8, WType, " & _
"AdditionalChargeName9, AdditionalChargeVal9, AdditionalChargeName10, AdditionalChargeVal10, ShipperID, ConsignerID, AgentID, SalespersonID, ShipperAddrID, ConsignerAddrID, AgentAddrID, AdditionalChargeName11, AdditionalChargeVal11, AdditionalChargeName12, AdditionalChargeVal12, AdditionalChargeName13, AdditionalChargeVal13, AdditionalChargeName14, AdditionalChargeVal14, AdditionalChargeName15, AdditionalChargeVal15, Voyage, PickUp, Intermodal, SedFilingFee, RoutingID, Routing, ManifestNumber, " & _
"CalcAdminFee, AWBDate, CTX, TCTX, TPTX, UserID, Closed, MAWBID, " & _
"rep_exp, ConsignerColoader, ShipperColoader, AgentNeutral, ManifestNo, MonitorArrivalDate, ClientCollectID, ClientsCollect, id_coloader, TotCarrierRate_Routing, FuelSurcharge_Routing, SecurityFee_Routing, " & _
"PickUp_Routing, SedFilingFee_Routing, Intermodal_Routing, PBA_Routing, " & _
"AdditionalChargeName1_Routing, AdditionalChargeName2_Routing, AdditionalChargeName3_Routing, AdditionalChargeName4_Routing, AdditionalChargeName5_Routing, AdditionalChargeName6_Routing, AdditionalChargeName7_Routing, AdditionalChargeName8_Routing, AdditionalChargeName9_Routing, AdditionalChargeName10_Routing, AdditionalChargeName11_Routing, AdditionalChargeName12_Routing, AdditionalChargeName13_Routing, AdditionalChargeName14_Routing, AdditionalChargeName15_Routing, '" & ObjectID2 & "', id_cliente_order, id_cliente_orderData " & _
"FROM Awb WHERE AwbID = '" & ObjectID2 & "')"
                                            'response.write ( "2.(Action=" & Action & ")(ObjectID2=" & ObjectID2 & ")<BR>" )
                                            'response.write ( SQLQuery & "<BR>" )
                                            Set rs = Conn.Execute(SQLQuery) 'REPLICA EXPORT A IMPORT
                                            
                                            SQLQuery = "SELECT LAST_INSERT_ID()" 'ULTIMO INSERT AL IMPORT
                                            'response.write (SQLQuery & "<BR>")
                                            Set rs = Conn.Execute(SQLQuery)                                            
                                            if Not rs.EOF then                                                
                                                AWBiID2 = CheckNum(rs(0))      
                                                'response.write ("Export Replica Correctamente<br>")
                                            end if
                                        
                                        else

                                            SQLQuery = "SELECT AwbID FROM Awbi WHERE ReplicaAwbID = " & ObjectID2
                                            'response.write (SQLQuery & "<BR>")
                                            Set rs = Conn.Execute(SQLQuery) 'SI NO ES INSERT LEE LA REPLICA DEL IMPORT
                                            if Not rs.EOF then                                            
                                                AWBiID2 = CheckNum(rs(0))
                                            end if
                                            'SI NO ENCUENTRA LA REPLICA EL AWBiID2 VIENE EN CERO

                                        end if
                                
                                        'response.write (AWBiID2 & "<BR>")

                                        if AWBiID2 > 0 then 'debe traer el id insertado o el leido de import

                                            SQLQuery = "SELECT ifnull(sum(ChargeID),0), ifnull(sum(InterCompanyID),0), ifnull(sum(InvoiceID),0) FROM ChargeItems WHERE AWBID = " &  AWBiID2 & " AND DocTyp = 1 AND Expired=0 AND CreatedDate > '2016-12-07'"
                                            'response.write (SQLQuery & "<BR><br>")
                                            Set rs = Conn.Execute(SQLQuery)
                                            if Not rs.EOF then
                                                
                                                InterCompanyID2 = rs(1)
                                                InvoiceID2 = rs(2)

                                                'response.write ( rs(0) & " " & InterCompanyID2 & " " & InvoiceID2 & "<br>")

                                                if rs(0) > 0 then 'SI TIENE CARGOS EN EL IMPORT LOS BORRA
                                                    'SQLQuery = "DELETE FROM ChargeItems WHERE ReplicaAwbID = " & ObjectID2
                                                    SQLQuery = "DELETE FROM ChargeItems WHERE AWBID = " & AWBiID2 & " AND DocTyp = 1 AND Expired=0"
                                                    'response.write (SQLQuery & "<BR>")
                                                    Set rs = Conn.Execute(SQLQuery)                                                    
                                                end if

                                                if InterCompanyID2 = 0 and InvoiceID2 = 0 then

                                                    if Action = 2 then 'SI ES UPDATE DEL EXPORT Y HAY REPLICA, LO ACTUALIZA                                                   
                                                        SQLQuery = "UPDATE Awbi a JOIN Awb b ON a.ReplicaAwbID = b.AwbID SET a.CreatedTime=DATE_FORMAT(NOW(),'%H%i%s'), " & _
                                                        "a.Expired=b.Expired, a.AwbNumber=b.AwbNumber, a.AccountShipperNo=b.AccountShipperNo, a.ShipperData=b.ShipperData, a.AccountConsignerNo=b.AccountConsignerNo, a.ConsignerData=b.ConsignerData, a.AgentData=b.AgentData, a.AccountInformation=b.AccountInformation, a.IATANo=b.IATANo, a.AccountAgentNo=b.AccountAgentNo, a.AirportDepID=b.AirportDepID, a.RequestedRouting=b.RequestedRouting, a.AirportToCode1=b.AirportToCode1, a.CarrierID=b.CarrierID, a.AirportToCode2=b.AirportToCode2, a.AirportToCode3=b.AirportToCode3, a.CarrierCode2=b.CarrierCode2, a.CarrierCode3=b.CarrierCode3, a.CurrencyID=b.CurrencyID, a.ChargeType=b.ChargeType, a.ValChargeType=b.ValChargeType, a.OtherChargeType=b.OtherChargeType, a.DeclaredValue=b.DeclaredValue, a.AduanaValue=b.AduanaValue, a.AirportDesID=b.AirportDesID, a.FlightDate1=b.FlightDate1, a.FlightDate2=b.FlightDate2, a.SecuredValue=b.SecuredValue, a.HandlingInformation=b.HandlingInformation, a.Observations=b.Observations, a.NoOfPieces=b.NoOfPieces, a.Weights=b.Weights, a.WeightsSymbol=b.WeightsSymbol, a.Commodities=b.Commodities, a.ChargeableWeights=b.ChargeableWeights, a.CarrierRates=b.CarrierRates, a.Carriersubtot=b.Carriersubtot, a.NatureQtyGoods=b.NatureQtyGoods, a.TotNoOfPieces=b.TotNoOfPieces, a.TotWeight=b.TotWeight, a.TotCarrierRate=b.TotCarrierRate, a.TotChargeWeightPrepaid=b.TotChargeWeightPrepaid, a.TotChargeWeightCollect=b.TotChargeWeightCollect, a.TotChargeValuePrepaid=b.TotChargeValuePrepaid, a.TotChargeValueCollect=b.TotChargeValueCollect, a.TotChargeTaxPrepaid=b.TotChargeTaxPrepaid, a.TotChargeTaxCollect=b.TotChargeTaxCollect, a.AnotherChargesAgentPrepaid=b.AnotherChargesAgentPrepaid, a.AnotherChargesAgentCollect=b.AnotherChargesAgentCollect, a.AnotherChargesCarrierPrepaid=b.AnotherChargesCarrierPrepaid, a.AnotherChargesCarrierCollect=b.AnotherChargesCarrierCollect, a.TotPrepaid=b.TotPrepaid, a.TotCollect=b.TotCollect, a.TerminalFee=b.TerminalFee, a.CustomFee=b.CustomFee, a.FuelSurcharge=b.FuelSurcharge, a.SecurityFee=b.SecurityFee, a.PBA=b.PBA, a.TAX=b.TAX, a.AdditionalChargeName1=b.AdditionalChargeName1, a.AdditionalChargeVal1=b.AdditionalChargeVal1, a.AdditionalChargeName2=b.AdditionalChargeName2, a.AdditionalChargeVal2=b.AdditionalChargeVal2, a.Invoice=b.Invoice, a.ExportLic=b.ExportLic, a.AgentContactSignature=b.AgentContactSignature, a.CommoditiesTypes=b.CommoditiesTypes, a.TotWeightChargeable=b.TotWeightChargeable, a.Instructions=b.Instructions, a.Agentsignature=b.Agentsignature, a.AdditionalChargeName3=b.AdditionalChargeName3, a.AdditionalChargeVal3=b.AdditionalChargeVal3, a.AdditionalChargeName4=b.AdditionalChargeName4, a.AdditionalChargeVal4=b.AdditionalChargeVal4, a.Countries='" & countries2 & "', a.HAwbNumber=b.HAwbNumber, a.ReservationDate=b.ReservationDate, a.DeliveryDate=b.DeliveryDate, a.DepartureDate=b.DepartureDate, a.Comment=b.Comment, a.AdditionalChargeName5=b.AdditionalChargeName5, a.AdditionalChargeVal5=b.AdditionalChargeVal5, a.AdditionalChargeName6=b.AdditionalChargeName6, a.AdditionalChargeVal6=b.AdditionalChargeVal6, a.Comment2=b.Comment2, a.ArrivalDate=b.ArrivalDate, a.HDepartureDate=b.HDepartureDate, a.Cont=b.Cont, a.Destinity=b.Destinity, a.TotalToPay=b.TotalToPay, a.Concept=b.Concept, a.FiscalFactory=b.FiscalFactory, a.ArrivalAttn=b.ArrivalAttn, a.ArrivalFlight=b.ArrivalFlight, a.Comment3=b.Comment3, a.DisplayNumber=b.DisplayNumber, a.AdditionalChargeName7=b.AdditionalChargeName7, a.AdditionalChargeVal7=b.AdditionalChargeVal7, a.AdditionalChargeName8=b.AdditionalChargeName8, a.AdditionalChargeVal8=b.AdditionalChargeVal8, a.WType=b.WType, " & _
                                                        "a.AdditionalChargeName9=b.AdditionalChargeName9, a.AdditionalChargeVal9=b.AdditionalChargeVal9, a.AdditionalChargeName10=b.AdditionalChargeName10, a.AdditionalChargeVal10=b.AdditionalChargeVal10, a.ShipperID=b.ShipperID, a.ConsignerID=b.ConsignerID, a.AgentID=b.AgentID, a.SalespersonID=b.SalespersonID, a.ShipperAddrID=b.ShipperAddrID, a.ConsignerAddrID=b.ConsignerAddrID, a.AgentAddrID=b.AgentAddrID, a.AdditionalChargeName11=b.AdditionalChargeName11, a.AdditionalChargeVal11=b.AdditionalChargeVal11, a.AdditionalChargeName12=b.AdditionalChargeName12, a.AdditionalChargeVal12=b.AdditionalChargeVal12, a.AdditionalChargeName13=b.AdditionalChargeName13, a.AdditionalChargeVal13=b.AdditionalChargeVal13, a.AdditionalChargeName14=b.AdditionalChargeName14, a.AdditionalChargeVal14=b.AdditionalChargeVal14, a.AdditionalChargeName15=b.AdditionalChargeName15, a.AdditionalChargeVal15=b.AdditionalChargeVal15, a.Voyage=b.Voyage, a.PickUp=b.PickUp, a.Intermodal=b.Intermodal, a.SedFilingFee=b.SedFilingFee, a.RoutingID=b.RoutingID, a.Routing=b.Routing, a.ManifestNumber=b.ManifestNumber, " & _
                                                        "a.CalcAdminFee=b.CalcAdminFee, a.AWBDate=b.AWBDate, a.CTX=b.CTX, a.TCTX=b.TCTX, a.TPTX=b.TPTX, a.UserID=b.UserID, a.Closed=b.Closed, a.MAWBID=b.MAWBID, " & _
                                                        "a.rep_exp=b.rep_exp, a.ConsignerColoader=b.ConsignerColoader, a.ShipperColoader=b.ShipperColoader, a.AgentNeutral=b.AgentNeutral, a.ManifestNo=b.ManifestNo, a.MonitorArrivalDate=b.MonitorArrivalDate, a.ClientCollectID=b.ClientCollectID, a.ClientsCollect=b.ClientsCollect, a.id_coloader=b.id_coloader, a.TotCarrierRate_Routing=b.TotCarrierRate_Routing, a.FuelSurcharge_Routing=b.FuelSurcharge_Routing, a.SecurityFee_Routing=b.SecurityFee_Routing, " & _
                                                        "a.PickUp_Routing=b.PickUp_Routing, a.SedFilingFee_Routing=b.SedFilingFee_Routing, a.Intermodal_Routing=b.Intermodal_Routing, a.PBA_Routing=b.PBA_Routing, " & _
                                                        "a.AdditionalChargeName1_Routing=b.AdditionalChargeName1_Routing, a.AdditionalChargeName2_Routing=b.AdditionalChargeName2_Routing, a.AdditionalChargeName3_Routing=b.AdditionalChargeName3_Routing, a.AdditionalChargeName4_Routing=b.AdditionalChargeName4_Routing, a.AdditionalChargeName5_Routing=b.AdditionalChargeName5_Routing, a.AdditionalChargeName6_Routing=b.AdditionalChargeName6_Routing, a.AdditionalChargeName7_Routing=b.AdditionalChargeName7_Routing, a.AdditionalChargeName8_Routing=b.AdditionalChargeName8_Routing, a.AdditionalChargeName9_Routing=b.AdditionalChargeName9_Routing, a.AdditionalChargeName10_Routing=b.AdditionalChargeName10_Routing, a.AdditionalChargeName11_Routing=b.AdditionalChargeName11_Routing, a.AdditionalChargeName12_Routing=b.AdditionalChargeName12_Routing, a.AdditionalChargeName13_Routing=b.AdditionalChargeName13_Routing, a.AdditionalChargeName14_Routing=b.AdditionalChargeName14_Routing, a.AdditionalChargeName15_Routing=b.AdditionalChargeName15_Routing, a.id_cliente_order=b.id_cliente_order, a.id_cliente_orderData=b.id_cliente_orderData " & _
                                                        "WHERE b.AwbID = " & ObjectID2
                                                        'response.write (SQLQuery & "<BR>")
                                                        Set rs = Conn.Execute(SQLQuery)
                                                    end if

                                                    if Action = 3 then 'SI BORRA EL EXPORT BORRA EL IMPORT SI NO HAY InterCompanyID2 NI InvoiceID2 
                                                        SQLQuery = "DELETE FROM Awbi WHERE ReplicaAwbID = " & ObjectID2
                                                        'response.write (SQLQuery & "<BR>")
                                                        Set rs = Conn.Execute(SQLQuery)
                                                    else
                                                        'SI NO ES DELETE DEL EXPORT REPLICA LOS CARGOS DE EXPORT A IMPORT
							                            SQLQuery = "INSERT INTO ChargeItems " & _
							                            "(AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, Expired, OverSold, Pos, ServiceID, ServiceName, PrepaidCollect, InvoiceID, AccountType, DocType, CalcInBL, InterChargeType, InterCompanyID, InterGroupID, InterProviderType, ItemName_Routing, batch_id, ReplicaAwbID) " & _
							                            "(SELECT " & AWBiID2 & ", CurrencyID, ItemID, Value, Local, AgentTyp, 1, ItemName, now(), CreatedTime, UserID, Expired, OverSold, Pos, ServiceID, ServiceName, PrepaidCollect, InvoiceID, AccountType, DocType, CalcInBL, InterChargeType, InterCompanyID, InterGroupID, InterProviderType, ItemName_Routing, 0, " & ObjectID2 & " " & _ 
                                                        "FROM ChargeItems WHERE AWBID = " & ObjectID2 & " AND DocTyp = 0 AND Expired=0)"
                                                        'response.write (SQLQuery & "<BR>")                                                    
                                                        Set rs = Conn.Execute(SQLQuery)
                                                        
                                                        'SE HACEN LOS AJUSTES DE RUBROS DE EXPORT A IMPORT
                                                        dim pos
                                                        dim AgentsPos(20)
                                                        AgentsPos(1) = 1 
	                                                    AgentsPos(2) = 2
	                                                    AgentsPos(3) = 6
	                                                    AgentsPos(4) = 7
	                                                    AgentsPos(5) = 9
	                                                    AgentsPos(6) = 10
	                                                    AgentsPos(7) = 11
	                                                    AgentsPos(8) = 12
	                                                    AgentsPos(9) = 13
	                                                    AgentsPos(10) = 14
	                                                    AgentsPos(11) = 15

                                                        '14 : CUSTOM FEE    15 : TERMINAL FEE    116 : PBA
                                                        dim itemids2(5)
                                                        itemids2(1) = 14
                                                        itemids2(2) = 15
                                                        itemids2(3) = 116

                                                        for i = 1 to 3  
                                                        pos = 999
                                                            SQLQuery = "SELECT MAX(b.Pos)+1 FROM ChargeItems b WHERE b.AWBID = " & AWBiID2 & " AND b.DocTyp = 1 AND b.Expired=0 AND AgentTyp = 1"
                                                            'response.write (SQLQuery & "<BR><br>")
                                                            Set rs = Conn.Execute(SQLQuery)
                                                            if Not rs.EOF then
                                                                pos = CheckNum(rs(0))
                                                            end if
                                                            closeOBJ rs

                                                            SQLQuery = "UPDATE ChargeItems SET AgentTyp = 1, Pos = " & pos & " WHERE AWBID = " & AWBiID2 & " AND DocTyp = 1 AND Expired=0 AND ItemID = " & itemids2(i)
                                                            'response.write (SQLQuery & "<BR><br>")
                                                            Conn.Execute(SQLQuery)

                                                            SQLQuery = "SELECT Value, ItemName, ItemName_Routing FROM ChargeItems WHERE AWBID = " & AWBiID2 & " AND DocTyp = 1 AND Expired=0 AND ItemID = " & itemids2(i)
                                                            'response.write (SQLQuery & "<BR><br>")
                                                            Set rs = Conn.Execute(SQLQuery)
                                                            if Not rs.EOF then
                                                                pos = AgentsPos(pos)
                                                                
                                                                if CheckNum(pos) > 0 then
                                                                
                                                                SQLQuery = "UPDATE Awbi SET AdditionalChargeVal" & pos & " = " & Replace(rs(0),",",".") & ", AdditionalChargeName" & pos & " = '" & rs(1) & "', AdditionalChargeName" & pos & "_routing = '" & rs(2) & "' WHERE AWBID = " & AWBiID2 
                                                                'response.write (SQLQuery & "<BR><br>")  
                                                                Conn.Execute(SQLQuery)

                                                                end if
                                                            end if
                                                            closeOBJ rs

                                                        next 

                                                    end if

                                                end if

                                            end if

                                        end if


                                        if Action = 3 then

                                            SQLQuery = "Delete from ChargeItems where AWBID=" & ObjectID2 & " AND DocTyp = 0 AND Expired=0"
                                            'response.write (SQLQuery & "<BR>")
                                            Conn.Execute(SQLQuery) 'BORRA CARGOS DEL EXPORT

                                        end if
                                        
                                        
                                        if AWBiID2 > 0 then 'SI ENCONTRO ID DE INSERT EN IMPORT O ID DE REPLICA REALIZADA ANTES

                                            response.write ("Rubros Replicados Correctamente Import " & AWBiID2 & "<br>")

                                        end if
                                        

                                    end if

                                end if

                            end if
                            
                            closeOBJ Conn

                        end if
                        'Action = tmpAction
                        %>

					<%else 'Import%>
						<!--#include file=Awbi.asp--> 
					<%end if%>
				<% Case 2 'Transportistas - Carriers %>
					<!--#include file=Carriers.asp--> 
				<% Case 3 'Transportistas-Salida %>
					<!--#include file=CarrierDepartures.asp--> 
				<% Case 4 'Confirmacion de Reserva %>
					<!--#include file=Rep_ConfReservation.asp--> 
				<% Case 5 'Transportistas-Rango %>
					<!--#include file=CarrierRanges.asp--> 
				<% Case 6 'House Cargo Manifiesto %>
					<!--#include file=Rep_HouseCargoMan.asp--> 	
				<% Case 7, 10 'Destinatarios - Consigneers, Shippers %>
					<!--#include file=Master.asp--> 
				<% Case 8 'Aeropuertos %>
					<!--#include file=Agents.asp--> 
				<% Case 9 'Aeropuertos %>
					<!--#include file=Airports.asp--> 
				<% Case 11 'Commodities %>
					<!--#include file=Commodities.asp--> 
				<% Case 12 'Monedas %>
					<!--#include file=Currencies.asp--> 
				<% Case 13 'Rangos %>
					<!--#include file=Ranges.asp--> 
				<% Case 14 'Rangos %>
					<!--#include file=Taxes.asp--> 
				<% Case 15 'Arribo %>
					<!--#include file=Rep_Arrival.asp--> 
				<% Case 18 'Rastreo / Tracking %>
                    <!--#include file=Tracking2.asp-->		 
					<!--#include file=Tracking.asp-->		 
				<% Case 21 'Guias %>
					<!--#include file=Guias.asp-->		 
				<% Case 22 'Mediciones %>
					<!--#include file=Mediciones.asp-->		 
				<% End Select
			Else
				Select Case GroupID
				Case 7, 10
					SearchSimilars "nombre_cliente", request.Form("Name"), GroupID, " ", SearchOption
				Case 8
					SearchSimilars "agente", request.Form("Name"), GroupID, " ", SearchOption
				Case 2, 12
					SearchSimilars "Name", request.Form("Name"), GroupID, " ", SearchOption
				Case 9
					SearchSimilars "AirportCode", request.Form("AirportCode"), GroupID, " ", SearchOption
				Case 11
					SearchSimilars "NameES", request.Form("NameES"), GroupID, " ", SearchOption
				End Select %>
				<!--#include file=Similars.asp--> 
			<% End If			
    End If 
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
