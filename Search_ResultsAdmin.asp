<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"

Session.LCID = 4106

Dim GroupID, HTMLCode, HTMLTitle, Table, OrderName, QuerySelect, DateFrom, DateTo 'QueryYearSelect
Dim Option1, Option2, Option3, Option4, Option5, Option6, Option7, Option8, Option9, Option10, AwbTableName
Dim Name, CarrierCode, CarrierID, AirportID, RangeID, AccountNo, IATANo, CurrencyCode, Val, Tax, ResultType
Dim AWBNumber, HAWBNumber, AgentID, ShipperID, AirportDepID, AirportDesID, AirportCode, AwbType, ReportType 
Dim elements, PageCount, AbsolutePage, HTMLHidden, Status, SalespersonID, JavaMsg, i, j, CD, MoreOptions  
Dim MMFrom, MMTo, YYFrom, YYTo, FECHAS

GroupID = CheckNum(Request.Form("GID"))

if GroupID = 0 then
    GroupID = CheckNum(Request("GID")) '2018-02-06
end if

	
if GroupID >= 1 and GroupID <=22 then

    MMFrom = CheckNum(Request.Form("MMFrom"))
    YYFrom = CheckNum(Request.Form("YYFrom"))
    MMTo = CheckNum(Request.Form("MMTo"))
    YYTo = CheckNum(Request.Form("YYTo"))

    Session("DateFrom") = Request.Form("DateFrom")
    Session("DateTo")  = Request.Form("DateTo")

    AwbType = CheckNum(Request.Form("AwbType"))

    FECHAS = " desde " & Request.Form("MMFromText") & " / " & YYFrom & " hasta "  & Request.Form("MMToText") & " /" & YYTo & " "
        
	AbsolutePage = CheckNum(Request.Form("P"))
	if AbsolutePage = 0 then
		 AbsolutePage = 1
	end if
	elements = 5
	PageCount = 0
    Select case GroupID
	case 1, 17, 22
			 OrderName = " order by a.CreatedDate Desc"
			 AwbType = CheckNum(Request.Form("AwbType"))

if AwbType = 0 then
    AwbType = CheckNum(Request("AwbType")) '2018-02-06
end if

			 if AwbType = 1 then
			 	QuerySelect = "select a.AWBID, a.CreatedTime, a.CreatedDate, a.AWBNumber, a.ShipperData, a.Expired, a.HAWBNumber from Awb a"
			 else 
			 	QuerySelect = "select a.AWBID, a.CreatedTime, a.CreatedDate, a.AWBNumber, a.ShipperData, a.Expired, a.HAWBNumber from Awbi a"
			 end if
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>No. de AWB</td><td class=titlelist><b>Embarcador</td><td class=titlelist><b>Status</td>"
			 AWBNumber = Request.Form("AWBNumber")
			 CarrierID = CheckNum(Request.Form("CarrierID"))
			 AgentID = CheckNum(Request.Form("AgentID"))
			 AirportDepID = CheckNum(Request.Form("AirportDepID"))
			 AirportDesID = CheckNum(Request.Form("AirportDesID"))
			 SalespersonID = CheckNum(Request.Form("SalespersonID"))
			
			 if GroupID=1 or GroupID=22 then
			 	Option1 = " a.Countries in " & Session("Countries") & " "
				 if AWBNumber <> "" then
						Option2 = " (a.AWBNumber like '%" & AWBNumber & "%' or a.HAWBNumber like '%" & AWBNumber & "%') "
				 end if
			 else
			 	Option1 = " a.Countries in " & Session("Countries") & " and (a.HAWBNumber='' or a.HAWBNumber=AWBNumber) "
				 if AWBNumber <> "" then
						Option2 = " (a.AWBNumber like '%" & AWBNumber & "%') "
				 end if
			 end if
			 
 			 if CarrierID <> 0 then
			 		Option3 = " a.CarrierID=" & CarrierID & " "
			 end if
 			 if AgentID <> 0 then
			 		Option4 = " a.AgentID=" & AgentID & " "
			 end if
 			 if AirportDepID <> 0 then
			 		Option5 = " a.AirportDepID=" & AirportDepID & " "
			 end if
 			 if AirportDesID <> 0 then
			 		Option6 = " a.AirportDesID=" & AirportDesID & " "
			 end if
 			 if SalespersonID <> 0 then
			 		Option7 = " a.SalespersonID=" & SalespersonID & " "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='AWBNumber' type=hidden value='" & AWBNumber & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='CarrierID' type=hidden value='" & CarrierID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AgentID' type=hidden value='" & AgentID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AirportDepID' type=hidden value='" & AirportDepID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AirportDesID' type=hidden value='" & AirportDesID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AwbType' type=hidden value='" & AwbType & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='SalespersonID' type=hidden value='" & SalespersonID & "'>"
	case 2
			 OrderName = " order by a.Countries, a.Name"
			 QuerySelect = GetSQLSearch (GroupID)
			 HTMLTitle = "<td class=titlelist><b>ID</td><td class=titlelist><b>C&oacute;digo Transportista</td><td class=titlelist><b>Transportista</td><td class=titlelist><b>Status</td>"
			 Name = Request.Form("Name")
			 CarrierCode = Request.Form("CarrierCode")
			 Option1 = " a.Countries in " & Session("Countries") & " "
			 if Name <> "" then
			 		Option2 = " a.Name like '%" & Name & "%' "
			 end if
 			 if CarrierCode <> "" then
			 		Option3 = " a.CarrierCode like '%" & CarrierCode & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='CarrierCode' type=hidden value='" & CarrierCode & "'>"
	case 3
			 OrderName = " order by b.Countries, b.Name, c.Name"
			 QuerySelect = 	"select a.CarrierDepartureID, a.CreatedTime, a.CreatedDate, b.Name, c.Name, c.AirportCode, a.Expired, a.CarrierID, a.AirportID from CarrierDepartures a, Carriers b, Airports c"
			 Option1 = " a.CarrierID=b.CarrierID and a.AirportID=c.AirportID and b.Countries in " & Session("Countries") & " "
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>Transportista</td><td class=titlelist><b>Aeropuerto</td><td class=titlelist><b>C&oacute;digo Aeropuerto</td><td class=titlelist><b>Status</td>"
			 CarrierID = CheckNum(Request.Form("CarrierID"))
			 AirportID = CheckNum(Request.Form("AirportID"))
			 if CarrierID <> 0 then
			 		Option2 = " a.CarrierID=" & CarrierID & " "
			 end if
			 if AirportID <> 0 then
			 		Option3 = " a.AirportID=" & AirportID & " "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name=CarrierID type=hidden value=" & CarrierID & ">"
			 HTMLHidden = HTMLHidden & "<INPUT name=AirportID type=hidden value=" & AirportID & ">"
		 	 elements = 6
	case 4, 6
			 OrderName = " group by a.AWBNumber order by a.CreatedDate Desc"
 			 AwbType = CheckNum(Request.Form("AwbType"))
			 if AwbType = 1 then
			 	QuerySelect = "select a.AWBID, a.CreatedTime, a.CreatedDate, a.HAWBNumber, a.ShipperData, a.ReservationDate from Awb a"
			 else
			 	QuerySelect = "select a.AWBID, a.CreatedTime, a.CreatedDate, a.HAWBNumber, a.ShipperData, a.ReservationDate from Awbi a"
			 end if
			 Option1 = " a.Countries in " & Session("Countries") & " "
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>No. de AWB</td><td class=titlelist><b>Embarcador</td><td class=titlelist><b>Status</td>"
			 AWBNumber = Request.Form("AWBNumber")
			 CarrierID = CheckNum(Request.Form("CarrierID"))
			 AgentID = CheckNum(Request.Form("AgentID"))
			 AirportDepID = CheckNum(Request.Form("AirportDepID"))
			 AirportDesID = CheckNum(Request.Form("AirportDesID"))
			 if AWBNumber <> "" then
			 	if GroupID=4 then
			 		Option2 = " a.AWBNumber like '%" & AWBNumber & "%' and a.HAWBNumber='' "
				else
					Option2 = " a.HAWBNumber like '%" & AWBNumber & "%' "
				end if
			 end if
 			 if CarrierID <> 0 then
			 		Option3 = " a.CarrierID=" & CarrierID & " "
			 end if
 			 if AgentID <> 0 then
			 		Option4 = " a.AgentID=" & AgentID & " "
			 end if
 			 if AirportDepID <> 0 then
			 		Option5 = " a.AirportDepID=" & AirportDepID & " "
			 end if
 			 if AirportDesID <> 0 then
			 		Option6 = " a.AirportDesID=" & AirportDesID & " "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='AWBNumber' type=hidden value='" & AWBNumber & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='CarrierID' type=hidden value='" & CarrierID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AgentID' type=hidden value='" & AgentID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AirportDepID' type=hidden value='" & AirportDepID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AirportDesID' type=hidden value='" & AirportDesID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AwbType' type=hidden value='" & AwbType & "'>"
	case 5
			 OrderName = " order by b.Countries, b.Name, c.Val"
			 QuerySelect = 	"select a.CarrierRangeID, a.CreatedTime, a.CreatedDate, b.Name, c.Val, a.Expired, a.CarrierID, a.RangeID from CarrierRanges a, Carriers b, Ranges c"
			 Option1 = " a.CarrierID=b.CarrierID and a.RangeID=c.RangeID and b.Countries in " & Session("Countries") & " "
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>Transportista</td><td class=titlelist><b>Rango</td><td class=titlelist><b>Status</td>"
			 CarrierID = CheckNum(Request.Form("CarrierID"))
			 RangeID = CheckNum(Request.Form("RangeID"))
			 if CarrierID <> 0 then
			 		Option2 = " a.CarrierID=" & CarrierID & " "
			 end if
			 if RangeID <> 0 then
			 		Option3 = " a.AirportID=" & RangeID & " "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name=CarrierID type=hidden value=" & CarrierID & ">"
			 HTMLHidden = HTMLHidden & "<INPUT name=RangeID type=hidden value=" & RangeID & ">"
	case 7
			 elements = 6
			 OrderName = " order by p.codigo, a.nombre_cliente"
			 QuerySelect = GetSQLSearch (GroupID)
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>Pais</td><td class=titlelist><b>Codigo</td><td class=titlelist><b>Destinatario</td><td class=titlelist><b>Status</td>"
			 Name = Request.Form("Name")
			 Option1 = " a.id_cliente = d.id_cliente " & _
							"and d.id_nivel_geografico = n.id_nivel " & _
							"and n.id_pais = p.codigo " & _
							"and a.es_consigneer = true "' & _
							'"and p.codigo in " & Session("Countries") & " "
			 if Name <> "" then
			 		Option2 = " a.nombre_cliente ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 8
			 elements = 6
			 OrderName = " order by a.agente"
			 QuerySelect = GetSQLSearch (GroupID)
				 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>Codigo</td><td class=titlelist><b>Nombre Agente</td><td class=titlelist><b>Contacto</td><td class=titlelist><b>Status</td>"
			 Name = Request.Form("Name")
			 if Name <> "" then
					Option1 = " a.agente ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 9
			 OrderName = " order by a.Name"
			 QuerySelect = 	GetSQLSearch (GroupID)
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>C&oacute;digo Aeropuerto</td><td class=titlelist><b>Aeropuerto</td><td class=titlelist><b>Status</td>"
			 Name = Request.Form("Name")
			 AirportCode = Request.Form("AirportCode")
			 if Name <> "" then
			 		Option1 = " a.Name like '%" & Name & "%' "
			 end if
			 if AirportCode <> "" then
			 		Option2 = " a.AirportCode like '%" & AirportCode & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name=AirportCode type=hidden value=" & AirportCode & ">"
	case 10
			 elements = 6
			 OrderName = " order by p.codigo, a.nombre_cliente"
			 QuerySelect = GetSQLSearch (GroupID)
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>Pais</td><td class=titlelist><b>Codigo</td><td class=titlelist><b>Embarcador</td><td class=titlelist><b>Status</td>"
			 Name = Request.Form("Name")
			 Option1 = " a.id_cliente = d.id_cliente " & _
							"and d.id_nivel_geografico = n.id_nivel " & _
							"and n.id_pais = p.codigo " & _
							"and a.es_shipper = true "
			 if Name <> "" then
			 		Option2 = " a.nombre_cliente ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 11
			 OrderName = " order by a.NameES"
			 QuerySelect = 	GetSQLSearch (GroupID)
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>C&oacute;digo SCR</td><td class=titlelist><b>Producto</td><td class=titlelist><b>Status</td>"
			 Name = Request.Form("Name")
			 if Name <> "" then
			 		Option1 = " a.NameES ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 12
			 OrderName = " order by a.Name"
			 QuerySelect = 	GetSQLSearch (GroupID)
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>C&oacute;digo Moneda</td><td class=titlelist><b>Moneda</td><td class=titlelist><b>Status</td>"
			 Name = Request.Form("Name")
			 CurrencyCode = Request.Form("CurrencyCode")
			 Option1 = " a.Countries in " & Session("Countries") & " "
			 if Name <> "" then
			 		Option2 = " a.Name like '%" & Name & "%' "
			 end if
			 if CurrencyCode <> "" then
			 		Option3 = " a.CurrencyCode like '%" & CurrencyCode & "%' "
			 end if			 
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='CurrencyCode' type=hidden value='" & CurrencyCode & "'>"
	case 13
			 OrderName = " order by a.Val"
			 QuerySelect = 	"select a.RangeID, a.CreatedTime, a.CreatedDate, a.RangeID, a.Val, a.Expired from Ranges a"
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>C&oacute;digo</td><td class=titlelist><b>Rango</td><td class=titlelist><b>Status</td>"
			 Val = Request.Form("Val")
			 if Val <> "" then
			 		Option1 = " a.Val like '%" & Val & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Val' type=hidden value='" & Val & "'>"
	case 14
			 OrderName = " order by a.Tax"
			 QuerySelect = 	"select a.TaxID, a.CreatedTime, a.CreatedDate, a.TaxID, a.Tax, a.Expired from Taxes a"
			 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>C&oacute;digo</td><td class=titlelist><b>Impuesto</td><td class=titlelist><b>Status</td>"
			 Tax = Request.Form("Tax")
			 Option1 = " a.Countries in " & Session("Countries") & " "
			 if Tax <> "" then
			 		Option2 = " a.Tax like '%" & Tax & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Tax' type=hidden value='" & Tax & "'>"
	case 15, 18
			 'OrderName = " group by a.HAWBNumber order by a.CreatedDate Desc"
			 OrderName = " order by a.CreatedDate Desc"
 			 AwbType = CheckNum(Request.Form("AwbType"))
			 if GroupID=15 then
                 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>No. AWB</td><td class=titlelist><b>Embarcador</td><td class=titlelist><b>Status</td>"
			     if AwbType = 1 then
			 	    QuerySelect = "select a.AWBID, a.CreatedTime, a.CreatedDate, a.HAWBNumber, a.ShipperData, a.ArrivalDate, a.AWBNumber from Awb a"
			     else
			 	    QuerySelect = "select a.AWBID, a.CreatedTime, a.CreatedDate, a.HAWBNumber, a.ShipperData, a.ArrivalDate, a.AWBNumber from Awbi a"
			     end if
             else
                 HTMLTitle = "<td class=titlelist><b>Fecha</td><td class=titlelist><b>No. AWB</td><td class=titlelist><b>Consignatario</td><td class=titlelist><b>Status</td>"
                if AwbType = 1 then
			 	    QuerySelect = "select a.AWBID, a.CreatedTime, a.CreatedDate, a.HAWBNumber, a.ConsignerData, a.ArrivalDate, a.AWBNumber, ConsignerID from Awb a"
			     else
			 	    QuerySelect = "select a.AWBID, a.CreatedTime, a.CreatedDate, a.HAWBNumber, a.ConsignerData, a.ArrivalDate, a.AWBNumber, ConsignerID from Awbi a"
			     end if
             end if
			 Option1 = " a.Countries in " & Session("Countries") & " "'" and a.HAWBNumber<>'' "
			 HAWBNumber = Request.Form("AWBNumber")
			 CarrierID = CheckNum(Request.Form("CarrierID"))
			 AgentID = CheckNum(Request.Form("AgentID"))
			 AirportDepID = CheckNum(Request.Form("AirportDepID"))
			 AirportDesID = CheckNum(Request.Form("AirportDesID"))
			 
             if GroupID = 15 then
                 if HAWBNumber <> "" then
			 		    Option2 = " (a.AWBNumber like '%" & HAWBNumber & "%' or a.HAWBNumber like '%" & HAWBNumber & "%') "
			     end if
 			 else                
                'Option2 = " (a.HAWBNumber like '%" & HAWBNumber & "%' and a.HAWBNumber <>'') "  2015-06-03 hhmm
                Option2 = " (a.AWBNumber like '%" & HAWBNumber & "%' or a.HAWBNumber like '%" & HAWBNumber & "%') "
             end if

             if CarrierID <> 0 then
			 		Option3 = " a.CarrierID=" & CarrierID & " "
			 end if
 			 if AgentID <> 0 then
			 		Option4 = " a.AgentID=" & AgentID & " "
			 end if
 			 if AirportDepID <> 0 then
			 		Option5 = " a.AirportDepID=" & AirportDepID & " "
			 end if
 			 if AirportDesID <> 0 then
			 		Option6 = " a.AirportDesID=" & AirportDesID & " "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='AWBNumber' type=hidden value='" & HAWBNumber & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='CarrierID' type=hidden value='" & CarrierID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AgentID' type=hidden value='" & AgentID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AirportDepID' type=hidden value='" & AirportDepID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AirportDesID' type=hidden value='" & AirportDesID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AwbType' type=hidden value='" & AwbType & "'>"
	case 16
 			 AwbType = CheckNum(Request.Form("AwbType"))
			 ReportType = CheckNum(Request.Form("ReportType")) 'Agente o Transportista
             ResultType = CheckNum(Request.Form("ResultType")) 'Comisiones o Carga
			 

             select Case ResultType
             Case 0 'COMISIONES
             
                 select Case ReportType
			     Case 0 'X TRANSPORTISTA
				     OrderName = " group by month(a.CreatedDate)"
				     QuerySelect = "select month(a.CreatedDate), count(a.AWBID), sum(a.TotNoOfPieces*1), sum(a.TotWeightChargeable*1), " & _
								    "sum(a.TotChargeWeightPrepaid*1), sum(a.TotChargeWeightPrepaid*1*b.ComisionRate/100), sum(a.TotChargeValuePrepaid*1), sum(a.TotChargeTaxPrepaid*1), " & _
								    "sum(a.AnotherChargesAgentPrepaid*1), sum(a.AnotherChargesCarrierPrepaid*1), sum(a.TotPrepaid*1), " & _
								    "sum(a.TotChargeWeightCollect*1), sum(a.TotChargeWeightCollect*1*b.ComisionRate/100), sum(a.TotChargeValueCollect*1), sum(a.TotChargeTaxCollect*1), " & _
								    "sum(a.AnotherChargesAgentCollect*1), sum(a.AnotherChargesCarrierCollect*1), sum(a.TotCollect*1)"
				     if AwbType = 1 then
					    QuerySelect = QuerySelect & " from Awb a, Carriers b "
					    'QueryYearSelect = "select distinct year(a.CreatedDate) from Awb a "
					    HTMLTitle = "<tr><td class=titlelist3 align=center colspan=20>Resultados de Guias EXPORT X TRANSPORTISTA</td></tr>"
				     else
					    QuerySelect = QuerySelect & " from Awbi a, Carriers b "
					    'QueryYearSelect = "select distinct year(a.CreatedDate) from Awbi a "
					    HTMLTitle = "<tr><td class=titlelist3 align=center colspan=20>Resultados de Guias IMPORT X TRANSPORTISTA</td></tr>"
				     end if
				     Option1 = " a.CarrierID=b.CarrierID and a.Countries in " & Session("Countries") & " and a.HAWBNumber='' "
				     HTMLTitle = HTMLTitle & _
							    "<tr><td class=titlelist3 align=center colspan=4>&nbsp;</td><td class=titlelist align=center colspan=8><b>PREPAID</b></td><td class=titlelist2 align=center colspan=8><b>COLLECT</b></td></tr>" & _
							    "<tr><td class=titlelist3 align=center><b>Fecha</b></td><td class=titlelist3 align=center><b>Cant.<br>Guias M.</b></td>" & _
							    "<td class=titlelist3 align=center><b>Cant.<br>Bultos</b></td><td class=titlelist3 align=center><b>Peso</b></td>" & _
							    "<td class=titlelist align=center><b>Cargos<br>x Peso</b></td><td class=titlelist align=center><b>Comision</b>" & _
							    "</td><td class=titlelist align=center><b>Cargos<br>x Valor</b></td><td class=titlelist align=center><b>Impuestos</b></td>" & _
							    "<td class=titlelist align=center><b>Cargos Pagar<br>al Agente</b></td><td class=titlelist align=center><b>Cargos Pagar<br>al Transportista</b></td>" & _
							    "<td class=titlelist align=center><b>TOTAL</b></td><td class=titlelist align=center><b>PROFIT</b></td>" & _
							    "<td class=titlelist2 align=center><b>Cargos<br>x Peso</b></td><td class=titlelist2 align=center><b>Comision</b>" & _
							    "</td><td class=titlelist2 align=center><b>Cargos<br>x Valor</b></td><td class=titlelist2 align=center><b>Impuestos</b></td>" & _
							    "<td class=titlelist2 align=center><b>Cargos Pagar<br>al Agente</b></td><td class=titlelist2 align=center><b>Cargos Pagar<br>al Transportista</b></td>" & _
							    "<td class=titlelist2 align=center><b>TOTAL</b></td><td class=titlelist2 align=center><b>PROFIT</b></td></tr>"
			     case 1 'X AGENTE
				    'select TotCarrierRate, AWBID from Awbi where AWBNumber='832-88807084'
				    'select sum(TotCarrierRate), AWBNumber, 1 from Awbi where HAWBNumber<>'' and AWBNumber='832-88807084' group by AWBNumber union
				    'select TotCarrierRate, AWBNumber, 0 from Awbi where HAWBNumber='' and AWBNumber='832-88807084' order by AWBNumber
				    '2.204621
				 
				     if AwbType = 1 then
					    QuerySelect = " from Awb a" 'como utiliza SQL UNION comprende 2 Selects que se unifican mas abajo
					    'QueryYearSelect = "select distinct year(a.CreatedDate) from Awb a "
					    HTMLTitle = "<tr><td class=titlelist3 align=center colspan=20>Resultados de Guias EXPORT X AGENTE</td></tr>"
				     else
					    QuerySelect = " from Awbi a" 'como utiliza SQL UNION comprende 2 Selects que se unifican mas abajo
					    'QueryYearSelect = "select distinct year(a.CreatedDate) from Awbi a "
					    HTMLTitle = "<tr><td class=titlelist3 align=center colspan=20>Resultados de Guias IMPORT X AGENTE</td></tr>"
				     end if
				     Option1 = " a.Countries in " & Session("Countries") & " "
				      'Option1 = " a.Countries = 'GT' "
				     HTMLTitle = HTMLTitle & _
				 			    "<tr><td class=titlelist3 align=center colspan=3>&nbsp;</td><td class=titlelist3 align=center colspan=3><b>House Revenue</b></td>" & _
							    "<td class=titlelist align=center colspan=3><b>Master Expense</b></td><td class=titlelist2 align=center colspan=5>&nbsp;</td></tr>" & _
							    "<tr><td class=titlelist3 align=center><b>Fecha</b></td><td class=titlelist3 align=center><b>Viaje</b></td>" & _
							    "<td class=titlelist3 align=center><b>M.AWB</b></td><td class=titlelist3 align=center><b>Air+Fuel+Sec</b></td>" & _
							    "<td class=titlelist3 align=center><b>Revenue<br>Intermodal</b></td><td class=titlelist3 align=center><b>Pick-Up</b></td>" & _
							    "<td class=titlelist align=center><b>Air+Fuel+Sec</b></td><td class=titlelist align=center><b>Intermodal</b></td>" & _
							    "<td class=titlelist align=center><b>Pick-Up</b></td><td class=titlelist2 align=center><b>Ch.Weight<br>(Kgs.)</b></td>" & _
							    "<td class=titlelist2 align=center><b>Ch.Weight<br>(Lbs.)</b></td>" & _
							    "<td class=titlelist2 align=center><b>ECONO<br>Admin Fee</b></td><td class=titlelist2 align=center><b>PROFIT<br>neto</b></td>" & _
							    "<td class=titlelist2 align=center><b>Commission<br>50%</b></td></tr>"
			     end select
             
             Case 1 'CARGA
             
                if AwbType = 1 then
                    AwbTableName = "Awb"
                else
                    AwbTableName = "Awbi"
                end if
                
                OrderName = " order by a.CreatedDate Desc"
 			    HTMLTitle = "<td class=titlelist><b>Agente</b></td><td class=titlelist><b>Origen</b></td><td class=titlelist><b>HAWB</b></td><td class=titlelist><b>KGS</b></td><td class=titlelist><b>Aerolinea</b></td><td class=titlelist><b>Fecha Salida</b></td><td class=titlelist><b>Fecha Llegada</b></td><td class=titlelist><b>Dias Transito</b></td><td class=titlelist><b>Shipper</b></td><td class=titlelist><b>Consignatario</b></td><td class=titlelist><b>Routing</b></td><td class=titlelist><b>Ruteada por</b></td>"
                QuerySelect = "select a.AgentData, b.Name, a.HAWBNumber, a.Weights, c.Name, a.HDepartureDate, a.ArrivalDate, a.ShipperData, a.ConsignerData, a.Routing, a.RoutingID" & _
                    " from " & AwbTableName & " a, Airports b, Carriers c"
                Option1 = " a.AirportDepID=b.AirportID and a.CarrierID=c.CarrierID" & _
                    " and a.HAWBNumber<>'' and a.Countries in " & Session("Countries")
             

             Case 2 'MEDICIONES 2016-02-12
                                
                AwbType = CheckNum(Request.Form("AwbType"))

                dim ExpImp, UserStr

                if AwbType = 1 then
                    ExpImp = "EXPORTACION"
                else
                    ExpImp = "IMPORTACION"
                end if

                QuerySelect = "SELECT * FROM mediciones a "
                Option1 = "" '" a.UserInsert = " & Session("OperatorID") & " AND Status = 1 "

                
                Dim Conn, SQLQuery, rs
                openConn Conn 'aereo
                
                Dim Segmentos, LabelUnoMain, LabelUnoDate, LabelUnoTime, LabelDosMain, LabelDosDate, LabelDosTime, LabelTresMain, LabelTresDate, LabelTresTime, LabelCuatroMain, LabelCuatroDate, LabelCuatroTime, LabelCincoMain, LabelCincoDate, LabelCincoTime, LabelUnoMedicion, LabelDosMedicion, cols

                'si pais no tiene registro de etiquetas, toma las default que son guatemala
                SQLQuery = "SELECT * FROM mediciones_labels WHERE countries IN  ('" & Request.Form("Countries") & "','GT') AND AwbType = " & AwbType & " ORDER BY id DESC LIMIT 1"
                set rs = Conn.Execute(SQLQuery)
                'response.write SQLQuery
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


                cols = 4

                If InStr(1,Segmentos, "1") > 0 Then    
                    cols = cols + 1
                End If

                If InStr(1,Segmentos, "2") > 0 Then
                    cols = cols + 1
                End If

                If InStr(1,Segmentos, "3") > 0 Then
                    cols = cols + 1
                End If

                If InStr(1,Segmentos, "4") > 0 Then
                    cols = cols + 1
                End If

                If InStr(1,Segmentos, "5") > 0 Then
                    cols = cols + 1
                End If


                HTMLTitle = "<thead>" & _ 
                "<tr><th><input type=button onclick=ReporteExcel() value='EXCEL'><th colspan=" & (cols + 2) & "><h2>FORMATO DE MEDICION DE " & ExpImp & "</h2></th><th> EDICION 2</tr>" & _ 
                "<tr><th colspan=2><h4> Usuario : " & Session("OperatorName") & "<th colspan=" & cols & "><h2>" & FECHAS & "<th colspan=2><h4>" & date & " " & time & "</h2></th></tr>" & _ 
                "<tr><th>Mawb'S:</th><th>Hawb'S</th><th>Bultos</th><th>Peso</th><th>Destino</th><th>Shipper'S</th>"

                If InStr(1,Segmentos, "1") > 0 Then    
                    HTMLTitle = HTMLTitle & "<th>" & LabelUnoMain & "</th>"
                End If

                If InStr(1,Segmentos, "2") > 0 Then
                    HTMLTitle = HTMLTitle & "<th>" & LabelDosMain & "</th>"
                End If

                If InStr(1,Segmentos, "3") > 0 Then
                    HTMLTitle = HTMLTitle & "<th>" & LabelTresMain & "</th>"
                End If

                If InStr(1,Segmentos, "4") > 0 Then
                    HTMLTitle = HTMLTitle & "<th>" & LabelCuatroMain & "</th>"
                End If

                If InStr(1,Segmentos, "5") > 0 Then
                    HTMLTitle = HTMLTitle & "<th>" & LabelCincoMain & "</th>"
                End If

                HTMLTitle = HTMLTitle & "<th>" & LabelUnoMedicion & "</th>"
                HTMLTitle = HTMLTitle & "<th>" & LabelDosMedicion & "</th>"
                HTMLTitle = HTMLTitle & "</tr></thead>"
                

                
                
                'MedicionID, CreatedDate, CreatedTime, AwbID, AwbNumber, HAwbNumber, AwbType, DateUno, TimeUno, DateDos, TimeDos, DateTres, TimeTres, DateCuatro, TimeCuatro, MedicionUno, MedicionDos, TotNoOfPieces, TotWeight, Destinity, ShipperData, UserInsert, UserUpdate, DateUpdate, Status 
                 
			     HTMLHidden = HTMLHidden & "<INPUT name='MMFrom' type=hidden value='" & MMFrom & "'>"
			     HTMLHidden = HTMLHidden & "<INPUT name='YYFrom' type=hidden value='" & YYFrom & "'>"
                 HTMLHidden = HTMLHidden & "<INPUT name='MMTo' type=hidden value='" & MMTo & "'>"
			     HTMLHidden = HTMLHidden & "<INPUT name='YYTo' type=hidden value='" & YYTo & "'>"
			     HTMLHidden = HTMLHidden & "<INPUT name='ResultType' type=hidden value='" & ResultType & "'>"
			     HTMLHidden = HTMLHidden & "<INPUT name='MMFromText' type=hidden value='" & Request.Form("MMFromText") & "'>"
                 HTMLHidden = HTMLHidden & "<INPUT name='MMToText' type=hidden value='" & Request.Form("MMToText") & "'>"
                 HTMLHidden = HTMLHidden & "<INPUT name='AwbType' type=hidden value='" & AwbType & "'>"

            case 3

                QuerySelect = "SELECT distinct a.AWBID, a.CreatedDate, b.CreatedTime, a.AwbNumber, a.HAwbNumber, a.UserID, b.DocTyp, b.CurrencyID, b.ItemID, b.ItemName, b.Value, b.PrepaidCollect, b.ServiceID, b.ServiceName, b.InvoiceID, b.ChargeID, b.DocType, c.FirstName, c.LastName FROM "

			     if AwbType = 1 then
			 	    QuerySelect = QuerySelect & "Awb a"
			     else
			 	    QuerySelect = QuerySelect & "Awbi a"
			     end if

                 QuerySelect = QuerySelect & " INNER JOIN ChargeItems b ON a.AwbID = b.AWBID AND b.Expired = '0' AND b.DocType IN (1,4) AND b.DocTyp = '0' AND b.InvoiceID > 0 LEFT JOIN Operators c ON c.OperatorID = a.UserID "
                 OrderName = " ORDER BY a.AWBID, b.ItemID"    
                Option1 = " a.Countries = '" & trim(Request.Form("Countries")) & "' "			     
                MoreOptions = ""


                HTMLTitle = HTMLTitle & "<thead>" 
                
                HTMLTitle = HTMLTitle & "<tr>" 

                if Request.Form("excel") = 1 then 
                HTMLTitle = HTMLTitle & "<th colspan=16>BITACORA DE FACTURACION " & trim(Request.Form("Countries")) & "</th>" 
                else
                HTMLTitle = HTMLTitle & "<th colspan=16>BITACORA DE FACTURACION " & trim(Request.Form("Countries")) & " <input type=button onclick=ReporteExcel() value='EXCEL'></th>"                 
                end if

                HTMLTitle = HTMLTitle & "</tr>" 

                HTMLTitle = HTMLTitle & "<tr>" 
                HTMLTitle = HTMLTitle & "<th></th>" 
                HTMLTitle = HTMLTitle & "<th>Fecha</th>" 
                HTMLTitle = HTMLTitle & "<th>Hora</th>" 
                HTMLTitle = HTMLTitle & "<th>Awb</th>" 
                HTMLTitle = HTMLTitle & "<th>Hawb</th>"
                HTMLTitle = HTMLTitle & "<th>Imp / Exp</th>"
        
                HTMLTitle = HTMLTitle & "<th>PreCol</th>"
                HTMLTitle = HTMLTitle & "<th>Usuario</th>"
                    
                HTMLTitle = HTMLTitle & "<th>Item</th>"
                HTMLTitle = HTMLTitle & "<th>Valor GTQ</th>"
                HTMLTitle = HTMLTitle & "<th>Valor USD</th>"
      
                HTMLTitle = HTMLTitle & "<th>Documento</th>" 
                HTMLTitle = HTMLTitle & "<th>SubTotal</th>" 
                HTMLTitle = HTMLTitle & "<th>Iva</th>" 
                HTMLTitle = HTMLTitle & "<th>Total</th>" 
                HTMLTitle = HTMLTitle & "</tr>" 
                HTMLTitle = HTMLTitle & "</thead>"

                 
			     HTMLHidden = HTMLHidden & "<INPUT name='MMFrom' type=hidden value='" & MMFrom & "'>"
			     HTMLHidden = HTMLHidden & "<INPUT name='YYFrom' type=hidden value='" & YYFrom & "'>"
                 HTMLHidden = HTMLHidden & "<INPUT name='MMTo' type=hidden value='" & MMTo & "'>"
			     HTMLHidden = HTMLHidden & "<INPUT name='YYTo' type=hidden value='" & YYTo & "'>"
			     HTMLHidden = HTMLHidden & "<INPUT name='ResultType' type=hidden value='" & ResultType & "'>"
			     HTMLHidden = HTMLHidden & "<INPUT name='MMFromText' type=hidden value='" & Request.Form("MMFromText") & "'>"
                 HTMLHidden = HTMLHidden & "<INPUT name='MMToText' type=hidden value='" & Request.Form("MMToText") & "'>"
                 'HTMLHidden = HTMLHidden & "<INPUT name='AwbType' type=hidden value='" & AwbType & "'>"
                 HTMLHidden = HTMLHidden & "<INPUT name='Countries' type=hidden value='" & Request.Form("Countries") & "'>"

             end select

			 
             AWBNumber = Request.Form("AWBNumber")
			 CarrierID = CheckNum(Request.Form("CarrierID"))
			 AgentID = CheckNum(Request.Form("AgentID"))
			 SalespersonID = CheckNum(Request.Form("SalespersonID"))
			 ShipperID = CheckNum(Request.Form("ShipperID"))
			 AirportDepID = CheckNum(Request.Form("AirportDepID"))
			 AirportDesID = CheckNum(Request.Form("AirportDesID"))


			 if AWBNumber <> "" then
			 		Option2 = " a.AWBNumber like '%" & AWBNumber & "%' "					
			 end if
 			 if CarrierID <> 0 then
			 		Option3 = " a.CarrierID=" & CarrierID & " "
			 end if
 			 if AgentID <> 0 then
			 		Option4 = " a.AgentID=" & AgentID & " "
			 end if
 			 if SalespersonID <> 0 then
			 		Option5 = " a.SalespersonID=" & SalespersonID & " "
			 end if
			 if ShipperID <> 0 then
			 		Option6 = " a.ShipperID=" & ShipperID & " "
			 end if
 			 if AirportDepID <> 0 then
			 		Option7 = " a.AirportDepID=" & AirportDepID & " "
			 end if
 			 if AirportDesID <> 0 then
			 		Option8 = " a.AirportDesID=" & AirportDesID & " "
			 end if

             if YYFrom <> 0 and MMFrom <> 0 then
			 		Option9 = " Year(a.CreatedDate)>=" & YYFrom & " and Month(a.CreatedDate)>=" & MMFrom & " "
			 end if

             if YYTo <> 0 and MMTo <> 0 then
			 		Option10 = " Year(a.CreatedDate)<=" & YYTo & " and Month(a.CreatedDate)<=" & MMTo & " "
			 end if

			 HTMLHidden = HTMLHidden & "<INPUT name='AWBNumber' type=hidden value='" & HAWBNumber & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='CarrierID' type=hidden value='" & CarrierID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AgentID' type=hidden value='" & AgentID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='SalespersonID' type=hidden value='" & SalespersonID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='ShipperID' type=hidden value='" & ShipperID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AirportDepID' type=hidden value='" & AirportDepID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AirportDesID' type=hidden value='" & AirportDesID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AwbType' type=hidden value='" & AwbType & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='MM' type=hidden value=''>"
			 HTMLHidden = HTMLHidden & "<INPUT name='YY' type=hidden value=''>"
             HTMLHidden = HTMLHidden & "<INPUT name='excel' type=hidden value=''>"

	end select

	'Construyendo el Query segun los parametros de busqueda seleccionados en la pagina anterior
	DateFrom = Request.Form("DateFrom")
	DateTo = Request.Form("DateTo")
	
	if isDate(DateFrom) then
		 Option9 = " a.CreatedDate>='" & ConvertDate(DateFrom,4) & "' "
	end if	
	if isDate(DateTo) then
		 Option10 = " a.CreatedDate<='" & ConvertDate(DateTo,4) & "' "
	end if
	HTMLHidden = HTMLHidden & "<INPUT name=DateFrom type=hidden value='" & DateFrom & "'>"
	HTMLHidden = HTMLHidden & "<INPUT name=DateTo type=hidden value='" & DateTo & "'>"
	MoreOptions = 0
	CreateSearchQuery QuerySelect, Option1, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option2, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option3, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option4, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option5, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option6, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option7, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option8, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option9, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option10, MoreOptions, " and "

	HTMLCode = ""
	'response.write QuerySelect & "<br>" & ResultType    
    response.write "<script>console.log('" & Replace(QuerySelect,"'","") & "');</script>" 

    if GroupID <> 16 then
 		DisplaySearchAdminResults HTMLCode
	else

        select Case ResultType
		Case 0 
            DisplayStats
		case 1 
            DisplayCargoStats
		case 2 
            DisplayMediciones(Segmentos)        
		case 3 
            DisplayBitacora
		end select
             
	end if
%>

<HTML><HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function NextPage(PageNo) {
	document.forma.P.value = PageNo;
	document.forma.submit();
}

function DisplayDetailStats (MM, YY) {
	document.forma.MM.value = MM;
	document.forma.YY.value = YY;
	document.forma.submit();
}

function ReporteExcel() {
    var action_tmp = document.forma.action;
    var target_tmp = document.forma.target;

    document.forma.action = "Search_ResultsAdmin.asp";
    document.forma.target = "_blank";
    document.forma.excel.value = 1;
    document.forma.submit();

    document.forma.action = action_tmp;
    document.forma.target = target_tmp;
}



</SCRIPT>
<%if ResultType = 2 or ResultType = 3 then%>				
        <LINK REL="stylesheet" type="text/css" HREF="img/2016.css">
<% else %>
        <LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<% end if %>



<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<%if JavaMsg <> "" then
			 Response.Write "<SCRIPT>alert('" & JavaMsg & "');</SCRIPT>"
		end if
	%>
	<FORM name="forma" action=<%if GroupID <>16 then%>"Search_ResultsAdmin.asp"<%else%>"Rep_DetailStats.asp"<%end if%> method="post" target=_self>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
  	<INPUT name="Action" type=hidden value=1>
	<INPUT name="P" type=hidden value=1>
  	<%=HTMLHidden%>
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<TR>
		<TD width=40% colspan=2 class=label align=right valign=top>

				<%if GroupID<>16 then%>				
					<TABLE cellspacing=3 cellpadding=2 width=100% align=center>
				<%else%>

                    <%if ResultType = 2 or ResultType = 3 then%>						

                        <% if Request.Form("excel") = 1 then 'solo en este reporte exporta a excel
                            Response.ContentType = "application/vnd.ms-excel"
                            Response.AddHeader "content-disposition", " filename=excelTest.xls"
                            response.write("<LINK REL='stylesheet' type='text/css' HREF='img/2016.css'>")
                            response.write ("<TABLE width=100% border=1>")
                        else
                            response.write ("<TABLE width=100% class='GridView'>")
                        end if %>

					    
                    <%else%>
                        <TABLE cellspacing=1 cellpadding=3 align=center border=1>
                    <%end if%>

				<%end if%>
					 <%=HTMLTitle%>
					 <%=HTMLCode%>
				</TABLE>


		</TD>
	  </TR>

<% if Request.Form("excel") = "" then %>

<% if PageCount > 1 then%>
		<TR>
		<TD width=40% colspan=2 class=label align=right valign=top>
				<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
				<TR>
				<TD class=label align=left valign="top" width=15%>
				<%if AbsolutePage > 1 then%>&nbsp;
								<a class=label onclick=JavaScript:NextPage("<%=(AbsolutePage-1)%>"); href=# target=_self><u><< Anterior</u></a>&nbsp;
				<%else%>
								<a class=label href="Search_admin.asp?GID=<%=GroupID%>" target=_self><u><< Regresar</u></a>&nbsp;

                                

				<%end if%>&nbsp;
				</TD>
				<TD class=label align=center>
							 <%
							 for i = 1 to PageCount
							 		 Response.write "&nbsp;<a class=label onclick=JavaScript:NextPage(" & i & ") href=#><u>" & i & "</u></a>&nbsp;"
							 		 if i <> PageCount then
							 		 		Response.write "<font class=label>|</font>" 
							 		 end if
									 if (i mod 20) = 0 then
									 		Response.write "<br>"
									 end if
							 next
							 %>
				</TD>
				<TD class=label align=right valign="top" width=15%>&nbsp;
				<%if PageCount <> AbsolutePage then%> 
						 <a class=label onclick=JavaScript:NextPage("<%=(AbsolutePage+1)%>"); href=# target=_self><u>Siguiente >></u></a>
				<%end if%>&nbsp;
				</TD>
				</TR>
				</TABLE>
		</TD>
	  </TR>
<%else%>
		<TR>
		<TD width=40% colspan=2 class=label align=right valign=top>
				<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
				<TR>
				<TD class=label align=left>
				<a class=label onclick=JavaScript:history.back(); href=# target=_self><u><< Regresar</u></a>
				</TD>
				</TR>
				</TABLE>
		</TD>
	  </TR>
<%end if ' PageCount %>		

<%end if 'excel %>		
	</TABLE>
  </FORM>				
</BODY>
</HTML>
<%
end if
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>

