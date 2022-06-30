<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"

Dim ObjectID, esquema, AwbType, Conn, rs, Action, QuerySelect, QuerySelect2, QuerySelect3, i, Asigns, Ides, cagen, ctrans, cotros, j, Pos, k, l, Res, Homologado, ValidoHomo, AgentTypeItemID

ObjectID = CheckNum(Request("OID"))
Action = Request("Action")
AwbType = Request("AT")
Pos = Request("Pos")
esquema = Request("esquema")

if Action <> "" then

    OpenConn Conn

    Response.Write "ENTRO (" & Action & ")(" & ObjectID & ")(" & esquema & ")<br>"

%>

<body onload="refresh();">

</body>

<script type="text/javascript">
    function refresh() {
        parent.document.forma.Action.value = 0;
        parent.document.forma.target = "";
        parent.document.forma.action = "AwbCharges.asp";
        
        //parent.location.reload(true);

        window.parent.document.location.reload(true);

    }
</script>
<%

    if Action = "asignar" or Action = "liberar" then 'asignar clientes

        Response.Write "(" & Request.Form("CHK") & ")<br>"

        Asigns=split(Request.Form("CHK"),",") 'siempre viene solo un registro NO puede traer los que el usuario seleccione

        For i=LBound(Asigns) to UBound(Asigns)    
                
            Ides=split(Asigns(i),"|") '0 ChargeID, 1 id_pedido, 2 pedido_erp, 5 cantidad rubros con pedido_id 

            select case Action 
                case "998", "994" 'delete / update

                    QuerySelect = "UPDATE ChargeItems SET Expired = 1 WHERE ChargeID = '" & Ides(0) & "'"
                    Response.Write QuerySelect & "<br>"
                    Conn.Execute(QuerySelect)
                
                    SaveChargeBL Conn, ObjectID

                case "liberar" 
                
                    'id_cliente = NULL, cliente_nombre = NULL,

                    QuerySelect = "UPDATE ChargeItems SET id_pedido = NULL, pedido_erp = NULL, InvoiceID = 0, DocType = 0 WHERE ChargeID = '" & Ides(0) & "'"
                    Response.Write QuerySelect & "<br>"
                    Conn.Execute(QuerySelect)
                
                
                case "asignar" 'solo cuando asigna cliente
   
                    Res = ValidaHomologacion("1", esquema, "01", "'" & Ides(3) & "'")   

                    On Error Resume Next
                        Homologado = IFNULL(Res(0,0))
                    If Err.Number <> 0 Then
                        Homologado = IFNULL(Res)
                    end if

                    'Response.Write "(Homologado=" & Homologado & ")<br>"

                    if Homologado <> "" then
                    
                        j = UBound(Res)

                        'Response.Write "(j=" & j & ")<br>"
                        
                        'if j > 0 then

                            'Response.Write "(" & Res(0,0) & ")<br>"
                            'Response.Write "(" & Res(1,0) & ")<br>"
                            'Response.Write "(" & Res(2,0) & ")<br>"
                            'Response.Write "(" & Res(3,0) & ")<br>"
    
                        'end if

                        QuerySelect = "UPDATE ChargeItems SET id_cliente = " & Ides(3) & ", cliente_nombre = '" & Ides(4) & "', id_pedido = NULL, pedido_erp = NULL, InvoiceID = 0, DocType = 0 WHERE ChargeID = '" & Ides(0) & "'"
                        Response.Write QuerySelect & "<br>"
                        Conn.Execute(QuerySelect)


                    else

                        response.write "<script" & ">alert('Cliente seleccionado no esta homologado');</script>"
    
                    end if
                    
            end select 

        Next 

        CloseOBJ Conn

        Response.Write "Finalizo Proceso<br>"

        Response.End

    end if




    '/////////////////////////////////////////////////// SECCION RUBROS 

    Dim CantItems, CreatedDate, CreatedTime, ItemCurrs, ItemIDs, ItemVals, ItemLocs, ItemNames, ItemOVals, ItemPPCCs, ItemServIDs, ItemChargeID, ItemPedErp, ItemServNames, ItemInvoices, ItemCalcInBls, ItemInRO, CType, ItemInterCompanyIDs, ItemDocType, ItemCli, ItemCliNom, ItemRegimen, ItemTarifaPrice, ItemTarifaTipo, ItemAgent

	FormatTime CreatedDate, CreatedTime
	
	CantItems = CheckNum(Request.Form("CantItems"))
	ItemCurrs = Split(Request.Form("ItemCurrs"), "|")
	ItemIDs = Split(Request.Form("ItemIDs"), "|")
	ItemServIDs = Split(Request.Form("ItemServIDs"), "|")
	ItemServNames = Split(Request.Form("ItemServNames"), "|")
	ItemVals = Split(Request.Form("ItemVals"), "|")
	ItemLocs = Split(Request.Form("ItemLocs"), "|")
	ItemNames = Split(Request.Form("ItemNames"), "|")
	ItemOVals = Split(Request.Form("ItemOVals"), "|")
	ItemPPCCs = Split(Request.Form("ItemPPCCs"), "|")
	ItemInvoices = Split(Request.Form("ItemInvoices"), "|")
	ItemDocType = Split(Request.Form("ItemDocType"), "|")
	ItemCalcInBLs = Split(Request.Form("ItemCalcInBls"), "|")
	ItemInRO = Split(Request.Form("ItemInRO"), "|")
	ItemInterCompanyIDs = Split(Request.Form("ItemInterCompanyIDs"), "|")                
    ItemChargeID = Split(Request.Form("ItemChargeID"), "|")
    ItemAgent = Split(Request.Form("ItemAgent"), "|")

    if Request.Form("ItemCli") = "" then
        ItemCli = Split("0", "|")    
    else
        ItemCli = Split(Request.Form("ItemCli"), "|")
    end if

    if Request.Form("ItemPedErp") = "" then
        ItemPedErp = Split(" ", "|")    
    else
        ItemPedErp = Split(Request.Form("ItemPedErp"), "|")
    end if

    if Request.Form("ItemCliNom") = "" then
        ItemCliNom = Split(" ", "|")    
    else    
        ItemCliNom = Split(Request.Form("ItemCliNom"), "|")
    end if

    if Request.Form("ItemRegimen") = "" then
        ItemRegimen = Split(" ", "|")    
    else    
        ItemRegimen = Split(Request.Form("ItemRegimen"), "|")
    end if

    if Request.Form("ItemTarifaPrice") = "" then
        ItemTarifaPrice = Split(" ", "|")    
    else    
        ItemTarifaPrice = Split(Request.Form("ItemTarifaPrice"), "|")
    end if

    if Request.Form("ItemTarifaTipo") = "" then
        ItemTarifaTipo = Split(" ", "|")    
    else    
        ItemTarifaTipo = Split(Request.Form("ItemTarifaTipo"), "|")
    end if

    select case Action 
        
        case "insert", "update", "borrar"

            cagen = -1
            ctrans = -1
            cotros = -1

            ValidoHomo = true
            QuerySelect3 = ""
            QuerySelect = ""

            '///////////////////////////////////////// RESET TODOS LOS VALORES

            QuerySelect = QuerySelect & ", CustomFee = ''"
            QuerySelect = QuerySelect & ", TerminalFee = ''"
            QuerySelect = QuerySelect & ", TotCarrierRate = '', TotCarrierRate_Routing = 0"
            QuerySelect = QuerySelect & ", FuelSurcharge = '', FuelSurcharge_Routing = 0"
            QuerySelect = QuerySelect & ", SecurityFee = '', SecurityFee_Routing = 0"
            QuerySelect = QuerySelect & ", PickUp = '', PickUp_Routing = 0"
            QuerySelect = QuerySelect & ", SedFilingFee = '', SedFilingFee_Routing = 0"
            QuerySelect = QuerySelect & ", Intermodal = '', Intermodal_Routing = 0"
            QuerySelect = QuerySelect & ", PBA = '', PBA_Routing = 0"


            if AwbType = 1 then

                QuerySelect = QuerySelect & ", CustomFee_Routing = 0"
                QuerySelect = QuerySelect & ", TerminalFee_Routing = 0"

            end if

            for j=1 to 15
                QuerySelect = QuerySelect & ", AdditionalChargeName" & j & " = '', AdditionalChargeVal" & j & " = '', AdditionalChargeName" & j & "_Routing = 0"
            next

            if AWBType <> 1 then 'import
                for j=1 to 6
                    QuerySelect = QuerySelect & ", OtherChargeName" & j & " = '', OtherChargeVal" & j & " = '', OtherChargeName" & j & "_Routing = 0"
                next
            end if

            '/////////////////////////////////////////////////////////////////este query se ejecuta hasta de ultimo 
            QuerySelect3 = "UPDATE " & Iif(AWBType = "1","Awb","Awbi") & " SET Expired = 0" & QuerySelect & " WHERE AwbID = " & ObjectID


            QuerySelect = ""

            for i=0 to CantItems

                if Action = "borrar" and Cdbl(Pos) = Cdbl(i) then
                    'cuando borra no acumula el rubro a la guia

                else

                    k = 0
                    l = 0

                    AgentTypeItemID = CheckNum(ItemAgent(i)) & "-" & CheckNum(ItemIDs(i))

                    select case AgentTypeItemID
			            case "0-14"	'CustomFee
                            QuerySelect = QuerySelect & ", CustomFee = " & CheckNum(ItemVals(i)) & ""
                            'ctrans = ctrans + 1
			            case "0-15"	'TerminalFee
                            QuerySelect = QuerySelect & ", TerminalFee = " & CheckNum(ItemVals(i)) & ""
                            'ctrans = ctrans + 1
			            case "0-11"	'TotCarrierRate
                            QuerySelect = QuerySelect & ", TotCarrierRate = " & CheckNum(ItemVals(i)) & ", TotCarrierRate_Routing = 0"
                            'ctrans = ctrans + 1
			            case "0-12"	'FuelSurcharge
                            QuerySelect = QuerySelect & ", FuelSurcharge = " & CheckNum(ItemVals(i)) & ", FuelSurcharge_Routing = 0"
                            'ctrans = ctrans + 1
			            case "0-13"	'SecurityFee
                            QuerySelect = QuerySelect & ", SecurityFee = " & CheckNum(ItemVals(i)) & ", SecurityFee_Routing = 0"
                            'ctrans = ctrans + 1
			            case "0-31"	'PickUp
                            QuerySelect = QuerySelect & ", PickUp = " & CheckNum(ItemVals(i)) & ", PickUp_Routing = 0"
                            'ctrans = ctrans + 1
			            case "0-38"	'SedFilingFee
                            QuerySelect = QuerySelect & ", SedFilingFee = " & CheckNum(ItemVals(i)) & ", SedFilingFee_Routing = 0"
                            'ctrans = ctrans + 1
			            case "0-115"    'Intermodal
                            QuerySelect = QuerySelect & ", Intermodal = " & CheckNum(ItemVals(i)) & ", Intermodal_Routing = 0"
                            'ctrans = ctrans + 1
			            case "0-116"	'PBA	
                            QuerySelect = QuerySelect & ", PBA = " & CheckNum(ItemVals(i)) & ", PBA_Routing = 0"
                            'ctrans = ctrans + 1
                        case else

                            select case CheckNum(ItemAgent(i))

                                case 0 'cargos transportista

                                    ctrans = ctrans + 1

                                    select case ctrans
                                    case 0 j = 3
                                    case 1 j = 4
                                    case 2 j = 5
                                    case 3 j = 8
                                    end select

                                    'k = j           'casilla en guia
                                    l = ctrans+1    'pos en rubros

    'response.write "(" & ItemNames(i) & ")(" & ctrans & ")(" & l & ")(" & j & ")<br>"

                                    QuerySelect = QuerySelect & ", AdditionalChargeName" & j & " = '" & ItemNames(i) & "', AdditionalChargeVal" & j & " = " & CheckNum(ItemVals(i)) & ", AdditionalChargeName" & j & "_Routing = 0"


                                case 1 'cargos agente

                                    cagen = cagen + 1

                                    select case cagen
                                    case 0 j = 1
                                    case 1 j = 2
                                    case 2 j = 6
                                    case 3 j = 7
                                    case 4 j = 9
                                    case 5 j = 10
                                    case 6 j = 11
                                    case 7 j = 12
                                    case 8 j = 13
                                    case 9 j = 14
                                    case 10 j = 15        
                                    end select	

                                    'k = j           'casilla en guia
                                    l = cagen+1     'pos en rubros

                                    QuerySelect = QuerySelect & ", AdditionalChargeName" & j & " = '" & ItemNames(i) & "', AdditionalChargeVal" & j & " = " & CheckNum(ItemVals(i)) & ", AdditionalChargeName" & j & "_Routing = 0"
                    
                                case 2 'otros import

                                    cotros = cotros + 1
                                    j = cotros+1

                                    'k = j           'casilla en guia
                                    l = cotros+1    'pos en rubros
                                    QuerySelect = QuerySelect & ", OtherChargeName" & j & " = '" & ItemNames(i) & "', OtherChargeVal" & j & " = " & CheckNum(ItemVals(i)) & ", OtherChargeName" & j & "_Routing = 0"

                            end select

			        end select


                end if

                'Homologado = "*" 'debe llevar algo

                'Response.Write "ItemAgent=" & ItemAgent(i) & " Agen=" & cagen & " Tran=" & ctrans & " Otr=" & cotros & "<br>"

                Response.Write "Pos=" & Pos & " i=" & i & " ItemAgent=" & CheckNum(ItemAgent(i)) & " l=" & l & "<br>"


                if Cdbl(Pos) = Cdbl(i) then


                    if Action = "insert" then                        

                        Res = ValidaHomologacion("1", esquema, "01", "'" &  CheckNum(ItemCli(i)) & "'")   

                        On Error Resume Next
                            Homologado = IFNULL(Res(0,0)) 'si trae valor 
                            ValidoHomo = true
                        If Err.Number <> 0 Then
                            Homologado = IFNULL(Res) 'aca asigna blancos
                            ValidoHomo = false
                        end if

                    end if

                    Response.Write "(Homologado=" & ValidoHomo & ")<br>"

                    if Action = "update" or Action = "borrar" then                        
                        'QuerySelect2 = "UPDATE ChargeItems SET Expired=1 WHERE ChargeID = " & ItemChargeID(i) 

                        QuerySelect2 = "UPDATE ChargeItems SET Expired=1 WHERE AwbID = " & ObjectID & " AND ItemID = " & CheckNum(ItemIDs(i))
                        Response.Write QuerySelect2 & "<br>"
                        Conn.Execute(QuerySelect2)
                    end if

        'Response.Write "Pos=" & Pos & " i=" & i & "<br>"
                
        'Response.Write "1*ENTRO (" & ItemCurrs(i) & ")(" & ItemIDs(i) & ")(" & ItemVals(i) & ")(" & ItemLocs(i) & ")(" & ItemAgent(i) & ")<br>"

        'Response.Write "2*ENTRO (" & ItemNames(i) & ")(" & ItemServIDs(i) & ")(" & ItemServNames(i) & ")(" & ItemPPCCs(i) & ")(" & ItemTarifaPrice(i) & ")(" & ItemRegimen(i) & ")(" & ItemTarifaTipo(i) & ")<br>"

        'Response.Write "3*ENTRO (" & ItemCli(i) & ")(" & ItemPedErp(i) & ")(" & ItemCliNom(i) & ")<br>"

                    if Action = "update" or (Action = "insert" and ValidoHomo = true) then                        
                
                        QuerySelect2 = "INSERT INTO ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, UserID, Pos, ServiceID, ServiceName, PrepaidCollect, CalcInBL, ItemName_Routing, TarifaPricing, Regimen, TarifaTipo, id_cliente, pedido_erp, cliente_nombre) VALUES " & _                    
                        "(" & ObjectID  & ", '" & ItemCurrs(i) & "', " & CheckNum(ItemIDs(i)) & ", " & CheckNum(ItemVals(i))  & ", " & CheckNum(ItemLocs(i))  & ", " & CheckNum(ItemAgent(i)) & ", " & IIf(AWBType = 1,"0","1") & ", " & _ 
                        "'" & ItemNames(i) & "', '" & CreatedDate & "', " & CreatedTime & ", " & Session("OperatorID") & ", " & CheckNum(l) & ", " & CheckNum(ItemServIDs(i)) & ", '" & ItemServNames(i) & "', " & CheckNum(ItemPPCCs(i))  & ", 0, 0, '" & CheckNum(ItemTarifaPrice(i)) & "', '" & ItemRegimen(i) & "', '" & ItemTarifaTipo(i) & "', " & CheckNum(ItemCli(i)) & ", '" & Trim(ItemPedErp(i)) & "', '" & Left(Trim(ItemCliNom(i)),99) & "')"    				
                        Response.Write QuerySelect2 & "<br>"
                        Conn.Execute(QuerySelect2)

                    end if


                else

                    if Action = "borrar" then                        

                        QuerySelect2 = "UPDATE ChargeItems SET Pos=" & CheckNum(l) & " WHERE ChargeID = " & ItemChargeID(i) 
                        Response.Write QuerySelect2 & "<br>"
                        Conn.Execute(QuerySelect2)

                    end if

                end if

            next

                    'Response.Write "<br><br>" & QuerySelect & "<br><br>"


            if ValidoHomo = true then

                if QuerySelect3 <> "" then
                    'Response.Write QuerySelect3 & "<br>"
                    Conn.Execute(QuerySelect3)
                end if

                if QuerySelect <> "" then
                    QuerySelect = "UPDATE " & Iif(AWBType = "1","Awb","Awbi") & " SET Expired = 0" & QuerySelect & " WHERE AwbID = " & ObjectID
                    'Response.Write QuerySelect & "<br>"
                    Conn.Execute(QuerySelect)
                end if


            else
                
                if Action = "insert" then
                    response.write "<script" & ">alert('Cliente seleccionado no esta homologado');</script>"
                end if

            end if

            Response.Write "Finalizo Proceso<br>"


    end select 

   CloseOBJ Conn

end if
%>

