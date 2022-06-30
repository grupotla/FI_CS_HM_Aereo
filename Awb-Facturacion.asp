<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, ObjectID, DocTyp, CountList9Values, aList9Values, QuerySelect, rs, conn, ConnMaster, FacID, FacType, FacStatus, i, esquema

GroupID = CheckNum(Request("GID"))
ObjectID = CheckNum(Request("ObjectID"))
DocTyp = CheckNum(Request("DocTyp"))
esquema = Request("esquema")
%>

<HTML><HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">

<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">

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
.style11 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
}
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

.ids    {   border:0px;
            color:white;
            font-weight:normal;
            background:gray;
            font-size: 10px;
            width:auto; 
            }
            
.readonly { border:0px;
            background:silver;
            color:navy;
            font-size: 10px; 
            font-family: Verdana, Arial, Helvetica, sans-serif;  
            width:auto; }            
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

<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="Javascript:self.focus();">

<%
    OpenConn Conn
	    'Obteniendo listado de Rubros
	    CountList9Values = -1
                          '     0           1       2           3     4     5       6       7           8               9           10          11      12     13  14  15  16               17
        QuerySelect = "Select ItemID, AgentTyp, CurrencyID, Local, Value, ItemName, Pos, ServiceID, ServiceName, PrepaidCollect, InvoiceID, CalcInBL, DocType, '', '', '', '', COALESCE(Regimen,'') from ChargeItems where Expired=0 and AwbID=" & ObjectID & " and DocTyp=" & DocTyp & " order by AgentTyp"
        'response.write QuerySelect & "<br>"
	    Set rs = Conn.Execute(QuerySelect)	
	    If Not rs.EOF Then
		    aList9Values = rs.GetRows
		    CountList9Values = rs.RecordCount - 1
	    End If

    openConnBAW Conn
    OpenConn2 ConnMaster

    for i=0 to CountList9Values

	    FacID = CheckNum(aList9Values(10,i))
        FacType = CheckNum(aList9Values(12,i))
        FacStatus = 0

        if FacID<>0 then
	        Select case FacType
            case 1
                set rs = Conn.Execute("select tfa_serie, tfa_correlativo, tfa_ted_id from tbl_facturacion where tfa_id=" & FacID)
                If Not rs.EOF Then
			        aList9Values(13,i) = "FC-" & rs(0) & "-" & rs(1)
                    FacStatus = CheckNum(rs(2))
                end if
		        CloseOBJ rs

            case 4
                set rs = Conn.Execute("select tnd_serie, tnd_correlativo, tnd_ted_id from tbl_nota_debito where tnd_id=" & FacID)
			        aList9Values(13,i) = "ND-" & rs(0) & "-" & rs(1)
                    FacStatus = CheckNum(rs(2))
		        CloseOBJ rs

            case 9,10 'recibido por pedido exactus
                QuerySelect = "SELECT DISTINCT COALESCE(a.pedido_erp,''), COALESCE(a.estado,0), COALESCE(b.fc_numero,''), COALESCE(b.fc_estado,0), COALESCE(b.fc_saldo,0), COALESCE(c.nc_numero,'')  FROM exactus_pedidos a LEFT JOIN exactus_pedidos_fc b ON a.id_pedido = b.id_pedido  LEFT JOIN exactus_pedidos_nc c ON a.id_pedido = c.id_pedido WHERE a.id_pedido = " & FacID & " "
                'response.write QuerySelect & "<br>"
                set rs = ConnMaster.Execute(QuerySelect)
                If Not rs.EOF Then
	    
                    if rs(0) <> "" then
                        'aList9Values(13,i) = "PE-" & rs(0) 
                        aList9Values(13,i) = FacID & " - " & rs(0) 
                        FacStatus = 90 'enviada
                    end if

                    if rs(2) <> "" then
                        'aList9Values(13,i) = "FC-" & rs(2) 
                        aList9Values(13,i) = FacID & " - " & rs(2) 
                        FacStatus = 91 'facturada
                    end if

                    'if  rs(5) <> "" then 'nunca entra aca
                    '    aList9Values(13,i) = "NC-" & rs(5) 
                    '    FacStatus = 92 'anulada
                    'end if

                end if		
		        CloseOBJ rs

            end Select

        End If


        'Indicando el Estado de Pago de la Factura/ND
        select Case FacStatus
        case 2
            aList9Values(14,i) = "<font color=blue>ABONADO</font>"
        case 4
            aList9Values(14,i) = "<font color=blue>PAGADO</font>"

        case 90 '2021-08-06
            aList9Values(14,i) = "<font color=blue>ENVIADO</font>"

        case 91 '2021-08-16
            aList9Values(14,i) = "<font color=blue>FACTURADO</font>"

        case 92 '2021-08-06
            aList9Values(14,i) = "<font color=blue>CANCELADO</font>"

        case Else
            aList9Values(14,i) = "<font color=red>PENDIENTE</font>"
        End Select


  
        'si funciona pero aun no esta autorizado mostrarlo solo 1237
        'if Session("OperatorID") = "1237" then 

            aList9Values(16,i) = "<img src='img/glyphicons_192_circle_remove1.png'>"

            'agregar eh_erp_esquema al filtro 2022-04-19 hoy se definio esta mejora
            QuerySelect = "SELECT a.codigo, COALESCE(eh_erp_codigo,''), COALESCE(eh_estado,0) " & _	
	"FROM vw_rubros_combinaciones a " & _
	"LEFT JOIN exactus_homologaciones ON codigo = eh_codigo AND eh_erp_categoria = '06' AND eh_estado = 1 AND eh_erp_esquema = '" & esquema & "' " & _     
    "WHERE a.id_servicio = " & aList9Values(7,i) & " AND a.id_rubro = " & aList9Values(0,i) & " " & _
    " AND a.d1 = '" & Iif(DocTyp = "1","IMPORT","EXPORT") & "' AND a.d2 = 'A' AND a.d3 = '" & aList9Values(17,i) & "' "
         
            'response.write QuerySelect & "<br>"
            set rs = ConnMaster.Execute(QuerySelect)
            If Not rs.EOF Then
                aList9Values(15,i) = rs(0) '"AE" & aList9Values(7,i) & "-IM-A-" & aList9Values(0,i)
       	        aList9Values(16,i) = "<font color=" & Iif(rs(2) = "1","green","white") & ">" & rs(1) & "</font>"
            end if
	        CloseOBJ rs
        
        'end if

        'response.write "(" & aList9Values(15,i) & ")(" & aList9Values(13,i) & ")(" & aList9Values(14,i) & ")" & "<br>"

    next



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
%>



    <center><h3 class="menu" ><font color=white><%=ObjectID & " - " & Iif(Request("HAWBNumber") = "", "AWBNumber : " & Request("AWBNumber"), "HAWBNumber : " & Request("HAWBNumber")) %>  &nbsp; - &nbsp; <%=Iif(DocTyp=0,"EXPORT","IMPORT")%></font></h3></center>

    <table width="80%" border="0">
        <tr>

            <%'if Session("OperatorID") = "1237" then %>
		    <td align="center" class="style4">
		        <font class="style8">Articulo</font>
            </td>

		    <td align="left" class="style4">
		        <font class="style8">Homologado</font>
            </td>          
            <%'end if %>

		    <td align="center" class="style4">
                <font class="style8">Servicio</font>
            </td>
		    <td align="center" class="style4">
                <font class="style8">Rubro</font>
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


		    <td align="center" class="style4">
		        <font class="style8">Pedido / Factura / ND</font>
            </td>
		    <td align="center" class="style4">
		        <font class="style8">Estado</font>
            </td>


        </tr>


<%
if CountList9Values>=0 then

	for i=0 to CountList9Values

%>

    <tr bgcolor="">

           
        <%'if Session("OperatorID") = "1237" then %>
	        <td align="right" class="style4" nowrap><%=aList9Values(15,i)%></td>

	        <td align="right" class="style4" style="background-color:white;text-align:center">
		        <%=aList9Values(16,i)%>		
            </td>    
        <%'end if %>

		<td align="right" class="style4"> 
			<input type="text" size="18" class="style10" value="<%=aList9Values(7,i) & " - " & aList9Values(8,i)%>" id="SVNO1" readonly>
		</td>

		<td align="right" class="style4">
			<input type="text" class="style10" value="<%=aList9Values(0,i) & " - " & aList9Values(5,i)%>" size="25" readonly>		
        </td>
                      				
		<td align="right" class="style4">
			<input type="text" class="style10" value="<%=aList9Values(2,i)%>" size="5" readonly>	 <!-- moneda -->			
        </td>
                                            				
		<td align="right" class="style4">
			<input type="text" class="style10" value="<%=aList9Values(4,i)%>" size="20" readonly>	 <!-- valor -->	
        </td>
                      				
		<td align="right" class="style4">
			<input type="text" size="5" class="style10" value="<%=IntLoc(CheckNum(aList9Values(3,i)))%>"  readonly>		
        </td>
               
		<td align="right" class="style4">
			<input type="text" size="5" class="style10" value="<%=PrepColl(CheckNum(aList9Values(9,i)))%>"  readonly>		
        </td>

        <td align="right" class="style4">
			<input type="text" class="style10" value="<%=aList9Values(13,i)%>" size="25" readonly>		
        </td>

		<td align="right" class="style4" style="background-color:white">
			<%=aList9Values(14,i)%>		
        </td>




	</tr>

<%
    next

end if
%>
    </table>

</BODY>
</HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>


