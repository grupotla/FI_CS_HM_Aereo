<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"

Dim Conn, rs, ObjectID, Name, Address, Phone1, Phone2, Attn, AccountNo, IATANo, QuerySelect, GroupID, AddressID, ClientType

ObjectID = CheckNum(Request("OID"))
GroupID = CheckNum(Request("GID"))
AddressID = CheckNum(Request("AID"))

if ObjectID <> 0 then
	QuerySelect = "select a.nombre_cliente, d.direccion_completa, d.""phone_number"", a.id_cliente, a.es_coloader " & _
							"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
							" where a.id_cliente = d.id_cliente" & _
							" and d.id_nivel_geografico = n.id_nivel" & _
							" and n.id_pais = p.codigo" & _
							" and a.id_cliente = " & ObjectID  & _
							" and d.id_direccion = " & AddressID
	OpenConn2 Conn
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
		Name = rs(0)
		Address = PurgeData(rs(1))
        ClientType = CheckNum(rs(4))
	end if
	CloseOBJ rs

	set rs = Conn.Execute("select numero_telefono from cli_telefonos where id_cliente=" & ObjectID)
	if Not rs.EOF then
		Phone1 = rs(0)
	end if
	CloseOBJ rs

	set rs = Conn.Execute("select nombres, telefono from contactos where id_cliente=" & ObjectID)
	if Not rs.EOF then
		Attn = rs(0)
		Phone2 = rs(1)
	end if
	CloseOBJ rs

	set rs = Conn.Execute("select no_cuenta, no_iata from clientes_aereo where id_cliente=" & ObjectID)
	if Not rs.EOF then
		AccountNo = rs(0)
		IATANo = rs(1)
	end if
	CloseOBJs rs, Conn
%>
<HTML>
<HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT LANGUAGE="JavaScript">

<%Select Case GroupID
Case 23%>	
	top.opener.document.forms[0].id_cliente_orderData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	top.opener.document.forms[0].id_cliente_order.value=<%=ObjectID%>;	    
<%Case 7%>
	top.opener.document.forms[0].AccountConsignerNo.value = '<%=AccountNo%>';
	top.opener.document.forms[0].ConsignerData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	top.opener.document.forms[0].ConsignerID.value=<%=ObjectID%>;
	top.opener.document.forms[0].ConsignerAddrID.value=<%=AddressID%>;
    top.opener.document.forms[0].ConsignerColoader.value=<%=ClientType%>;
<%Case 10%>
	top.opener.document.forms[0].AccountShipperNo.value = '<%=AccountNo%>';
	top.opener.document.forms[0].ShipperData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	top.opener.document.forms[0].ShipperID.value=<%=ObjectID%>;
	top.opener.document.forms[0].ShipperAddrID.value=<%=AddressID%>;
    top.opener.document.forms[0].ShipperColoader.value=<%=ClientType%>;
<%Case 21%>
	top.opener.document.forms[0].ClientCollectID.value = '<%=ObjectID%>';
	top.opener.document.forms[0].ClientsCollect.value = '<%=Name%>';
<%Case 22%>	
	top.opener.document.forms[0].ColoaderData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	top.opener.document.forms[0].id_coloader.value=<%=ObjectID%>;
	//top.opener.document.forms[0].ConsignerAddrID.value=<%=AddressID%>;    
<%Case 24%>
	top.opener.document.forma.ConsignerID.value=<%=ObjectID%>;
	top.opener.document.forma.ConsignerData.value = '<%=PurgeData(Name)%>';
<%End Select%>	
	top.close();
</SCRIPT>
</BODY>
</HTML>
<%else%>
<SCRIPT LANGUAGE="JavaScript">
    top.close();
</SCRIPT>
<%end if%>		
