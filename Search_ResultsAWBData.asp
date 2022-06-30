<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, HTMLCode, HTMLTitle, Table, QuerySelect, MoreOptions, Agents
Dim Option1, Option2, Option3, Option4, Conn, rs, JavaMsg, i, link, ItemPos
Dim Name, AccountNo, IATANo, NameES, OrderWord, Routing, AWBType, ItemID, Identifier
Dim elements, PageCount, AbsolutePage, HTMLHidden, Status, ItemName, ListColor

GroupID = CheckNum(Request.Form("GID"))

if GroupID >= 7 and GroupID <=24 then
	AbsolutePage = CheckNum(Request.Form("P"))
	if AbsolutePage = 0 then
		 AbsolutePage = 1
	end if
	elements = 5
	PageCount = 0
	OrderWord = "a.Countries, a.Name"
    Select case GroupID
	case 7, 21, 22, 23, 24
			 'OrderWord = "p.codigo, a.nombre_cliente"
             OrderWord = "a.nombre_cliente"
			 QuerySelect = 	"select a.id_cliente, p.codigo, a.nombre_cliente, d.id_direccion, a.es_coloader from clientes a, direcciones d, niveles_geograficos n, paises p "						
			 HTMLTitle = "<td class=titlelist><b>Pais</b></td><td class=titlelist><b>Codigo</b></td><td class=titlelist><b>Nombre</b></td>"
			 Name = Request.Form("Name")
			 Option1 = " a.id_cliente = d.id_cliente " & _
							"and d.id_nivel_geografico = n.id_nivel " & _
							"and n.id_pais = p.codigo " & _
							"and a.es_consigneer = true " & _
							"and a.id_estatus in (1,2) "
							
			 if Name <> "" then
			 		Option2 = " a.nombre_cliente ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 8
			 OrderWord = "a.agente"
			 'QuerySelect = "select a.AccountNo, a.Name, a.Address, a.Phone1, a.Phone2, a.IATANo, Countries, a.Address2, AgentID from Agents a "
			 QuerySelect = "select a.accountno, a.agente, a.direccion, a.telefono, a.fax, contacto, a.iatano, a.countries, a.agente_id, a.es_neutral from agentes a"
			 HTMLTitle = "<td class=titlelist><b>Codigo</b></td><td class=titlelist><b>Nombre Agente</td>"
			 Name = Request.Form("Name")
			 Option1 = "a.activo=true "
			 if Name <> "" then
			 		Option2 = " a.agente ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 10
			 'OrderWord = "p.codigo, a.nombre_cliente"
             OrderWord = "a.nombre_cliente"
			 QuerySelect = 	"select a.id_cliente, p.codigo, a.nombre_cliente, d.id_direccion, a.es_coloader from clientes a, direcciones d, niveles_geograficos n, paises p "						
			 HTMLTitle = "<td class=titlelist><b>Pais</b></td><td class=titlelist><b>Codigo</b></td><td class=titlelist><b>Nombre</b></td>"
			 Name = Request.Form("Name")
			 Option1 = " a.id_cliente = d.id_cliente " & _
							"and d.id_nivel_geografico = n.id_nivel " & _
							"and n.id_pais = p.codigo " & _
							"and a.es_shipper = true " & _
							"and a.id_estatus in (1,2) " 'Ticket#2016111704000465  Problemas al ingresar al modulo de Aereo 
                            'a partir del nuevo catalogo los clientes nacen con estatus 2, pero en aereo pueden seleccionarlos
                            'para proceder a generar guia CIF 2016-11-18
			 if Name <> "" then
			 		Option2 = " a.nombre_cliente ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 11
			 OrderWord = "a.NameES"
			 QuerySelect = 	"select a.CommodityCode, a.NameES, a.CommodityId, a.TypeVal from Commodities a "
			 HTMLTitle = "<td class=titlelist><b>C&oacute;digo Producto</b></td><td class=titlelist><b>Nombre</b></td>"
			 NameES = Request.Form("NameES")
			 Option1 = " a.Expired=0 "
			 if NameES <> "" then
			 		Option2 = " a.NameES ilike '%" & NameES & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='NameES' type=hidden value='" & NameES & "'>"
	case 17
			 OrderWord = "id_routing desc"
			 QuerySelect = 	"select a.id_routing, a.routing from routings a left join seguros b on (a.id_routing=b.id_routing and b.anulado=false) "
			 HTMLTitle = "<td class=titlelist><b>Routing</b></td>"
			 Routing = Request.Form("Routing")
			 AWBType = CheckNum(Request.Form("AT"))
			 'Option1 = " a.id_transporte=1 and id_routing_type=2 "'Transporte Aereo, Routings Tipo Internos
			 Option1 = " a.id_transporte=1 and a.id_routing_type=2 and (a.activo=true or a.seguro=true) and a.borrado=false "'Transporte Aereo, Routings Tipo Internos

             'Option1 = Option1 & " and (a.bl_id = 0 or a.bl_id IS NULL) AND a.id_pais IN " & Session("Countries") '2016-04-01
             Option1 = Option1 & " and (a.bl_id = 0 or a.bl_id IS NULL) "  '2016-04-29

             Option1 = Option1 & " and a.id_cliente_order IS NOT NULL  "  '2016-10-10 para pruebas, esto quitarlo

			 if Routing <> "" then
			 		Option2 = " a.routing ilike '%" & Routing & "%' "
			 end if

			 if AWBType = 1 then 'Export
			 		Option3 = " a.import_export=false "
			 else 'Import
			 		Option3 = " a.import_export=true "
			 end if


             if AWBType = 1 then 'Export
             end if

			 HTMLHidden = HTMLHidden & "<INPUT name='Routing' type=hidden value='" & Routing & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='AT' type=hidden value='" & AWBType & "'>"
	case 18, 19
			 OrderWord = "a.desc_rubro_es"
			 QuerySelect = 	"select a.id_rubro, a.desc_rubro_es from rubros a "
			 HTMLTitle = "<td class=titlelist><b>ID</b></td><td class=titlelist><b>Routing</b></td>"
			 Name = Request.Form("Name")
			 Option1 = " a.id_estatus=1 "'Rubro Activo
			 if Name <> "" then
			 		Option2 = " a.desc_rubro_es ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 20
			 Agents = Request("ST")
			 ItemPos = CheckNum(Request("N"))
			 Name = Request("Name")
				 
			 select Case Agents
			 Case 0 'Linea Aereo
				 OrderWord = "a.name"
				 QuerySelect = 	"select a.carrier_id, a.name, b.es_afecto, 0 from carriers a, regimen_tributario b "
				 HTMLTitle = "<td class=titlelist><b>Codigo</b></td><td class=titlelist><b>L&iacute;nea&nbsp;Aerea</b></td>"
				 Option1 = " a.tiporegimen=b.id_regimen and trim(a.name)<>'' "
				 if Name <> "" then
						Option2 = " a.name ilike '%" & Name & "%' "
				 end if
			 Case 1 'Agentes
				 OrderWord = "a.agente"
				 QuerySelect = 	"select a.agente_id, a.agente, b.es_afecto, a.es_neutral from agentes a, regimen_tributario b "
				 HTMLTitle = "<td class=titlelist><b>Codigo</b></td><td class=titlelist><b>Agente</b></td>"
				 Option1 = " a.tiporegimen=b.id_regimen and a.activo=true and trim(a.agente)<>'' "
				 if Name <> "" then
						Option2 = " a.agente ilike '%" & Name & "%' "
				 end if
			 Case 2 'Naviera
				 OrderWord = "a.nombre"
				 QuerySelect = 	"select a.id_naviera, a.nombre, b.es_afecto, 0 from navieras a, regimen_tributario b "
				 HTMLTitle = "<td class=titlelist><b>Codigo</b></td><td class=titlelist><b>Naviera</b></td>"
				 Option1 = "a.tiporegimen=b.id_regimen and a.activo=true and trim(a.nombre)<>'' "
				 if Name <> "" then
						Option2 = " a.nombre ilike '%" & Name & "%' "
				 end if
			 Case 3	'Proveedores (Otros)
				 OrderWord = "a.nombre"
				 QuerySelect = 	"select a.numero, a.nombre, b.es_afecto, 0 from proveedores a, regimen_tributario b "
				 HTMLTitle = "<td class=titlelist><b>Codigo</b></td><td class=titlelist><b>Proveedor</b></td>"
				 Option1 = " a.tiporegimen=b.id_regimen and a.status in (0,1) and trim(a.nombre)<>'' "
				 if Name <> "" then
						Option2 = " a.nombre ilike '%" & Name & "%' "
				 end if
			 End Select
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='ST' type=hidden value='" & Agents & "'>"
	end select

	MoreOptions = 0
	CreateSearchQuery QuerySelect, Option1, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option2, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option3, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option4, MoreOptions, " and "

    if GroupID = 17 And AWBType = 1 then 'Export solo cuando es export realiza el union 2017-06-26 Ticket#2017062604000161  RO Exportación aerea PTY 
        QuerySelect = QuerySelect & " UNION " & Replace(QuerySelect,"a.import_export=false","a.import_export=true") & " and a.id_pais_origen IN " & Session("Countries")                            
    end if    
    
    QuerySelect = QuerySelect & " Order By " & OrderWord    

    response.write "<script>console.log('" & Replace(QuerySelect,"'","") & "');</script>" 

	'response.write GroupID & "<br>" 
	'response.write QuerySelect & "<br>"
    'response.write AWBType & "<br>"
    'response.write "<" & "script" & ">console.log(""" & QuerySelect & """);<" & "/" & "script>" 

    HTMLCode = ""

select Case GroupID
case 7, 8, 10, 11, 17, 18, 19, 20, 21, 22, 23, 24
	OpenConn2 Conn
case else
	OpenConn Conn
end select

'Buscando los archivos que coinciden con el query de Busqueda
    Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
	'if False then
		'Obteniendo la cantidad de resultados por busqueda
		rs.PageSize = 10
		'Saltando a la pagina seleccionada
  	  	rs.AbsolutePage = AbsolutePage
		PageCount = rs.PageCount
		'Desplegando los resultados de la pagina
		select Case GroupID
		case 7, 10, 21, 22, 23, 24
			for i=1 to rs.PageSize
                if CheckNum(rs(4))=0 then
                    ListColor = "list"
                    Identifier = ""
                else
                    ListColor = "listwarning"
                    Identifier = "[Coloader]"
                end if
				HTMLCode = HTMLCode & "<tr><td class=" & ListColor & "><a class=labellist href=SetMaster.asp?GID=" & GroupID & "&OID=" & rs(0) & "&AID=" & rs(3) & ">" & rs(1) & "</a></td>" & _
				"<td class=" & ListColor & "><a class=labellist href=SetMaster.asp?GID=" & GroupID & "&OID=" & rs(0) & "&AID=" & rs(3) & ">" & rs(0) & "</a></td>" & _
				"<td class=" & ListColor & "><a class=labellist href=SetMaster.asp?GID=" & GroupID & "&OID=" & rs(0) & "&AID=" & rs(3) & ">" & rs(2) & "&nbsp;" & Identifier & "</a></td></tr>"
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		case 8
			for i=1 to rs.PageSize
                if CheckNum(rs(9))=0 then
                    ListColor = "list"
                    Identifier = ""
                else
                    ListColor = "listwarning"
                    Identifier = "[Neutral]"
                end if
				link = "<td class=" & ListColor & "><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0].AccountAgentNo.value = '"& rs(0) & "';" & _
				"top.opener.document.forms[0].IATANo.value = '" & rs(6) & "';" & _
				"top.opener.document.forms[0].AgentData.value = '" & rs(1) & "\n" & rs(2) & "\n" & _
				rs(3) & "    " & rs(4)
				if rs(5) <> "" then
					link = link & "\nATTN: " & rs(5) & "\n"
				end if
				link = link & "';top.opener.document.forms[0].AgentID.value=" & rs(8) & ";" & _
                "top.opener.document.forms[0].AgentNeutral.value = '" & CheckNum(rs(9)) & "';top.close();"">"
				
				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(8) & "</a></td>" & _
				link & rs(1) & "&nbsp;" & Identifier & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		case 11
			for i=1 to rs.PageSize
				link = "<td class=list><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0].Commodities.value = top.opener.document.forms[0].Commodities.value + '"& rs(2) & "\n';" & _
				"top.opener.document.forms[0].NatureQtyGoods.value = top.opener.document.forms[0].NatureQtyGoods.value + '"& rs(1) & "\n';" & _
				"top.opener.document.forms[0].CommoditiesTypes.value = top.opener.document.forms[0].CommoditiesTypes.value + '"& rs(3) & ",';" & _
				"top.close();"">"
				
				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(2) & "</a></td>" & _
				link & rs(1) & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		case 17
			for i=1 to rs.PageSize
				link = "<td class=list><a class=labellist href='SetRouting.asp?RID=" & rs(0) & "&AT=" & AWBType & "'>"

				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(1) & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		case 18
			ItemName = Request("N")
			ItemID = Request("NID")
			for i=1 to rs.PageSize
				link = "<td class=list><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0]." & ItemID & ".value = "& rs(0) & ";" & _
				"top.opener.document.forms[0].N" & ItemID & ".value = '"& trim(rs(1)) & "';" & _
				"top.opener.document.forms[0]." & ItemName & ".value = '"& trim(rs(1)) & "';" & _
				"top.close();"">"
				
				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(0) & "</a></td>" & _
				link & trim(rs(1)) & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		case 19 'Para insertar Rubros en Costos de MAWB
			ItemID = Request("N")
			for i=1 to rs.PageSize
				link = "<td class=list><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0].I" & ItemID & ".value = "& rs(0) & ";" & _
				"top.opener.document.forms[0].N" & ItemID & ".value = '"& trim(rs(1)) & "';" & _
				"top.close();"">"
				
				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(0) & "</a></td>" & _
				link & trim(rs(1)) & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		case 20
			for i=1 to rs.PageSize
				if CheckNum(rs(3))=0 then
                    ListColor = "list"
                    Identifier = ""
                else
                    ListColor = "listwarning"
                    Identifier = "[Neutral]"
                end if
                
                link = "<td class=" & ListColor & "><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0].SI" & ItemPos & ".value = "& rs(0) & ";" & _
                "top.opener.document.forms[0].SN" & ItemPos & ".value = '"& trim(replace(rs(1),"""","",1,-1)) & "';" & _
				"top.opener.document.forms[0].SAF" & ItemPos & ".value = "& rs(2) & ";" & _
				"top.opener.document.forms[0].SNEU" & ItemPos & ".value = "& rs(3) & ";" & _
                "top.opener.ValidarDoble("& ItemPos & ");top.close();"">"				
				
				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(0) & "</a></td>" & _
				link & trim(rs(1)) & "&nbsp;" & Identifier & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		end select
	else
		JavaMsg = "No hay resultados para esta busqueda"

        if GroupID = 17 then 
                        
            QuerySelect = "select a.id_routing, a.routing, b.anulado, a.id_transporte, a.id_routing_type, a.activo, a.seguro, a.borrado, a.bl_id, a.no_bl, a.import_export, a.id_pais from routings a left join seguros b on (a.id_routing=b.id_routing) where a.routing ilike '%" & Routing & "%' limit 1"
            'response.write QuerySelect & "<br>"
            Set rs = Conn.Execute(QuerySelect)
	        if Not rs.EOF then

                JavaMsg = "Routing " & Routing & "\n=================================\n"

                if rs("anulado") = "1" then 
                    JavaMsg = JavaMsg & "Seguro fue anulado\n"
                end if

                if rs("id_transporte") <> "1" then
                    JavaMsg = JavaMsg & "Transporte no es aereo\n"
                end if

                if rs("id_routing_type") <> "2" then
                    JavaMsg = JavaMsg & "Tipo de Routing no es interno\n"
                end if

                if rs("activo") <> "1" then
                    JavaMsg = JavaMsg & "No esta Activo\n"
                end if

                if rs("seguro") <> "1" then
                    JavaMsg = JavaMsg & "No tiene seguro\n"
                end if

                if rs("borrado") <> "0" then
                    JavaMsg = JavaMsg & "Fue borrado\n"
                end if

                if rs("bl_id") <> "0" then
                    JavaMsg = JavaMsg & "Ya tiene BL : " & rs("bl_id") & "\n"
                end if

                if rs("no_bl") <> "" then
                    JavaMsg = JavaMsg & "BL No. : " & rs("no_bl") & "\n"
                end if



			     if AWBType = 1 then 'Export

                    if rs("import_export") <> "0" then
                        JavaMsg = JavaMsg & "No es Export\n"
                    end if

			     else 'Import

                    if rs("import_export") <> "1" then
                        JavaMsg = JavaMsg & "No es import\n"
                    end if
			 		    
			     end if

                'if rs("id_pais") <> "" then
                '    JavaMsg = JavaMsg & "Pais es : " & rs("id_pais") & "\n"
                'end if

            end if

        end if

	end if
CloseOBJs rs, Conn

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
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<%if JavaMsg <> "" then
			 Response.Write "<SCRIPT>alert('" & JavaMsg & "');</SCRIPT>"
		end if
	%>
	<FORM name="forma" action="Search_ResultsAWBData.asp" method="post" target=_self>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
 	<INPUT name="Action" type=hidden value=1>
	<INPUT name="P" type=hidden value=1>
	<INPUT name="N" type=hidden value="<%=Request("N")%>">
	<INPUT name="NID" type=hidden value="<%=ItemID%>">
    <%=HTMLHidden%>

    
	<TABLE cellspacing=0 cellpadding=2 width=430 align=center>
		<TR>
		<TD width=430 colspan=2 class=label align=right valign=top>
			<TABLE cellspacing=5 cellpadding=2 width=430 align=center>
				<%=HTMLTitle%>
				<%=HTMLCode%>
			</TABLE>
		</TD>
	  </TR>

      

<% if PageCount > 1 then%>
		<TR>
		<TD width=430 colspan=2 class=label align=right valign=top>
				<TABLE cellspacing=5 cellpadding=2 width=430 align=center>
				<TR>
				<TD class=label align=left valign="top" width=65>
				<%if AbsolutePage > 1 then%>&nbsp;
								<a class=label onclick=JavaScript:NextPage("<%=(AbsolutePage-1)%>"); href=# target=_self><u><< Anterior</u></a>&nbsp;
				<%else%>
								<a class=label href="Search_AWBData.asp?GID=<%=GroupID%>" target=_self><u><< Regresar</u></a>&nbsp;

				<%end if%>&nbsp;
				</TD>
				<TD class=label align=center width="300">
							 <%
							 for i = 1 to PageCount
							 		 Response.write " <a class=label onclick=JavaScript:NextPage(" & i & ") href=#><u>" & i & "</u></a> "
							 		 if i <> PageCount then
							 		 		Response.write "<font class=label>|</font>" 
							 		 end if
									 'if (i mod 12) = 0 then
									 	'	Response.write "<br>"
									 'end if
							 next
							 %>
				</TD>
				<TD class=label align=right valign="top" width=65>&nbsp;
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
<%end if%>		
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