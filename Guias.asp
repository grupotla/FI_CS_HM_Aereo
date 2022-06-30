<%

iPerfilOpcion = PerfilOpcion()

'iPerfilOpcion = PerfilOpcion2(CheckNum(Request("GID")),Session("OperatorID"))

if iArr2.Item(Request("GID") & "Log") then

Checking "0|1"   

dim Guia, Comentarios, SelIniciada, Iniciada, UpdateDate, UpdatedTime, Estatus, SelEstatus, UserCreateStr, UserModifyStr
if CountTableValues >= 0 then
    Guia = aTableValues(1, 0)		    

	Estatus = aTableValues(2, 0)
    CreatedDate = ConvertDate(aTableValues(3, 0),2) 
    CreatedTime = aTableValues(10, 0)

    UserCreate = aTableValues(4, 0)
	UserModify = aTableValues(5, 0)	    

    UpdateDate = ConvertDate(aTableValues(11, 0),2)     
    UpdatedTime = aTableValues(12, 0)

	Expired = aTableValues(6, 0)	
    Iniciada = aTableValues(7, 0)	    

    CarrierID = aTableValues(8, 0)	
    Comentarios = aTableValues(9, 0)
    
    'GuideID, GuideNumber, GuideStatus, CreatedDate, CreatedUser, UpdatedUser, GuideActive, GuideType, GuideCarrierID, Comentarios, CreatedTime, UpdatedDate, UpdatedTime FROM "


else    
    
    CarrierID = Request("CarrierID")     

end if

    if Request("SelIniciada") = "" then
        SelIniciada = 2
    else
        SelIniciada = Request("SelIniciada") 
    end if

    if Request("SelEstatus") = "" then
        SelEstatus = 0
    else
        SelEstatus = Request("SelEstatus") 
    end if


    if CountTableValues = -1 then
        Expired = 1
    end if
    
    

    UserCreateStr = "" 
    UserModifyStr = ""


    'response.write "(Action=" & Action & ")"
    'response.write "(Request(CarrierID)=" & Request("CarrierID") & ")"
    'response.write "(CarrierID=" & CarrierID & ")"

if Session("OperatorLevel") = 0 then
	OpenConn2 Conn	
	'Obteniendo el listado de Grupos
	set rs = Conn.Execute("select id_grupo, nombre_grupo from grupos where id_estatus=1 order by nombre_grupo")
	if Not rs.EOF then
		aList3Values = rs.GetRows
		CountList3Values = rs.RecordCount-1
	end if
	CloseOBJ rs
	if UserCreate <> 0 then
		set rs = Conn.Execute("select pw_gecos from usuarios_empresas where id_usuario=" & UserCreate)
		if Not rs.EOF then
			UserCreate = UCASE(rs(0))
		end if
		CloseOBJ rs
	else
		UserCreate = ""
	end if
		
	if UserModify <> 0 then
		set rs = Conn.Execute("select pw_gecos from usuarios_empresas where id_usuario=" & UserModify)
		if Not rs.EOF then
			UserModify = UCASE(rs(0))
		end if
		CloseOBJ rs
	else
		UserModify = ""
	end if	
	CloseOBJ Conn
end if

%>
<!DOCTYPE html>
<html>
<head>



<TITLE>AWB - Aimar - Administración</TITLE>


<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
    function validar(Action) {

        if (Action != 3 && Action != 5) {

            if (!valTxt(document.forma.CarrierID, 1, 10)) { return (false) };

            if (document.forma.guia0.style.display != 'none')  		    
                if (!valTxt(document.forma.guia0, 3, 10)) { return (false) };

            if (document.forma.Expired.checked == false)                
                if (!valTxt(document.forma.Comentarios, 3, 10)) { return (false) };
        }

        if (Action == 5) {        
            document.forma.OID.value = 0;            
        }

        if (Action == 3 || Action == 5) {
            //document.forma.SelIniciada.value = 0;
            document.forma.guia0.value = '';
            //document.forma.guia1.value = '';
            //document.forma.guia2.value = '';
            document.forma.CD.value = '';
        }

		document.forma.Action.value = Action;
		document.forma.submit();
    }

    function ChangeSelIniciada() {
        document.forma.submit();
    }


</SCRIPT>

<style type="text/css">    
    .GridView thead td { background: #000066; font-weight :bolder; color:White;  }    
    .GridView tbody td { *text-align:center }    
    .GridView td, .GridView a { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 7.6pt; *text-align:center; }        
    
    .GridView a { color: navy; display:block; border:0px solid red; text-decoration: none;}                
    
    .GridView tbody tr:nth-child(odd) { background: rgb(182,216,237); }
    .GridView tbody tr:nth-child(even) { background: #ffffff; }
    .GridView tbody tr:hover { background: rgb(191,191,191); }
    .readonly { background-color:silver }    
    .input { background-color:white }    
</style>


<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">

</head>

<body>
<!--
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="Javascript:self.focus();">
-->


<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="SO" type=hidden value="<%=SearchOption%>">
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">


<table width=100% border=0 align=center valign=top>
<tr>

    <TD colspan=2 class="label" align=center>
    
        <b>AEROLINEA :</b><br />

	    <select name="CarrierID" id="Transportista" onChange="document.forma.OID.value='';document.forma.submit();" >
	    <option value="">Seleccionar</option>
	    <%        
            OpenConn Conn

            'Obteniendo listado de Carriers
	        Set rs = Conn.Execute("select CarrierID, Name, Countries from Carriers where Expired = 0 and Countries in " & Session("Countries") & " order by Name, Countries")
	        If Not rs.EOF Then
   		        aList1Values = rs.GetRows
       	        CountList1Values = rs.RecordCount
            End If
	        CloseOBJ rs

		    For i = 0 To CountList1Values-1
		        if aList1Values(0,i) = CarrierID then
			        CarrierName = aList1Values(1,i)
			        Countries = aList1Values(2,i)
		        end if
	            %>
	            <option value="<%=aList1Values(0,i)%>"><%=aList1Values(1,i) & " - " & aList1Values(2,i) & " - " & aList1Values(0,i)%></option>
	            <%
   		    Next
	    %>
	    </select>
    </TD>

</tr>
<tr>
<td width=100% valign=top style="display:block;height:400px;overflow-y:scroll;">

    <SELECT class="label" name="SelIniciada" id="SelIniciada" onChange="ChangeSelIniciada()" style="display:none">    
    <option value=0 <% if SelIniciada = 0 then %> selected <% end if %> >Inactivos</option>
    <option value=1 <% if SelIniciada = 1 then %> selected <% end if %> >Activos</option>
    <option value=2 <% if SelIniciada = 2 then %> selected <% end if %> >Activos / Inacttivos</option>
    </SELECT>
    <span class="label">Seleccione Estado de Guias</span>
    <SELECT class="label" name="SelEstatus" id="SelEstatus" onChange="ChangeSelIniciada()">        
    <option value=0 <% if SelEstatus = 0 then %> selected <% end if %> >Todas sin Awb / Hawb</option>
    <option value=1 <% if SelEstatus = 1 then %> selected <% end if %> >Tiene Awb / Hawb</option>
    </SELECT>

    <table border="0" align="center" width="100%" class="GridView" cellpadding="2" cellspacing="0" >
    <thead>
    <tr>
        <td>ID</td>
        <td>Guia</td>
        <!--<td>Estado</td>-->
        <td>Usuario Crea</td>        
        <td>Fecha Crea</td>        
        <td>Estado</td>
        <td>Usuario Modif</td>        
        <td>Fecha Modif</td>        
        <!--<td>Documento</td>-->
    </tr>
    </thead>

    <tbody>

<%    
    Function OnClick(index,text,rgb,doc)
        'if doc = "" then
            OnClick = "<a " & rgb & " href='InsertData.asp?GID=21&OID=" & index & "&SelIniciada=" & SelIniciada & "&SelEstatus=" & SelEstatus & "' title='Editar Guia " & index & "'>"  & text & "</a>"
        'else
        '    OnClick = "<a href='#' " & rgb & ">"  & text & "</a>"
        'end if
    End Function

 
    QuerySelect = "SELECT GuideID, GuideNumber, GuideStatus, g.CreatedDate, " 
    QuerySelect = QuerySelect & "ifnull(concat(a.FirstName , ' ' , a.LastName),''), ifnull(concat(b.FirstName , ' ' , b.LastName),''), " 
    QuerySelect = QuerySelect & "GuideActive, GuideType, GuideCarrierID, Comentarios, g.CreatedTime, UpdatedDate, UpdatedTime " 
    QuerySelect = QuerySelect & "FROM  Guides g " 
    QuerySelect = QuerySelect & "LEFT JOIN Operators a ON a.OperatorID = CreatedUser " 
    QuerySelect = QuerySelect & "LEFT JOIN Operators b ON b.OperatorID = UpdatedUser " 
    QuerySelect = QuerySelect & "WHERE GuideCountry = '" & Session("OperatorCountry") & "' AND GuideStatus = '" & SelEstatus & "' AND GuideCarrierID = '" & CarrierID & "' "

    if SelIniciada < 2 then
    QuerySelect = QuerySelect & "AND GuideActive = " & cInt(Request("SelIniciada")) & " "
    end if

    'QuerySelect = QuerySelect & "ORDER BY GuideID DESC" 

    QuerySelect = QuerySelect & "ORDER BY GuideNumber" 

    Conn.Errors.Clear()
    on error resume next

    Set rs = Conn.Execute(QuerySelect)         
    dim rgb, usado, fec, estado, documento, rs_ch, ConnBaw1, rs_bw

    openConnBAW ConnBaw1

    usado = 0
    if Conn.Errors.Count = 0 then
	    If Not rs.EOF Then
		    Do While Not rs.EOF

                documento = ""
                'QuerySelect = "SELECT distinct b.AWBID, b.DocTyp, b.InvoiceID, b.DocType, b.Expired FROM Awb a, ChargeItems b WHERE a.AwbNumber = '" & rs(1) & "' AND a.AwbID = b.AWBID AND b.Expired = '0' AND b.DocType IN (1,4) AND DocTyp = '0' LIMIT 1"
                'Set rs_ch = Conn.Execute(QuerySelect)  
	            'If Not rs_ch.EOF Then                    
                '    if rs_ch(3) = 1 then 'factura
                        'QuerySelect = "SELECT tfa_serie, tfa_correlativo, tfa_id, tfa_hbl, tfa_mbl FROM tbl_facturacion WHERE tfa_id = '" & rs_ch(2) & "' AND tfa_ted_id != '3' LIMIT 1"
                        'Set rs_bw = ConnBaw1.Execute(QuerySelect)  
                        'documento = rs_bw(0) & rs_bw(1)
                '    end if                    
                '    if rs_ch(3) = 4 then 'nota credito
                        'QuerySelect = "SELECT tnc_serie, tnc_correlativo FROM tbl_nota_credito WHERE tnc_id = '" & rs_ch(2) & "' AND tnc_ted_id != '3' LIMIT 1"
                        'Set rs_bw = ConnBaw1.Execute(QuerySelect)  
                        'documento = rs_bw(0) & rs_bw(1)
                '    end if
                'end if

                rgb = ""                    
                if CInt(rs(6)) = 1 then 'activo
                    if CInt(rs(7)) = 1 then 'iniciada
                        estado = "Iniciada"
                        rgb = " style='color:brown' "
                        usado = 1
                    else
                        estado = "Activa"
                    end if                        
                else 
                    estado = "Inactiva"
                    rgb = " style='color:gray' "
                end if

                if rs(2) = "1" then 'tiene awb
                    estado = "Awb/Hawb"
                    rgb = " style='color:black' "
                end if

                'if rs(3) = 0 then
                '    fec = ""
                'else
                '    fec = Mid(rs(3),7,2) & "/" & Mid(rs(3),5,2) & "/" & Left(rs(3),4) 
                'end if

                %>
                <tr>
                <td><%=OnClick(rs(0),rs(0),rgb,documento)%></td>
                <td><%=OnClick(rs(0),rs(1),rgb,documento)%></td>                                            
                <td><%=OnClick(rs(0),rs(4),rgb,documento)%></td>                    
                <td><%=OnClick(rs(0),rs(3),rgb,documento)%></td>
                <td title="<%=rs(9)%>"><%=OnClick(rs(0),estado,rgb,documento)%></td>
                <td><%=OnClick(rs(0),rs(5),rgb,documento)%></td>                    
                <%
                if rs(12) = 0 then
                fec = ""
                else
                fec = Mid(rs(12),7,2) & "/" & Mid(rs(12),5,2) & "/" & Left(rs(12),4) 
                end if
                %>
                <td><%=OnClick(rs(0),fec,rgb,documento)%></td>

                <!--<td><%=OnClick(rs(0),documento,rgb,documento)%></td>-->

                
                </tr>        
                <%
                rs.MoveNext
            Loop            		    
	    End If
    else   
        response.write("Error en query revisar " & QuerySelect )
    end if
        
	CloseOBJ rs

    if CountTableValues > -1 then
        QuerySelect = "SELECT  " 
        QuerySelect = QuerySelect & "ifnull(concat(a.FirstName , ' ' , a.LastName),''), ifnull(concat(b.FirstName , ' ' , b.LastName),'') "     
        QuerySelect = QuerySelect & "FROM  Guides g " 
        QuerySelect = QuerySelect & "LEFT JOIN Operators a ON a.OperatorID = CreatedUser " 
        QuerySelect = QuerySelect & "LEFT JOIN Operators b ON b.OperatorID = UpdatedUser " 
        QuerySelect = QuerySelect & "WHERE GuideID = '" & aTableValues(0, 0) & "' "
        Set rs = Conn.Execute(QuerySelect)         
	    If Not rs.EOF Then

            UserCreateStr = rs(0) 
            UserModifyStr = rs(1)

        end if
        CloseOBJ rs
    end if





    documento = "No tiene documento"

    if Guia <> "" then
        QuerySelect = "SELECT distinct b.AWBID, b.DocTyp, b.InvoiceID, b.DocType, b.Expired FROM Awb a, ChargeItems b WHERE a.AwbNumber = '" & Guia & "' AND a.AwbID = b.AWBID AND b.Expired = '0' AND b.DocType IN (1,4) AND DocTyp = '0' LIMIT 1"
        'response.write(QuerySelect & "<br>")
        Set rs_ch = Conn.Execute(QuerySelect)  
	    If Not rs_ch.EOF Then
            if rs_ch(3) = 1 then 'factura
                QuerySelect = "SELECT tfa_serie, tfa_correlativo, tfa_id, tfa_hbl, tfa_mbl FROM tbl_facturacion WHERE tfa_id = '" & rs_ch(2) & "' AND tfa_ted_id != '3' LIMIT 1"
            end if                    
            if rs_ch(3) = 4 then 'nota credito
                QuerySelect = "SELECT tnc_serie, tnc_correlativo FROM tbl_nota_credito WHERE tnc_id = '" & rs_ch(2) & "' AND tnc_ted_id != '3' LIMIT 1"
            end if

            'response.write(QuerySelect & "<br>")

            if rs_ch(3) = 1 or rs_ch(3) = 4 then 'factura
                Set rs_bw = ConnBaw1.Execute(QuerySelect)  
                documento = rs_bw(0) & rs_bw(1)
            end if

        end if

    end if

    CloseOBJ Conn

    CloseOBJ ConnBaw1


    dim ValidaBorrar
    dim ValidaActualizar
    
    ValidaBorrar = "if (confirm('Esta seguro de Borrar este registro ? ')) JavaScript" & ":validar(3); else return false;"
    ValidaActualizar = "JavaScript" & ":validar(2)"    

    if documento <> "No tiene documento" then
        ValidaBorrar = "alert('No se puede borrar porque ya tiene facturacion')"
        'ValidaActualizar = "alert('No se puede actualizar porque ya tiene facturacion')"        
    end if


    if Estatus = 1 then
        ValidaBorrar = "alert('No se puede borrar porque ya tiene relacionado AWB / HWB')"
        'ValidaActualizar = "alert('No se puede actualizar porque ya tiene relacionado AWB / HWB')"        
    end if

    



    %>

    </tbody>
    </table>
    <% 'if usado = 1 then %>
    <% 'response.write("<font color=brown>Correlativos en marron estan apartados para ser utilizados </font>") %>
    <% 'end if %>

</td>
<td width=25% valign=top align=left>


    <h1 class="label" style="background:orange;color:black;padding:4px;">Documento de Facturacion : </h1>
    <h1 class="label" style="background:blue;color:white;padding:4px;"><%=response.write(documento)%></h1>

	<TABLE cellspacing=3 cellpadding=2 width=300px align=center border=0  style="border:1px solid #6699CC">	
    <tr>
    <td colspan=3 align=center style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12pt; background:#000066; color:white">Administrador de Guias </td>
    </tr>
		<%if SearchOption = 1 then%>
		<TR><TD class="label" align=center colspan="2"><b>Agentes:</b></TD></TR> 
		<%end if%>
		<TR><TD class="label" align=right><b>ID :</b></TD><TD class="label" align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 		
		<TR><TD class="label" align=right><b>Guia :</b></TD><TD class="label" align=left>
            <INPUT TYPE=hidden name="GuiaAnt" id="GuiaAnt" value="<%=Guia%>" maxlength="20" size="20" readonly >

            <% 'response.write ( "(" & trim(UserCreate) & ")(" & Session("OperatorID") & ")" ) %>

            <INPUT TYPE=text name="guia0" id="Guia" value="<%=Guia%>" maxlength="20" size="20" <%If trim(UserCreate) <> "" And CInt(UserCreate) <> CInt(Session("OperatorID")) Then %> readonly class="label readonly" <% else %> class="label input" <% end if %> >

            </TD></TR>

		<TR><TD class="label" align=right><b>Comentarios :</b></TD></TR>
		<TR><TD class="label" align=right colspan="2">
 
 <!-- < % If trim(UserModify) <> "" And UserModify <> Session("OperatorID") AND Comentarios <> "" Then%> readonly class="label readonly" < % else %> class="label input" < % end if %>  -->

            <textarea name="Comentarios" id="Comentarios" cols="44" rows="4" class="label input"><%=Comentarios%></textarea>

            </TD></TR>

        <TR><TD class="label" align=right><INPUT name=Expired TYPE=checkbox class="label" <%If Expired = 1 Then%> checked <%End If%>></TD><TD class="label" align=left><b>Activa</b></TD></TR>

        <TR><TD class="label" align=right><INPUT name=Iniciada TYPE=checkbox class="label" <%If Iniciada = 1 Then%> checked <%End If%>></TD><TD class="label" align=left><b>Iniciada</b></TD></TR>

        <TR><TD class="label" align=right><INPUT name=Estatus TYPE=checkbox class="label" <%If Estatus = 1 Then%> checked <%End If%> ></TD><TD class="label" align=left nowrap><b>Tiene Awb/Hawb</b></TD></TR>

        <%
        fec = Mid(CreatedDate,9,2) & "/" & Mid(CreatedDate,6,2) & "/" & Left(CreatedDate,4) & " " & CreatedTime
        %>

		<TR><TD class="label" align=right><b>Creado :</b></TD><TD class="label" align=left style="background:silver"><%=UserCreateStr%></TD></TR>
		<TR><TD class="label" align=right><b>Fecha :</b></TD><TD class="label" align=left style="background:silver"><%=fec%></TD></TR>

        <%
        if UpdatedTime = 0 then
            fec = ""
        else
            fec = Mid(UpdatedTime,7,2) & "/" & Mid(UpdatedTime,5,2) & "/" & Left(UpdatedTime,4) & " " & Right(UpdatedTime,6)
        end if                       
        %>


		<TR><TD class="label" align=right nowrap><b>Modificado :</b></TD><TD class="label" align=left style="background:lightblue"><%=UserModifyStr%></TD></TR>
		<TR><TD class="label" align=right nowrap><b>Fecha :</b></TD><TD class="label" align=left style="background:lightblue"><%=fec%></TD></TR>

		<TR>
		    <TD colspan="2" class="label" align=center>
			    <TABLE cellspacing=0 cellpadding=2 width=200>
			    <TR />
			    <%if CountTableValues = -1 then%>                        
                    
				    <TD class="label" align=center colspan=2><INPUT <%=IIf(iArr2.Item(Request("GID") & "Ins") = "1","","disabled")%> name=enviar type=button onClick="JavaScript:validar(4)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class="label"></TD>

			    <%else%>

				    <%if SearchOption = 1 then%>
				    <!-- <TD class="label" align=center colspan=2><INPUT name=enviar type=button onClick="top.opener.document.forms[0].AccountAgentNo.value='<%=AccountNo%>';top.opener.document.forms[0].IATANo.value='<%=IATANo%>';top.opener.document.forms[0].AgentData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%>&nbsp;&nbsp;&nbsp;&nbsp;<%=Phone2%>';top.opener.document.forms[0].AgentID.value=<%=ObjectID%>;top.close();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class="label"></TD> -->
				    <%end if%>
                                        
				    <TD class="label" align=center colspan=2><INPUT <%=IIf(iArr2.Item(Request("GID") & "Upd") = "1","","disabled")%> name=enviar type=button value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class="label" onClick="<%=ValidaActualizar%>"></TD>
                    
				    <TD class="label" align=center colspan=2><INPUT <%=IIf(iArr2.Item(Request("GID") & "Del") = "1","","disabled")%> name=enviar type=button value="&nbsp;&nbsp;Borrar&nbsp;&nbsp;" class="label" onClick="<%=ValidaBorrar%>"></TD>

				    <TD class="label" align=center colspan=2><INPUT <%=IIf(iArr2.Item(Request("GID") & "Ins") = "1","","disabled")%> name=enviar type=button value="&nbsp;&nbsp;Nuevo&nbsp;&nbsp;" class="label" onClick="JavaScript:validar(5)"></TD>

			    <%end if%>
			    </TR>
			    </TABLE>
		    <TD>
		</TR>
	
	</TABLE>

</td>
</tr>
</table>

</FORM>

<script>
    <% if (Session("OperatorLevel") = 0) then %>
        selecciona('forma.BGID','<%=BusinessGID%>');
    <% end if %>    
    //document.getElementById('SelIniciada').value = '<%=SelIniciada%>';
    //ChangeSelIniciadaMaster(document.getElementById('SelIniciada').value);
    document.getElementById('Transportista').value = '<%=CarrierID%>';
    document.getElementById('Transportista').focus();
</script>
</BODY>
</HTML>



<%

else

    response.write("No tiene permisos para esta opcion")
    
end if

%>