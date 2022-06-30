<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
-->
<%
Response.CharSet = "utf-8"
%>
    <div id="myProgress">
        <div id="myBar">10%</div>
    </div>

<%
Checking "0|1|2"

Dim TrackingID, DocTyp

if CountTableValues >= 0 then
    TrackingID = aTableValues(0, 0)
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	AWBID = aTableValues(3, 0)
	ConsignerID = aTableValues(4, 0)
	Comment = aTableValues(5, 0)
	Val = aTableValues(6, 0)
    DocTyp = aTableValues(7, 0)
else
	AWBID=CheckNum(request("AWBID"))
	ConsignerID=CheckNum(request("CID"))
	CreatedDate = ""
	CreatedTime = ""
end if

Set aTableValues = Nothing

   

    
    Dim NotifyAgentID, NotifyClientID, NotifyShipperID, Header, selected, agente_nom
	
	OpenConn2 Conn
		'Obteniendo el listado de Status 
		SQLQuery = "select id, estatus, notificar_agente, notificar_cliente, notificar_shipper, publico from aimartrackings where air=1 and activo=1 "
                
        if AwbType = 1 then
        SQLQuery = SQLQuery & " and import = 0 order by estatus"
        else
        SQLQuery = SQLQuery & " and import = 1 order by estatus"
        end if

        set rs = Conn.Execute(SQLQuery)

        'response.write SQLQuery

	    'LLenando el listado html de los Status seleccionados
		Do While not rs.EOF

			selected = ""
			if Action = 99 AND CheckNum(Request.Form("BLStatus")) = rs(0) then
				selected = " selected "
			end if

			Estate = Estate & "<option value='" & rs(0) & "' " & selected & ">" & rs(1) & " --I"
            if rs(2)=1 then
                Estate = Estate & "A"
            end if
            if rs(3)=1 then
                Estate = Estate & "C"
            end if
            if rs(4)=1 then
                Estate = Estate & "S"
            end if
            if rs(5)=1 then
                Estate = Estate & "P"
            end if
            Estate = Estate & "</option>"
            
            RangeID = RangeID & "RangeID[" & rs(0) & "]='" & rs(1) & "';" & vbCrLf
            NotifyAgentID = NotifyAgentID & "NotifyAgentID[" & rs(0) & "]='" & rs(2) & "';" & vbCrLf
            NotifyClientID = NotifyClientID & "NotifyClientID[" & rs(0) & "]='" & rs(3) & "';" & vbCrLf
            NotifyShipperID = NotifyShipperID & "NotifyShipperID[" & rs(0) & "]='" & rs(4) & "';" & vbCrLf
			rs.MoveNext
		Loop
	CloseOBJ rs

    if Action = 99 and Comment = "" then
		SQLQuery = "select comentario from tracking_comentarios where id_estatus_pg = " & CheckNum(Request.Form("BLStatus")) & " "
        set rs = Conn.Execute(SQLQuery)
        if Not rs.EOF then		    
	        Comment = rs(0)
		end if        
	    CloseOBJ rs        
    end if




    aList1Values = GetaList1Values(AWBID,AWBType)  


    agente_nom = ""
    if CheckNum(aList1Values(4,0)) > 0 then
        SQLQuery = "SELECT agente FROM agentes WHERE agente_id = " & CheckNum(aList1Values(4,0)) & " AND activo = 't' "		
        'response.write(SQLQuery & "<br>")
        set rs = Conn.Execute(SQLQuery)
        if Not rs.EOF then		    
	        agente_nom = rs(0)
		end if        
	    CloseOBJ rs
    end if


    
    
    
    

    
    dim arr_contact_count_tmp, arr_contact_tmp 
    arr_contact_count_tmp = -1
    SQLQuery = ""
    if Action = 1 or Action = 2 or Action = 99 then 'desde el select de estatus para display de los contactos correspondientes
                           
        SQLQuery = GetSQLQuery(CheckNum(Request.Form("NAgentID")), CheckNum(Request.Form("NClientID")), CheckNum(Request.Form("NShipperID")), aList1Values, AwbType)  
		if SQLQuery <> "" then		    		    
            'response.write(SQLQuery & "<br>")
			set rs = Conn.Execute(SQLQuery)
		    if Not rs.EOF then
		        arr_contact_tmp = rs.GetRows
		        arr_contact_count_tmp = rs.RecordCount -1
		    end if
		    CloseOBJ rs
		end if

    end if

    CloseOBJ Conn

    


    dim modeDev, modeDevStr(2) 
    modeDevStr(1) = ""

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    if Action = 1 or Action = 2 then '2020-08-20


        On Error Resume Next

        modeDev = WsModeDev("aereo", Session("OperatorCountry"), "ModeDev2", 1) 'ws21 = 1
        modeDevStr(1) = modeDev(1)
  
        If Err.Number <> 0 Then
            response.write Err.Number & "::" & Err.description        
        end if

         if InStr(1,modeDevStr(1),"wsNotif") > 0 then '2020-09-04 contactos divisiones TRACKING PRUEBAS copia = Si
                
            dim result, MyEmail, produccion, ws21

            'if InStr(1,modeDevStr(1),"Test0") > 0 then '2020-09-04 contactos divisiones TRACKING PRUEBAS status = Inactivo
                produccion = "0" 'envia TEST en subject y no copia nadie solo a Desarrollo
            'end if

            if InStr(1,modeDevStr(1),"Test1") > 0 then '2020-09-04 contactos divisiones TRACKING PRUEBAS status = Activo
                produccion = "1" 'envia emails normalmente a todos los contactos
            end if

            'if InStr(1,modeDevStr(1),"ws210") > 0 then '2020-09-04 contactos divisiones TRACKING PRUEBAS tipo_contacto = Consulta
                ws21 = "0"  'utiliza el web service 32
            'end if

            if InStr(1,modeDevStr(1),"ws211") > 0 then '2020-09-04 contactos divisiones TRACKING PRUEBAS tipo_contacto <> Consulta
                ws21 = "1"  'utiliza el web service 21
            end if


             MyEmail =  "(" & TrackingID & ")(" & "" & ")(" & AWBID & ")(" & DocTyp & ")(" & Request.Form("BLStatus") & ")(" & produccion & ")(" & Session("Login") & ")(" & Request.ServerVariables("REMOTE_ADDR")  & ")<br>" 

             On Error Resume Next
    
                 result = WsNotification(TrackingID, "", AWBID, DocTyp, Request.Form("BLStatus"), produccion, Session("Login"), Request.ServerVariables("REMOTE_ADDR"), ws21)

                 response.write "&nbsp;<span style='font-size:16px;color:red'>" & result(1) & "</span>"

                 'response.write "(stat=" & result(0) & ")<br>"
                 'response.write "msg=" & result(1) & "<br>"
                 'response.write "(sent_si=" & result(2) & ")<br>"
                 'response.write "(sent_no=" & result(3) & ")<br>"
                 'response.write "(tracking_id=" & result(4) & ")<br>"
                 'response.write "(product=" & result(5) & ")<br>"
                 'response.write "(sub_product=" & result(6) & ")<br>"
                 'response.write "(impex=" & result(7) & ")<br>"
                 'response.write "(bl_id=" & result(8) & ")<br>"
                 'response.write "(status_id=" & result(9) & ")<br>"
                 'response.write "(produccion=" & result(10) & ")<br>"
                 'response.write "(user=" & result(11) & ")<br>"
                 'response.write "(ip=" & result(12) & ")<br>"
                 'response.write "(Countries=" & result(13) & ")<br>"
                 'response.write "(CountriesDest=" & result(14) & ")<br>"

                 MyEmail = ""
                 MyEmail = MyEmail &  "(stat=" & result(0) & ")<br>"
                 MyEmail = MyEmail &  "msg=" & result(1) & "<br>"
                 MyEmail = MyEmail &  "(sent_si=" & result(2) & ")<br>"
                 MyEmail = MyEmail &  "(sent_no=" & result(3) & ")<br>"
                 MyEmail = MyEmail &  "(tracking_id=" & result(4) & ")<br>"
                 MyEmail = MyEmail &  "(product=" & result(5) & ")<br>"
                 MyEmail = MyEmail &  "(sub_product=" & result(6) & ")<br>"
                 MyEmail = MyEmail &  "(impex=" & result(7) & ")<br>"
                 MyEmail = MyEmail &  "(bl_id=" & result(8) & ")<br>"
                 MyEmail = MyEmail &  "(status_id=" & result(9) & ")<br>"
                 MyEmail = MyEmail &  "(produccion=" & result(10) & ")<br>"
                 MyEmail = MyEmail &  "(user=" & result(11) & ")<br>"
                 MyEmail = MyEmail &  "(ip=" & result(12) & ")<br>"
                 MyEmail = MyEmail &  "(Countries=" & result(13) & ")<br>"
                 MyEmail = MyEmail &  "(CountriesDest=" & result(14) & ")<br>"

            If Err.Number <> 0 Then
                MyEmail = MyEmail & Err.Number & ":" & Err.description            
            end if
    
            'MyEmail = "(" & Session("OperatorCountry") & ")(modeDev=" & modeDevStr(1) & ")(test=" & produccion & ")(ws21=" & ws21 & ")<br>" & MyEmail
    
             'result = WsSendMails("GT", "soporte7@aimargroup.com", "Aereo Monitoreo",  Base64Encode(MyEmail), "Monitoreo", "", "", "")

             'response.write result(1) & ".."

         end if 'modDev

    end if





    'Enviando Notificaciones por Mail para Agente, Cliente, Shipper cuando se ingresa o actualiza informacion
    Select Case Action
    case 1,2
           
        if InStr(1,modeDevStr(1),"sendNotif") > 0 then '2020-09-04 contactos divisiones TRACKING PRUEBAS rechazo = Si

		'para pruebas se comento el if        
        'SE AGREGO GTLTF 2015-11-18
        'OR aList1Values(5,0) = "GTLTF" 
		'if aList1Values(5,0) = "BZ" OR aList1Values(5,0) = "GT" OR aList1Values(5,0) = "SV" OR aList1Values(5,0) = "HN" OR aList1Values(5,0) = "NI" OR aList1Values(5,0) = "CR" OR aList1Values(5,0) = "PA" then
            if SQLQuery <> "" then '2015-09-23 para que no de error
                SQLQuery = SendNotification(Request.Form("BLStatusName"), SQLQuery, AwbType, Request.Form("Comment"), aList1Values, CheckNum(Request.Form("NShipperID")),0)
            end if
        'end if

        end if

    End Select     






    Comment2 = "Comentario para guia aerea:<b> " & aList1Values(0,0)
    if aList1Values(1,0) <> "" then
        Comment2 = Comment2 & "-" & aList1Values(1,0)
    end if
    Comment2 = Comment2 & "</b>"

%>
    <center><font color=blue face=arial><b>    
    <%=aList1Values(5,0)%>
    ::
    <%if AWBType = 1 then %> EXPORT <% else %> IMPORT <% end if %>
     
    </b></font></center>



<HTML>
<HEAD><TITLE>Aimar - Aereo</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=utf8">
<!--
<meta charset="UTF-8">
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
-->
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var RangeID = new Array();
var NotifyAgentID = new Array();
var NotifyClientID = new Array();
var NotifyShipperID = new Array();
<%=RangeID%>
<%=NotifyAgentID%>
<%=NotifyClientID%>
<%=NotifyShipperID%>

    function SelectBLStatus() {  
        document.forma.NAgentID.value = NotifyAgentID[document.forma.BLStatus.value];
        document.forma.NClientID.value = NotifyClientID[document.forma.BLStatus.value];
        document.forma.NShipperID.value = NotifyShipperID[document.forma.BLStatus.value];
		document.forma.Action.value = 99;
        document.forma.submit();
    }

	function validar(Action) {
        
        if (Action == 1 || Action == 2) 
            SendingTracking();

        move();

        if (Action == 3) {
           document.forma.Action.value = Action;
           document.forma.submit();		
        }

	   	if (Action != 0) {
			if (!valSelec(document.forma.BLStatus)){return (false)};
			if (!valTxt(document.forma.Comment, 3, 5)){return (false)};
			document.forma.BLStatusName.value = RangeID[document.forma.BLStatus.value];
            document.forma.NAgentID.value = NotifyAgentID[document.forma.BLStatus.value];
            document.forma.NClientID.value = NotifyClientID[document.forma.BLStatus.value];
            document.forma.NShipperID.value = NotifyShipperID[document.forma.BLStatus.value];
			document.forma.Action.value = Action;
		} else {
			document.forma.OID.value = 0;
		}

		document.forma.submit();		
	 }
	 
	 function SetValidar() {
		var SetIngresar = document.getElementById('IngresarTracking');
		var SetEditar = document.getElementById('EditarTracking');
		SetIngresar.style.visibility = "visible";
		SetEditar.style.visibility = "hidden";
		document.forma.Comment.value = "";
		document.forma.CD.value = "";
		document.forma.CT.value = "";
	 }

_editor_url = "Javascripts/";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);

if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }

/*
if (win_ie_ver >= 5.5) {
  document.write('<scr' + 'ipt src="' +_editor_url + 'editor.js"');
  document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { 
  document.write('<scr' + 'ipt src="' +_editor_url + 'editor2.js"');
  document.write(' language="Javascript1.2"></scr' + 'ipt>');  
}
*/

    function move() {
        window.location = '#';
        document.forma.style.display = "none";
        document.getElementById('myBar').style.display = "block";
        var elem = document.getElementById("myBar");
        var width = 10;
        var id = setInterval(frame, 65);
        function frame() {
            if (width >= 100) {
                clearInterval(id);
            } else {
                width++;
                elem.style.width = width + '%';
                elem.innerHTML = width * 1 + '%';
            }
        }
    }

    function SendingTracking() {
        var SendingID = document.getElementById("SendingID"); 
        SendingID.value = 'Enviando..';
        SendingID.style.display = 'inline';      
    }

</script>

<style>
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
</style>


<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<!-- <BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"> -->

<textarea id="SendingID" name="SendingID" rows=1 style='display:none;border:0px;color:navy;font-size:20px; font-family:Verdana' ></textarea>

	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<TR>
	<TD valign="top">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="AWBID" type=hidden value="<%=AWBID%>">
	<INPUT name="CID" type=hidden value="<%=ConsignerID%>">
	<INPUT name="AT" type=hidden value="<%=AWBType%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
		<TR><TD class=label align=left colspan="2"><%=Comment2%></TD></TR> 
		<TR><TD class=label align=right><b>Estado:</b></TD><TD class=label align=left>
		<input name="BLStatusName" type=hidden value="">
        <input name="NAgentID" type=hidden value="0">
        <input name="NClientID" type=hidden value="0">
        <input name="NShipperID" type=hidden value="0">
		<select name="BLStatus" id="Status de la Carga" class="label" onchange="SelectBLStatus()">
			<option value="-1">SELECCIONAR</option>
			<%=Estate%>
		</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Comentario:</b></TD><TD class=label align=left><Textarea name="Comment" id="Comentario" cols="40" rows="10"><%=Comment%></Textarea></TD></TR>
		<TR><TD class=label align=right><b>Fecha&nbsp;Creaci&oacute;n:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
		<TR><TD class=label align=right><b>C&oacute;digo:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
			<%if CountTableValues < 0 then%>
				<TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Ingresar&nbsp;&nbsp;" class=label></TD>
			<%else%>
				<TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(0)" value="&nbsp;&nbsp;Nuevo&nbsp;&nbsp;" class=label></TD>
				<TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
				<TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(3)" value="&nbsp;&nbsp;Eliminar&nbsp;&nbsp;" class=label></TD>
			<%end if%>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</FORM>
	</TABLE>

    

    <h3 class=label>Contactos a Notificar</h3>
    Agente : <%=agente_nom%>
    <table border=0 width="100%">
        <tr><th class=label style='background:silver'>Tipo</th><th class=label style='background:silver'>Nombre</th><th class=label style='background:silver'>Email</th></tr>
<%
    if arr_contact_count_tmp > -1 then
        for i=0 to arr_contact_count_tmp

            Select Case arr_contact_tmp(8,i)            
            Case "Agente","Consigneer","Shipper","Coloader","Notify"

            response.write ( "<tr>" )
            response.write ( "<td class=label style='border-bottom:1px solid gray'>" & arr_contact_tmp(8,i) & "</td>" )
			
            Select Case arr_contact_tmp(8,i)            
            Case "Agente"
            arr_contact_tmp(0,i) = "<a href=http://10.10.1.20/catalogo_admin/agentes-detalle.php?agente_id=" & arr_contact_tmp(12,i) & " target=_blank >" & arr_contact_tmp(0,i) & "</a>"
            
            Case "Consigneer","Shipper","Coloader","Notify"
            arr_contact_tmp(0,i) = "<a href=http://10.10.1.20/catalogo_admin/clientes-detalle.php?id_cliente=" & arr_contact_tmp(12,i) & " target=_blank >" & arr_contact_tmp(0,i) & "</a>"
            end Select
            
			response.write ( "<td class=label style='border-bottom:1px solid gray'>" & arr_contact_tmp(0,i) & "</td>" )
			response.write ( "<td class=label style='border-bottom:1px solid gray'>" & arr_contact_tmp(1,i) & "</td>" )            

            'response.write ( "<td><input type=text readonly name='email_cta[" & i & "]' value='" & arr_contact_tmp(2,i) & "'></td>" )            
            'if trim(arr_contact_tmp(2,i)) <> "" then
            '    response.write ( "<td><input type=checkbox name='email_send[" & i & "]' value='" & i & "'></td>" )
            'else
            '    response.write ( "<td></td>" )
            'end if

            response.write ( "</tr>" )

            end Select

        next 
    end if
%>
    </table>


	</TD>
	<TD>&nbsp;&nbsp;</TD>
	<TD valign="top">
	<iframe id="trackhistory" name="trackhistory" src="TrackingHistory.asp?AWBID=<%=AWBID%>&CID=<%=ConsignerID%>&AWBType=<%=AWBType%>&AWBNumber=<%=aList1Values(0,0)%>&pais=<%=aList1Values(5,0)%>&HAWBNumber=<%=aList1Values(1,0)%>" frameborder="0" framespacing="0" scrolling="auto" width="600" height="400">
	Tu browser no soporta esta funcionalidad, favor contactar a soporte.
	</iframe>
	</TD>
	</TR>
	</TABLE>
</BODY>
<script language="javascript1.2">
if (win_ie_ver >= 5.5) {
	editor_generate('Comment');  
} else { 
	//editor_generate('Comentario');  
}

selecciona('forma.BLStatus','<%=Val%>');


<% if Action = 98 then 'viene de TrackingHistyory %>
   
    SelectBLStatus();    

<% end if %>


</SCRIPT>

</HTML>