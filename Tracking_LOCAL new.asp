<meta http-equiv="Content-Type" content="text/html; charset=utf8">

<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
-->

<%

Checking "0|1|2"

    Function contactoXpais1(contactoxpais,email,pais)
        dim temp2, TestArray, dato, temp3                   
        temp3 = ""
        contactoXpais1 = ""
        if contactoxpais <> "" then
            contactoxpais = Replace(contactoxpais,"{","")
            contactoxpais = Replace(contactoxpais,"}","")
            TestArray = Split(contactoxpais,",")
            For Each dato In TestArray
                temp2 = Split(dato,":")
                'response.write("*(" & dato & ")(" & Replace(temp2(0),"""","") & ")(" & pais & ")<br>")
                if Replace(temp2(0),"""","") = pais then
                    temp3 = temp2(1)
                    temp3 = Replace(temp3,"""","")                            
                    'response.write(temp2(0) & " " & email & "<br>")
                end if               
            next 
        end if

        if temp3 = "" then
            temp3 = email              
        end if

        contactoXpais1 = temp3
    End Function

    Function in_array(element, arr)
        dim j
        For j=0 To Ubound(arr) 
            If Trim(arr(j)) = Trim(element) Then 
                in_array = True
                Exit Function
            Else 
                in_array = False
            End If  
        Next 
    End Function 



Sub SendNotification (BL, HBLNumber, Status, Countries, SQLQuery, AwbType, Comentario, ConsignerID, ShipperID  )

    Dim Conn, rs, i, BodySpanish, BodyEnglish, Body_b, eSendYes, Header, Logo, ubicacion, pais_origen_nombre, emailAnt, from, subject, BodyNoEnvio, result, webp_url, webp_tex, flag_img, eSendYesAge
  
    Dim headers, ContactoEmail, ContactoTel, local_, atentamente, Divisiones, eSendNo, EmailStr, CountListaMails, ListaMails

    OpenConn2 Conn

    'response.write SQLQuery
    CountListaMails = -1
    Set rs = Conn.Execute(SQLQuery)
    if Not rs.EOF then
        ListaMails = rs.GetRows
        CountListaMails = rs.RecordCount - 1
    end if
    CloseOBJ rs

    SQLQuery = "SELECT nombre_cliente FROM clientes WHERE id_cliente = '" & ConsignerID & "' and id_estatus in (1,2)"
    'response.write(SQLQuery & "<br>") 
    set rs = Conn.Execute(SQLQuery)
    if Not rs.EOF then
        ConsignerID = rs(0)
	end if
	CloseOBJ rs

    SQLQuery = "SELECT nombre_cliente FROM clientes WHERE id_cliente = '" & ShipperID & "' AND id_estatus in (1,2)"
    'response.write(SQLQuery & "2.(1)<br>") 
    set rs = Conn.Execute(SQLQuery)
    'response.write(SQLQuery & "2.(2)<br>") 
	if Not rs.EOF then
        ShipperID = rs(0)
	end if
	CloseOBJ rs

    CloseOBJ Conn


    select Case Countries
    Case "N1"
        Logo = "grh.bmp"
        webp_url = "www.aimargroup.com"
        webp_tex = "www.aimargroup.com"
        atentamente = "GRH"
    Case "BZLTF","GTLTF","SVLTF","HNLTF","NILTF","CRLTF","PALTF"
        Logo = "logo_latin_new.jpg"
        webp_url = "http://www.latinfreightneutral.com"
        webp_tex = "www.latinfreightneutral.com"
        atentamente = "Latin Freight"
        'flag_img = "<img src='www.latinfreightneutral.com/img/" & LCase(Countries) & "-flag.png' height=16>&nbsp;"
    Case else
        Logo = "aimargroup.jpg"
        webp_url = "http://www.aimargroup.com"
        webp_tex = "www.aimargroup.com"
        atentamente = "Aimar Group"
    end Select

    if AwbType = 1 then
        ubicacion = "EXPORT"
    else
        ubicacion = "IMPORT"
    end if
							
    'pais_origen_nombre = TranslateCountry (Countries)	
    pais_origen_nombre = TranslateCountry (Mid(Countries, 1, 2))	
    
    Logo = "<img src='" & webp_url & "/img/" & Logo & "' height=60>"
    flag_img = "<img src='" & webp_url & "/img/" & Mid(LCase(Countries), 1, 2) & "-flag.png' height=16>&nbsp;"    
    headers = Logo & flag_img & pais_origen_nombre & " " & ubicacion 


    if Request.Servervariables("REMOTE_ADDR") = "127.0.0.1" or Request.Servervariables("REMOTE_ADDR") = "::1" or Request.Servervariables("REMOTE_ADDR") = "localhost"  then
        local_ = 1
    else
        local_ = 0
    end if    
    
    ContactoEmail = ""
    ContactoTel = ""

    if CountListaMails > -1 then
        for i=0 to CountListaMails
            if ListaMails(8,i) = "Contacto" then
                ContactoEmail = contactoXpais1(ListaMails(11,i),ListaMails(1,i),Countries)
                'ContactoEmail = ListaMails(1,i)
                ContactoTel = ListaMails(2,i)
            end if
        next 
    end if



    subject = ""
    if local_ = 1 then
        subject = Request.Servervariables("REMOTE_ADDR")
    end if

    subject = subject & "Status Notification Air "

    if trim(HBLNumber) = "" then    
        subject = subject & "AWB:" & BL & " / " & ConsignerID
    else	
        subject = subject & "HAWB:" & HBLNumber & " / " & ConsignerID	    
    end if	



    '/////////////////////////////////////////////////CLIENTE LATINO ESPAÑOL//////////////////////////////////////////////////////////
	BodySpanish = headers & "<p>Estimado Cliente : </p><p>A continuaci&oacute;n le damos a conocer el status actual de su mercaderia amparada con la siguiente informaci&oacute;n : </p>"	
    if trim(HBLNumber) = "" then
        BodySpanish = BodySpanish & "<p><b>AWB : </b>" & BL & "<br>"
    else
	    BodySpanish = BodySpanish & "<p><b>HAWB : </b>" & HBLNumber & "<br>"
    end if
	BodySpanish = BodySpanish & "<b>Cliente : </b>" & ConsignerID & "<br><b>Shipper : </b>" & ShipperID & "<br><b>STATUS : </b>" & Status & "<br>"	
	BodySpanish = BodySpanish & "<b>Observaciones : </b><font color='green'>" & Comentario & "</font></p>"
	BodySpanish = BodySpanish & "<p style='text-align:justify'>Si necesita mayor informaci&oacute;n de su carga puede visitar nuestro tracking en la pagina web: <a href='" & webp_url & "'>" & webp_tex & "</a> o bien consultar con nuestro departamento de Servicio al Cliente " & ContactoEmail & " Telefono: " & ContactoTel & ".</p>"
	BodySpanish = BodySpanish & "<p>Para solicitar un usuario y password para acceso al tracking puede comunicarse con nuestro representante de Servicio al Cliente.</p>"
	BodySpanish = BodySpanish & "<p>Estamos para servirle,</p><p>Atentamente,</p>" & atentamente
	BodySpanish = BodySpanish & "<p><b>IMPORTANTE:<br>Favor no responder este email ya que fue enviado desde un sistema automaticamente y no tendr&aacute; respuesta desde esta direcci&oacute;n de correo.</b></p>"
	BodySpanish = BodySpanish & "<font color='white'>" & Session("Login") & " " & Request.Servervariables("REMOTE_ADDR") & "</font>"


    '/////////////////////////////////////////////////AGENTES INGLES//////////////////////////////////////////////////////////
	BodyEnglish = headers & "<p>Dear Agent : </p><p>Here we present the current status of your shipment with the following information : </p>"
    if trim(HBLNumber) = "" then
        BodyEnglish = BodyEnglish & "<p><b>AWB : </b>" & BL & "<br>"        
    else
	    BodyEnglish = BodyEnglish & "<p><b>HAWB : </b>" & HBLNumber & "<br>"        
    end if	
	BodyEnglish = BodyEnglish & "<b>Consignee : </b>" & ConsignerID & "<br><b>Shipper : </b>" & ShipperID & "<br><b>STATUS : </b><font color=green>" & Status & "</font><br></p>"
	BodyEnglish = BodyEnglish & "<p style='text-align:justify'>If you need any additional information please visit our tracking on web page <a href='" & webp_url & "'>" & webp_tex & "</a> or you can also contact our Customer Service Department e-mail: " & ContactoEmail & " Phone: " & Replace(ContactoTel,"Supervisora de Servicio al cliente"," Customer Service Manager") & ".</p>"
	BodyEnglish = BodyEnglish & "<p>To request a username and password to access to the tracking please contact a customer service representative.</p>"
	BodyEnglish = BodyEnglish & "<p>Cordially,</p>" & atentamente
	BodyEnglish = BodyEnglish & "<p> <b>IMPORTANT:<br>Please do not reply to this email, it is sent from an automated system, there will be no response from this address. For assistance contact customer service department.</b></p>"
	BodyEnglish = BodyEnglish & "<font color='white'>" & Session("Login") & " " & Request.Servervariables("REMOTE_ADDR") & "</font>"

    BodyEnglish = BodyEnglish & "<br>"
    BodyEnglish = BodyEnglish & "<hr>"
    BodyEnglish = BodyEnglish & "<br>"
    
	BodyEnglish = BodyEnglish & "<p>Estimado Agente : </p><p>A continuaci&oacute;n le damos a conocer el status actual de su mercaderia amparada con la siguiente informaci&oacute;n : </p>"	
    if trim(HBLNumber) = "" then
        BodyEnglish = BodyEnglish & "<p><b>AWB : </b>" & BL & "<br>"
    else
	    BodyEnglish = BodyEnglish & "<p><b>HAWB : </b>" & HBLNumber & "<br>"
    end if
	BodyEnglish = BodyEnglish & "<b>Cliente : </b>" & ConsignerID & "<br><b>Shipper : </b>" & ShipperID & "<br><b>STATUS : </b><font color=green>" & Status & "</font><br>"		
	BodyEnglish = BodyEnglish & "<p style='text-align:justify'>Si necesita mayor informaci&oacute;n de su carga puede visitar nuestro tracking en la pagina web: <a href='" & webp_url & "'>" & webp_tex & "</a> o bien consultar con nuestro departamento de Servicio al Cliente " & ContactoEmail & " Telefono: " & ContactoTel & ".</p>"
	BodyEnglish = BodyEnglish & "<p>Para solicitar un usuario y password para acceso al tracking puede comunicarse con nuestro representante de Servicio al Cliente.</p>"
	BodyEnglish = BodyEnglish & "<p>Estamos para servirle,</p><p>Atentamente,</p>" & atentamente
	BodyEnglish = BodyEnglish & "<p><b>IMPORTANTE:<br>Favor no responder este email ya que fue enviado desde un sistema automaticamente y no tendr&aacute; respuesta desde esta direcci&oacute;n de correo.</b></p>"
	BodyEnglish = BodyEnglish & "<font color='white'>" & Session("Login") & "</font>"


    

    '/////////////////////////////////////////////////CLIENTE INGLES//////////////////////////////////////////////////////////
	Body_b = headers & "<p>Dear Consignee : </p><p>Here we present the current status of your shipment with the following information : </p>"
    if trim(HBLNumber) = "" then
        Body_b = Body_b & "<p><b>AWB : </b>" & BL & "<br>"        
    else
	    Body_b = Body_b & "<p><b>HAWB : </b>" & HBLNumber & "<br>"        
    end if	
	Body_b = Body_b & "<b>Consignee : </b>" & ConsignerID & "<br><b>Shipper : </b>" & ShipperID & "<br><b>STATUS : </b><font color=green>" & Status & "</font><br></p>"
    Body_b = Body_b & "<b>Comments : </b><font color='green'>" & Comentario & "</font></p>"
	Body_b = Body_b & "<p style='text-align:justify'>If you need any additional information please visit our tracking on web page <a href='" & webp_url & "'>" & webp_tex & "</a> or you can also contact our Customer Service Department e-mail: " & ContactoEmail & " Phone: " & Replace(ContactoTel,"Supervisora de Servicio al cliente"," Customer Service Manager") & ".</p>"
	Body_b = Body_b & "<p>To request a username and password to access to the tracking please contact a customer service representative.</p>"
	Body_b = Body_b & "<p>Cordially,</p>" & atentamente
	Body_b = Body_b & "<p> <b>IMPORTANT:<br>Please do not reply to this email, it is sent from an automated system, there will be no response from this address. For assistance contact customer service department.</b></p>"
	Body_b = Body_b & "<font color='white'>" & Session("Login") & " " & Request.Servervariables("REMOTE_ADDR") & "</font>"


    

		
    from = "tracking@aimargroup.com"    
    
    if CountListaMails > -1 then        

            eSendYes = ""
            eSendYesAge = "" 
            emailAnt  = ""   
            BodyNoEnvio = ""
			eSendNo = ""
            Divisiones = ""

            Dim arr1()            
            redim arr1(CountListaMails)
            'for i=0 to CountListaMails
            '    arr1(i) = ""
            'next

            for i=0 to CountListaMails
                'response.write( ListaMails(11,i) & " " & ListaMails(1,i) & " " & Countries & "<br>" )
                EmailStr = contactoXpais1(ListaMails(11,i),ListaMails(1,i),Countries)
                'EmailStr = ListaMails(1,i)

                result = -1
                'Se verifica si tiene @ para que se interprete como correo
                if InStr(1,EmailStr,"@")>0 then
                                        
                    'if UCase(EmailStr) = UCase(emailAnt) then
                    If in_array(EmailStr, arr1 ) Then
                        result = -3 'email repetido
                        'response.write "Ya existe " & EmailStr & "<br>"
                    else

                        arr1(i) = EmailStr
                        'response.write "nuevo " & arr1(i) & "<br>"
                        'response.write ( "(" & Mid(Countries, 1, 2) & ")<br>" )
                        if ListaMails(8,i) <> "Desarrollo" then 'Copia
                            if ListaMails(9,i) = "Si" then 'Copia
                                '//////////////////////////////////////////////////////////////////////////////////////////////////////
                                if local_ = 0 then
                                    if ListaMails(8,i) = "Agente" then                            
                                        if InStr(1, EmailStr, "aimargroup") = 0 then
    'comentado temporalmente solicitado por Ticket#2015062404000623 — Fwd: Re: NUEVO SISTEMA DE AUTO NOTIFICACIONES AIMAR Linea 6
                                            'SendMails2 BodyEnglish, EmailStr, subject, from, result 
                                            'eSendYesAge = eSendYesAge & " " & EmailStr & "<br>"                                    
                                        end if
                                    else
                                        'if Countries <> "GT" then
                                            if Left(Countries, 2) = "BZ" then
                                                SendMails2 Body_b, EmailStr, subject, from, result 
                                            else
                                                SendMails2 BodySpanish, EmailStr, subject, from, result 
                                            end if
                                        'end if
                                    end if
                                else
                                    result = -2
                                end if
                            end if
                            
                            'response.write ( EmailStr & "<br>" )

                            if result = -1 then                            
                                eSendNo = eSendNo & " " & EmailStr & "<br>" 
                            else                            
                                eSendYes = eSendYes & " " & EmailStr & " " & result & "<br>" 
                            end if
                            emailAnt = UCase(EmailStr)
                        end if


                        Select Case ListaMails(8,i)            
                        Case "Soporte","Contacto","Monitor"
                            Divisiones = Divisiones & "No." & i & "<br>"
                            Divisiones = Divisiones & "TipoPersona : " & ListaMails(8,i) & "<br>"
                            Divisiones = Divisiones & "Nombre : " & ListaMails(0,i) & "<br>"
                            Divisiones = Divisiones & "Email : " & EmailStr & " (" & ListaMails(1,i) & ")<br>"
                            Divisiones = Divisiones & "Telefono : " & ListaMails(2,i) & "<br>"
                            Divisiones = Divisiones & "Pais : " & ListaMails(3,i) & "<br>"
                            Divisiones = Divisiones & "Copia : " & ListaMails(9,i) & "<br>"
                            Divisiones = Divisiones & "Rechazo : " & ListaMails(10,i) & "<br>"
                            Divisiones = Divisiones & "Send : " & result & "<br><br>"
                        end Select


                    end if
                end if

                if result = -1 then
                    Select Case ListaMails(8,i)            
                    Case "Agente","Consigneer","Shipper","Coloader","Notify"
				        BodyNoEnvio = BodyNoEnvio & "<font color=red>IMPORTANTE:</font> No se pudo enviar estatus al "
				        'if ListaMails(3,i) = 2 then BodyNoEnvio = BodyNoEnvio & "Contacto del "
				        BodyNoEnvio = BodyNoEnvio & ListaMails(8,i) & "<br>"
                        BodyNoEnvio = BodyNoEnvio & "<b>" & ListaMails(12,i) & " - " & ListaMails(0,i) & "</b> E-mail : " & EmailStr & "<br>"
				        if trim(EmailStr) = "" then
					        BodyNoEnvio = BodyNoEnvio & "no tiene cuenta de correo, favor de revisar y actualizar<br><br>"
				        else
					        BodyNoEnvio = BodyNoEnvio & "tiene error en cuenta de correo, favor de revisar y actualizar<br><br>"
                        end if                    
                    end Select
                end if
            Next

            if eSendYes <> "" then
                eSendYes = "<br><br><table border=0 width=99% align=left><tr><td>Correos Enviados:<br>" & eSendYes & "</td></tr></table>"
            end if 

            if eSendNo <> "" then
                eSendNo = "<br><br><table border=0 width=99% align=left><tr><td>Correos No Enviados:<br>" & eSendNo & "</td></tr></table>"
            end if

		    if eSendYesAge <> "" then
		        eSendYesAge = "<br><br><table border=0 width=99% align=left><tr><td>Correos de Agentes Enviados:<br>" & eSendYesAge & "</td></tr></table>"            
		    end if


		    'response.write("DESARROLLO<BR>")                
		    'A DESARROLLO 				
            'if Countries <> "GT" then		
            for i=0 to CountListaMails			
				'response.write (  ListaMails(8,i) & " " &  ListaMails(9,i) & "<br>" )			
                if ListaMails(8,i) = "Desarrollo" and ListaMails(9,i) = "Si" then 'Copia
                    'response.write ( "*" & ListaMails(1,i) & "<br>" )
                    SendMails2 BodySpanish & eSendYes & eSendYesAge & Divisiones, ListaMails(1,i), subject, from, result 
                end if
            Next
            'end if

		    'response.write("RECHAZOS<BR>")
		    'RECHAZOS
		    if trim(BodyNoEnvio) <> "" then
				BodyNoEnvio = BodyNoEnvio & "A continuacion el mensaje original: <br><br>"
				BodyNoEnvio = BodyNoEnvio & "<hr><b>SUBJECT : " & subject & "</b><br><br><hr>" 
                                        
                if Mid(Countries, 1, 2) = "BZ" then                                                
                    BodyNoEnvio = BodyNoEnvio & Body_b
                else
                    BodyNoEnvio = BodyNoEnvio & BodySpanish
                end if      
                
                'if Countries <> "GT" then          	

		        emailAnt  = ""
                for i=0 to CountListaMails
                    if InStr(1,ListaMails(1,i),"@")>0 then
                        if UCase(ListaMails(1,i)) <> UCase(emailAnt) then
                            if ListaMails(10,i) = "Si" then 'Rechazo
                                if ListaMails(8,i) = "Desarrollo" then 
                                    SendMails2 BodyNoEnvio, ListaMails(1,i), "NO SE ENVIO:" & subject, from, result                                 
                                else
                                    if local_ = 0 then
                                        SendMails2 BodyNoEnvio, ListaMails(1,i), "NO SE ENVIO:" & subject, from, result                                 
                                    end if
                                end if
                            end if
                            emailAnt = ListaMails(1,i)
                        end if
                    end if
                Next

                'end if

			end if



    end if

    Set ListaMails = Nothing

End Sub




if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	AWBID = aTableValues(3, 0)
	ConsignerID = aTableValues(4, 0)
	Comment = aTableValues(5, 0)
	Val = aTableValues(6, 0)
else
	AWBID=CheckNum(request("AWBID"))
	ConsignerID=CheckNum(request("CID"))
	CreatedDate = ""
	CreatedTime = ""
end if

Set aTableValues = Nothing

	CountList1Values = -1 
	OpenConn Conn

        SQLQuery = "select AWBNumber, HAWBNumber, ifnull(ConsignerID,0), ifnull(ShipperID,0), ifnull(AgentID,0), Countries, ifnull(id_coloader,0), ifnull(id_cliente_order,0), 1 as no from "
		if AWBType = 1 then
		    SQLQuery = SQLQuery & "Awb"
		else 
            SQLQuery = SQLQuery & "Awbi"
		end if
        'esto debido a que hubo un caso que no trajo nada 
        SQLQuery = SQLQuery & " where AWBID=" & AWBID & " union select 'n/a','n/a',0,0,0,'',0,0,2 as no order by no asc"
		'response.write(SQLQuery & "<br>")
		set rs = Conn.Execute(SQLQuery)
		if Not rs.EOF then
		    aList1Values = rs.GetRows
			CountList1Values = rs.RecordCount - 1
		end if        
	CloseOBJs rs, Conn





Dim NotifyAgentID, NotifyClientID, NotifyShipperID, Header, selected, agente_nom
	
	OpenConn2 Conn
		'Obteniendo el listado de Status Terrestre Local
		SQLQuery = "select id, estatus, notificar_agente, notificar_cliente, notificar_shipper, publico from aimartrackings where air=1 and activo=1 "
                
        if AwbType = 1 then
        SQLQuery = SQLQuery & " and import = 0 order by estatus"
        else
        SQLQuery = SQLQuery & " and import = 1 order by estatus"
        end if

        set rs = Conn.Execute(SQLQuery)

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


    agente_nom = ""
    if CheckNum(aList1Values(4,0)) > 0 then
        SQLQuery = "SELECT agente FROM agentes WHERE agente_id = " & CheckNum(aList1Values(4,0)) & " AND activo = 't' "		
        set rs = Conn.Execute(SQLQuery)
        if Not rs.EOF then		    
	        agente_nom = rs(0)
		end if        
	    CloseOBJ rs
    end if


    Function EmailsExternos(sort,id,titulo,pais)  
        dim SQLQuery1
        if titulo = "Agente" then           
'2015-08-07
'SQLQuery1 = "SELECT agente_id as cod, agente as nombre, trim(correo) as email, 1 as nivel, '" & titulo & "' as tipo FROM agentes WHERE agente_id = " & id & " AND activo = 't' " &_
'"UNION " & _
'"SELECT agente_id as cod, nombres as nombre, trim(email) as email, 2 as nivel, '" & titulo & "' as tipo FROM agentes_contactos WHERE id_pais LIKE '%" & pais & "%' AND agente_id = " & id & " AND activo = 't' "
'SQLQuery1 = "SELECT agente_id as cod, nombres as nombre, trim(email) as email, 2 as nivel, '" & titulo & "' as tipo FROM agentes_contactos WHERE  agente_id = " & id & " AND activo = 't'"
'2015-11-18
'EmailsExternos = "SELECT agente_id as cod, nombres as nombre, trim(email) as email, 2 as nivel, '" & titulo & "' as tipo FROM agentes_contactos WHERE  agente_id = " & id & " AND activo = 't'"
'id_pais = '" & pais & "' AND 2015-10-18 para pruebas lo quite

'2016-03-07 hoy aun se usara este catalogo, hasta que se publique el nuevo catalogo_admin (yii)
EmailsExternos = "SELECT nombres as nombre, trim(email) as email, telefono, '' as pais, '' as area, '' as impexp, '' as carga, '' as tranship, '" & titulo & "' as tipo_persona, 'Si' as copia, '' as rechazo, '' as contactoxpais, agente_id as id_catalogo, id_contacto, " & sort & " as sort FROM agentes_contactos WHERE agente_id = " & id & " AND activo = 't'"
        
        else            

'2015-08-07
'SQLQuery1 = "SELECT id_cliente as cod, nombre_cliente as nombre, trim(email) as email, 1 as nivel, '" & titulo & "' as tipo FROM clientes WHERE id_cliente = " & id & " AND id_estatus = 1 " & _
'"UNION " & _
'"SELECT id_cliente as cod, nombres as nombre, trim(email) as email, 2 as nivel, '" & titulo & "' as tipo FROM contactos WHERE id_cliente = " & id & " AND activo = 't' "

'EmailsExternos = "SELECT id_cliente as cod, nombres as nombre, trim(email) as email, 2 as nivel, '" & titulo & "' as tipo FROM contactos WHERE id_cliente = " & id & " AND activo = 't' "

EmailsExternos = "SELECT nombres as nombre, trim(email) as email, '' as telefono, '' as pais, '' as area, '' as impexp, '' as carga, '' as tranship, '" & titulo & "' as tipo_persona, 'Si' as copia, '' as rechazo, '' as contactoxpais, id_cliente as id_catalogo, contacto_id as id_contacto, " & sort & " as sort FROM contactos WHERE id_cliente = " & id & " AND activo = 't' "

        end if	  
        
    End Function
    

    SQLQuery = ""


    dim arr_contact_count_tmp, arr_contact_tmp, CoLoaderID, IsConsignerID 
    
    arr_contact_count_tmp = -1

    'response.write ( "(" & Action & ")<br>" )

    if Action = 1 or Action = 2 or Action = 99 then 'desde el select de estatus para display de los contactos correspondientes
    
        AgentID = CheckNum(Request.Form("NAgentID"))
        IsConsignerID = CheckNum(Request.Form("NClientID"))
        ShipperID = CheckNum(Request.Form("NShipperID"))

        if aList1Values(6,0) > 0 then
            CoLoaderID = 1
        else
            CoLoaderID = 0
        end if





        '2016-03-07 hoy aun no se usa este codigo hasta que catalogos este publicado
        'if AgentID = 1 then  
        '    SQLQuery = SQLQuery & " OR (catalogo = 'AGENTE' AND id_catalogo = '" & aList1Values(4,0) & "') "
        'end if
        
        'temporalmente los clientes de tabla contactos 2016-02-25
        'if CoLoaderID = 1 and (ShipperID = 1 or ConsignatarioID = 1) then            
        '    SQLQuery = SQLQuery & " OR (catalogo = 'CLIENTE' AND id_catalogo = '" & aList1Values(6,0) & "') "
        'else
        '    if ShipperID = 1 then
        '        SQLQuery = SQLQuery & " OR (catalogo = 'CLIENTE' AND id_catalogo = '" & aList1Values(3,0) & "') "
        '    end if        
        '    if ConsignatarioID = 1 then
    	'        SQLQuery = SQLQuery & " OR (catalogo = 'CLIENTE' AND id_catalogo = '" & aList1Values(2,0) & "') "
        '    end if
        'end if

        SQLQuery = "SELECT nombre, email, telefono, pais, area, impexp, carga, tranship, tipo_persona, copia, rechazo, contactoxpais, id_catalogo, id_contacto, case when tipo_persona = 'Desarrollo' then 10 when tipo_persona = 'Contacto' then 9 when tipo_persona = 'Soporte' then 8 when tipo_persona = 'Shipper' then 6 when tipo_persona = 'Coloader' then 5 when tipo_persona = 'Consigneer' then 4 when tipo_persona = 'Agente' then 3 else 0 end as sort FROM contactos_divisiones " & _ 
        "WHERE status = 'Activo' AND area ILIKE '%Aereo%' AND " & _ 
        "( (catalogo = 'USUARIO' AND pais ILIKE '%""" & aList1Values(5,0) & """%') " & SQLQuery & ") "

        if AwbType = "1" then 
            SQLQuery = SQLQuery & " AND impexp ILIKE '%Export%'"
        else
            SQLQuery = SQLQuery & " AND impexp ILIKE '%Import%'" 
        end if

		'response.write "(" & aList1Values(7,0) & ")"

        if aList1Values(7,0) > "0" then

            if SQLQuery <> "" then
                SQLQuery = SQLQuery & " UNION "
            end if
            SQLQuery = SQLQuery & EmailsExternos(7,CheckNum(aList1Values(7,0)),"Notify","")
        
            if ShipperID = 1 then
                if SQLQuery <> "" then
                    SQLQuery = SQLQuery & " UNION "
                end if
                SQLQuery = SQLQuery & EmailsExternos(6,CheckNum(aList1Values(3,0)),"Shipper","")
            end if

        else

            '2016-03-07 hoy aun sigue este codigo
            if AgentID = 1 then                        
                if SQLQuery <> "" then
                    SQLQuery = SQLQuery & " UNION "
                end if
                SQLQuery = SQLQuery & EmailsExternos(3,CheckNum(aList1Values(4,0)),"Agente",aList1Values(5,0))
            end if
        
            if CoLoaderID = 1 and (ShipperID = 1 or IsConsignerID = 1) then            
                if SQLQuery <> "" then
                    SQLQuery = SQLQuery & " UNION "
                end if
                SQLQuery = SQLQuery & EmailsExternos(5,CheckNum(aList1Values(6,0)),"Coloader","")
            else
                if ShipperID = 1 then
                    if SQLQuery <> "" then
                        SQLQuery = SQLQuery & " UNION "
                    end if
                    SQLQuery = SQLQuery & EmailsExternos(6,CheckNum(aList1Values(3,0)),"Shipper","")
                end if

		        if IsConsignerID = 1 then
		            if SQLQuery <> "" then
                        SQLQuery = SQLQuery & " UNION "
                    end if
                    SQLQuery = SQLQuery & EmailsExternos(4,CheckNum(aList1Values(2,0)),"Consigneer","")
                end if
            end if
        end if

		if SQLQuery <> "" then
		    SQLQuery = "select * from (" & SQLQuery & ") x order by sort" 
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
    

    'response.write("(" & aList1Values(0,0) & ")(" & aList1Values(1,0) & ")(" & aList1Values(5,0) & ")(" & aList1Values(2,0) & ")(" & aList1Values(3,0)  & ")<br>" )

    'Enviando Notificaciones por Mail para Agente, Cliente, Shipper cuando se ingresa o actualiza informacion
    Select Case Action
    case 1,2
		'para pruebas se comento el if        
        'SE AGREGO GTLTF 2015-11-18
        'OR aList1Values(5,0) = "GTLTF" 
		'if aList1Values(5,0) = "BZ" OR aList1Values(5,0) = "GT" OR aList1Values(5,0) = "SV" OR aList1Values(5,0) = "HN" OR aList1Values(5,0) = "NI" OR aList1Values(5,0) = "CR" OR aList1Values(5,0) = "PA" then
            if SQLQuery <> "" then '2015-09-23 para que no de error
                SendNotification aList1Values(0,0), aList1Values(1,0), Request.Form("BLStatusName"), aList1Values(5,0), SQLQuery, AwbType, Request.Form("Comment"), aList1Values(2,0), aList1Values(3,0) 
            end if
        'end if
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
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
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
</script>

<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
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
		<TR><TD class=label align=right><b>Fecha&nbsp;Creación:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
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
	<iframe id="trackhistory" name="trackhistory" src="TrackingHistory.asp?AWBID=<%=AWBID%>&CID=<%=ConsignerID%>&AWBType=<%=AWBType%>&AWBNumber=<%=aList1Values(0,0)%>&pais=<%=aList1Values(5,0)%>" frameborder="0" framespacing="0" scrolling="auto" width="600" height="400">
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