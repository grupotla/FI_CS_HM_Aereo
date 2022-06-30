<%
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





    
Function SendNotification(Status, SQLQuery, AwbType, Comentario, aList1Valuesi, SendShipper, Mode)

    Dim BLi, HBLNumberi, Countriesi, ConsignerIDi, ShipperIDi, Coloaderi, Notifyi 

    dim SendCoLoader, SendNotify

    if CheckNum(aList1Valuesi(6,0)) > 0 then 'coloader
        SendCoLoader = 1
    else
        SendCoLoader = 0
    end if

    if CheckNum(aList1Valuesi(7,0)) > 0 then 'notify
        SendNotify = 1
    else
        SendNotify = 0
    end if


    'response.write "" & Status & "<br><br>" & SQLQuery & "<br><br>" & AwbType & "<br><br>" & Comentario & "<br><br>" & SendShipper & "<br><br>"

    'response.write aList1Valuesi(0,0) & "<br><br>" & aList1Valuesi(1,0) & "<br><br>" & aList1Valuesi(2,0) & "<br><br>" & aList1Valuesi(3,0) & "<br><br>" & aList1Valuesi(4,0) & "<br><br>" & aList1Valuesi(5,0) & "<br><br>" & aList1Valuesi(6,0) & "<br><br>" & aList1Valuesi(7,0) & "<br><br>" & aList1Valuesi(8,0) & "<br><br>" & aList1Valuesi(9,0) & "<br><br>"

    BLi          = aList1Valuesi(0,0) 
    HBLNumberi   = aList1Valuesi(1,0) 
    Countriesi   = aList1Valuesi(5,0) 
    ConsignerIDi = aList1Valuesi(2,0) 
    ShipperIDi   = aList1Valuesi(3,0) 
    Coloaderi    = aList1Valuesi(6,0) 
    Notifyi      = aList1Valuesi(7,0) 

    Dim Conn, rs, i, BodySpanish, BodyEnglish, Body_b, eSendYes, Header, ubicacion, emailAnt, subject, BodyNoEnvio, result, flag_img, eSendYesAge  
    Dim headers, ContactoEmail, ContactoTel, local_, Divisiones, eSendNo, EmailStr, CountListaMails, ListaMails
    'atentamente, Logo, webp_url, webp_tex, iFromAddress, pais_origen_nombre

    OpenConn2 Conn
    
    'response.write "select routing, order_no, no_embarque FROM routings WHERE id_routing = '" & aList1Valuesi(9,0) & "'" & "<br><br>"

    dim routing, order_no, no_embarque
    Set rs = Conn.Execute("select routing, order_no, no_embarque FROM routings WHERE id_routing = '" & aList1Valuesi(9,0) & "'")
    if Not rs.EOF then
        ListaMails = rs.GetRows        
        routing = ListaMails(0,0)
        order_no = ListaMails(1,0)
        no_embarque = ListaMails(2,0)
    end if
    CloseOBJ rs



    'response.write SQLQuery
    CountListaMails = -1
    Set rs = Conn.Execute(SQLQuery)
    if Not rs.EOF then
        ListaMails = rs.GetRows
        CountListaMails = rs.RecordCount - 1
    end if
    CloseOBJ rs

    SQLQuery = "SELECT id_cliente, nombre_cliente FROM clientes WHERE id_cliente IN (" & ConsignerIDi & "," & ShipperIDi & "," & Coloaderi & "," & Notifyi & ") and id_estatus in (1,2)"
	'response.write(SQLQuery & "<br>") 
    Set rs = Conn.Execute(SQLQuery)	
	if Not rs.EOF then
		do while Not rs.EOF

            if CheckNum(ConsignerIDi) = CheckNum(rs(0)) then
                ConsignerIDi = rs(1)
            end if

            if CheckNum(ShipperIDi) = CheckNum(rs(0)) then
                ShipperIDi = rs(1)
            end if

            if CheckNum(Coloaderi) = CheckNum(rs(0)) then
                Coloaderi = rs(1)
            end if

            if CheckNum(Notifyi) = CheckNum(rs(0)) then
                Notifyi = rs(1)
            end if

            'response.write(rs(0) & " " & rs(1) & "<br>") 
            rs.MoveNext
        loop
    end if
    CloseOBJ rs

    'SQLQuery = "SELECT nombre_cliente FROM clientes WHERE id_cliente = '" & ConsignerIDi & "' and id_estatus in (1,2)"
    'response.write(SQLQuery & "<br>") 
    'set rs = Conn.Execute(SQLQuery)
    'if Not rs.EOF then
    '   ConsignerIDi = rs(0)
	'end if
	'CloseOBJ rs

    'SQLQuery = "SELECT nombre_cliente FROM clientes WHERE id_cliente = '" & ShipperIDi & "' AND id_estatus in (1,2)"    
    'set rs = Conn.Execute(SQLQuery)    
	'if Not rs.EOF then
    '    ShipperIDi = rs(0)
	'end if
	'CloseOBJ rs

    CloseOBJ Conn
    
    if 1 = 1 then

        'Logo = Iif(isNull(aTableValues5(20,0)),"","<img src='data:image/jpeg;base64," & aTableValues5(20,0) & "'>")            
        'webp_url = Iif(isNull(aTableValues5(18,0)),"",aTableValues5(18,0))    
        'atentamente = Iif(isNull(aTableValues5(19,0)),"",aTableValues5(19,0))    

        'select Case Countriesi
        'Case "N1"
        '    Logo = "grh.bmp"
        '    webp_url = "www.aimargroup.com"
        '    webp_tex = "www.aimargroup.com"
        '    atentamente = "GRH"
        'Case "BZLTF","GTLTF","SVLTF","HNLTF","NILTF","CRLTF","PALTF"
        '    Logo = "logo_latin_new.jpg"
        '    webp_url = "http://www.latinfreightneutral.com"
        '    webp_tex = "www.latinfreightneutral.com"
        '    atentamente = "Latin Freight"
        '    'flag_img = "<img src='www.latinfreightneutral.com/img/" & LCase(Countries) & "-flag.png' height=16>&nbsp;"
        'Case else
        '    Logo = "aimargroup.jpg"
        '    webp_url = "http://www.aimargroup.com"
        '    webp_tex = "www.aimargroup.com"
        '    atentamente = "Aimar Group"
        'end Select

        if AwbType = 1 or AwbType = "1" then
            ubicacion = "EXPORT"
        else
            ubicacion = "IMPORT" '2
        end if
							
        'pais_origen_nombre = TranslateCountry (Countries)	
        'pais_origen_nombre = TranslateCountry (Mid(Countriesi, 1, 2))	
    
        'Logo = "<img src='" & webp_url & "/img/" & Logo & "' height=60>"
        'flag_img = "<img src='" & webp_url & "/img/" & Mid(LCase(Countriesi), 1, 2) & "-flag.png' height=16>&nbsp;"    
        'headers = Logo & flag_img & pais_origen_nombre & " AIR " & ubicacion 
        headers = "<img src='data:image/jpeg;base64,#*logo*#'>#*nombre_pais*# AIR " & ubicacion 

        if Request.Servervariables("REMOTE_ADDR") = "127.0.0.1" or Request.Servervariables("REMOTE_ADDR") = "::1" or Request.Servervariables("REMOTE_ADDR") = "localhost"  then
            local_ = 1
        else
            local_ = 0

            'esto se comento 2018-08-02 cuando ya esta enviando emails corractamente
            'if Mode = 1 then 'Mode es 1 cuando viene de php
            '    local_ = 1 'forzar a que no envie a nadie mas solo a mi correo
            'else                            
            '    local_ = 0 'Mode es 0 cuando viene de rastreo
            'end if
        end if    
        'local_ = 1
        
        if local_ <> 0 then
            response.write "LOCALHOST :: (" & local_ & ")<br>"
        end if
    
        ContactoEmail = ""
        ContactoTel = ""

        if CountListaMails > -1 then
            for i=0 to CountListaMails
                if ListaMails(8,i) = "Contacto" then
                    ContactoEmail = contactoXpais1(ListaMails(11,i),ListaMails(1,i),Countriesi)
                    'ContactoEmail = ListaMails(1,i)
                    ContactoTel = ListaMails(2,i)
                end if
            next 
        end if



        if ContactoEmail = "" or ContactoTel = "" then        
            response.write "<font color=red> Se requiere informacion del contacto principal (" & Countriesi & ") (" & ContactoEmail & ")(" & ContactoTel & "). </font><br>"
        end if


        subject = ""
        if local_ = 1 then
            subject = Request.Servervariables("REMOTE_ADDR")
        end if

        subject = subject & "Status Notification "

        if order_no <> "" then
            subject = subject & "PO " & order_no & " / "
        end if

        if routing <> "" then
            subject = subject & "RO " & routing & " / "
        end if

        'no_embarque = ListaMails(2,0)

        subject = subject & "S: " & ShipperIDi & " / C: " & ConsignerIDi & " / "

        if trim(HBLNumberi) = "" then    
            subject = subject & "AWB: " & BLi
        else	
            subject = subject & "HAWB: " & HBLNumberi
        end if	



        '/////////////////////////////////////////////////CLIENTE LATINO ESPAÑOL//////////////////////////////////////////////////////////
	    BodySpanish = headers & "<p>Estimado Cliente : </p><p>A continuaci&oacute;n le damos a conocer el status actual de su mercaderia amparada con la siguiente informaci&oacute;n : </p>"	
        if trim(HBLNumberi) = "" then
            BodySpanish = BodySpanish & "<p><b>AWB : </b>" & BLi & "<br>"
        else
	        BodySpanish = BodySpanish & "<p><b>HAWB : </b>" & HBLNumberi & "<br>"
        end if
	    BodySpanish = BodySpanish & "<b>Cliente : </b>" & ConsignerIDi & "<br><b>Shipper : </b>" & ShipperIDi & "<br>"

        'response.write "**(" & SendConsigner & ")(" & SendShipper & ")(" & SendCoLoader & ")(" & SendNotify & ")" 

        if SendCoLoader = 1 then
	        BodySpanish = BodySpanish & "<b>Coloader : </b>" & Coloaderi & "<br>"
        end if

        if SendCoLoader = 1 or SendShipper = 1 then
            if SendNotify = 1 then 'id_cliente_order notify
                BodySpanish = BodySpanish & "<b>Notify Party : </b>" & Notifyi & "<br>"        
            end  if 
        end if 

        if right(Countriesi,3) = "TLA" then

        BodySpanish = BodySpanish & "<b>STATUS : </b>" & Status & "<br>"	
	    BodySpanish = BodySpanish & "<b>Observaciones : </b><font color='green'>" & Comentario & "</font></p>"
	    BodySpanish = BodySpanish & "<p style='text-align:justify'>Si necesita mayor informaci&oacute;n de su carga puede consultar nuestro departamento de Operaciones <a href='mailto:" & LCase(ContactoEmail) & "'>" & ContactoEmail & "</a> o al tel&eacute;fono: " & ContactoTel & ".</p>"
	    BodySpanish = BodySpanish & "<p>Estamos para servirle,</p><p>Atentamente,</p>#*firma*#"
	    BodySpanish = BodySpanish & "<p><b>IMPORTANTE:<br>Favor no responder este email ya que fue enviado desde un sistema automaticamente y no tendr&aacute; respuesta desde esta direcci&oacute;n de correo.</b></p>"
	    BodySpanish = BodySpanish & "<font color='white'>" & Session("Login") & " " & Request.Servervariables("REMOTE_ADDR") & "</font>"
        
        else

        BodySpanish = BodySpanish & "<b>STATUS : </b>" & Status & "<br>"	
	    BodySpanish = BodySpanish & "<b>Observaciones : </b><font color='green'>" & Comentario & "</font></p>"
	    BodySpanish = BodySpanish & "<p style='text-align:justify'>Si necesita mayor informaci&oacute;n de su carga puede visitar nuestro tracking en la pagina web: <a href='http://#*home_page*#'>#*home_page*#</a> o bien consultar con nuestro departamento de Servicio al Cliente <a href='mailto:" & LCase(ContactoEmail) & "'>" & ContactoEmail & "</a> Telefono: " & ContactoTel & ".</p>"
	    BodySpanish = BodySpanish & "<p>Para solicitar un usuario y password para acceso al tracking puede comunicarse con nuestro representante de Servicio al Cliente.</p>"
	    BodySpanish = BodySpanish & "<p>Estamos para servirle,</p><p>Atentamente,</p>#*firma*#"
	    BodySpanish = BodySpanish & "<p><b>IMPORTANTE:<br>Favor no responder este email ya que fue enviado desde un sistema automaticamente y no tendr&aacute; respuesta desde esta direcci&oacute;n de correo.</b></p>"
	    BodySpanish = BodySpanish & "<font color='white'>" & Session("Login") & " " & Request.Servervariables("REMOTE_ADDR") & "</font>"
        
        end if

        'response.write right(Countriesi,3) & "<br><br>"
        'response.write "<table>" & BodySpanish & "<br><br>"
        'response.end
    

        '/////////////////////////////////////////////////AGENTES INGLES//////////////////////////////////////////////////////////
	    BodyEnglish = headers & "<p>Dear Agent : </p><p>Here we present the current status of your shipment with the following information : </p>"
        if trim(HBLNumberi) = "" then
            BodyEnglish = BodyEnglish & "<p><b>AWB : </b>" & BLi & "<br>"        
        else
	        BodyEnglish = BodyEnglish & "<p><b>HAWB : </b>" & HBLNumberi & "<br>"        
        end if	
	    BodyEnglish = BodyEnglish & "<b>Consignee : </b>" & ConsignerIDi & "<br><b>Shipper : </b>" & ShipperIDi & "<br>"

        if SendCoLoader = 1 then
	        BodyEnglish = BodyEnglish & "<b>Coloader : </b>" & Coloaderi & "<br>"	    
        end if

        if SendCoLoader = 1 or SendShipper = 1 then
            if SendNotify = 1 then 'id_cliente_order notify
                BodyEnglish = BodyEnglish & "<b>Notify Party : </b>" & Notifyi & "<br>"                	
            end  if 
        end if 



        
        if right(Countriesi,3) = "TLA" then

        BodyEnglish = BodyEnglish & "<b>STATUS : </b><font color=green>" & Status & "</font><br></p>"
        BodyEnglish = BodyEnglish & "<p style='text-align:justify'>If you need any additional information please contact our Operations Department e-mail: <a href='mailto:" & LCase(ContactoEmail) & "'>" & ContactoEmail & "</a> or by phone: " & Replace(ContactoTel,"Supervisora de Servicio al cliente","") & ".</p>"
	    BodyEnglish = BodyEnglish & "<p>Cordially,</p>#*firma*#"
	    BodyEnglish = BodyEnglish & "<p> <b>IMPORTANT:<br>Please do not reply to this email, it is sent from an automated system, there will be no response from this address. For assistance contact customer service department.</b></p>"
	    BodyEnglish = BodyEnglish & "<font color='white'>" & Session("Login") & " " & Request.Servervariables("REMOTE_ADDR") & "</font>"

        else

        BodyEnglish = BodyEnglish & "<b>STATUS : </b><font color=green>" & Status & "</font><br></p>"
        BodyEnglish = BodyEnglish & "<p style='text-align:justify'>If you need any additional information please visit our tracking on web page <a href='http://#*home_page*#'>#*home_page*#</a> or you can also contact our Customer Service Department e-mail: <a href='mailto:" & LCase(ContactoEmail) & "'>" & ContactoEmail & "</a> Phone: " & Replace(ContactoTel,"Supervisora de Servicio al cliente"," Customer Service Manager") & ".</p>"
	    BodyEnglish = BodyEnglish & "<p>To request a username and password to access to the tracking please contact a customer service representative.</p>"
	    BodyEnglish = BodyEnglish & "<p>Cordially,</p>#*firma*#"
	    BodyEnglish = BodyEnglish & "<p> <b>IMPORTANT:<br>Please do not reply to this email, it is sent from an automated system, there will be no response from this address. For assistance contact customer service department.</b></p>"
	    BodyEnglish = BodyEnglish & "<font color='white'>" & Session("Login") & " " & Request.Servervariables("REMOTE_ADDR") & "</font>"

        end if


        BodyEnglish = BodyEnglish & "<br>"
        BodyEnglish = BodyEnglish & "<hr>"
        BodyEnglish = BodyEnglish & "<br>"
    
	    BodyEnglish = BodyEnglish & "<p>Estimado Agente : </p><p>A continuaci&oacute;n le damos a conocer el status actual de su mercaderia amparada con la siguiente informaci&oacute;n : </p>"	
        if trim(HBLNumberi) = "" then
            BodyEnglish = BodyEnglish & "<p><b>AWB : </b>" & BLi & "<br>"
        else
	        BodyEnglish = BodyEnglish & "<p><b>HAWB : </b>" & HBLNumberi & "<br>"
        end if

        BodyEnglish = BodyEnglish & "<b>Cliente : </b>" & ConsignerIDi & "<br><b>Shipper : </b>" & ShipperIDi & "<br>"

        if SendCoLoader = 1 then
	        BodyEnglish = BodyEnglish & "<b>Coloader : </b>" & Coloaderi & "<br>"
        end if

        if SendCoLoader = 1 or SendShipper = 1 then
            if SendNotify = 1 then 'id_cliente_order notify
                BodyEnglish = BodyEnglish & "<b>Notify Party : </b>" & Notifyi & "<br>"        
            end  if 
        end if 

        BodyEnglish = BodyEnglish & "<b>STATUS : </b>" & Status & "<br>"
        
        if right(Countriesi,3) = "TLA" then

	    BodyEnglish = BodyEnglish & "<p style='text-align:justify'>Si necesita mayor informaci&oacute;n de su carga puede consultar nuestro departamento de Operaciones <a href='mailto:" & LCase(ContactoEmail) & "'>" & ContactoEmail & "</a> o tel&eacute;fono: " & ContactoTel & ".</p>"
	    BodyEnglish = BodyEnglish & "<p>Estamos para servirle,</p><p>Atentamente,</p>#*firma*#"
	    BodyEnglish = BodyEnglish & "<p><b>IMPORTANTE:<br>Favor no responder este email ya que fue enviado desde un sistema automaticamente y no tendr&aacute; respuesta desde esta direcci&oacute;n de correo.</b></p>"
	    BodyEnglish = BodyEnglish & "<font color='white'>" & Session("Login") & "</font>"

        else

	    BodyEnglish = BodyEnglish & "<p style='text-align:justify'>Si necesita mayor informaci&oacute;n de su carga puede visitar nuestro tracking en la pagina web: <a href='http://#*home_page*#'>#*home_page*#</a> o bien consultar con nuestro departamento de Servicio al Cliente <a href='mailto:" & LCase(ContactoEmail) & "'>" & ContactoEmail & "</a> Telefono: " & ContactoTel & ".</p>"
	    BodyEnglish = BodyEnglish & "<p>Para solicitar un usuario y password para acceso al tracking puede comunicarse con nuestro representante de Servicio al Cliente.</p>"
	    BodyEnglish = BodyEnglish & "<p>Estamos para servirle,</p><p>Atentamente,</p>#*firma*#"
	    BodyEnglish = BodyEnglish & "<p><b>IMPORTANTE:<br>Favor no responder este email ya que fue enviado desde un sistema automaticamente y no tendr&aacute; respuesta desde esta direcci&oacute;n de correo.</b></p>"
	    BodyEnglish = BodyEnglish & "<font color='white'>" & Session("Login") & "</font>"

        end if

        '/////////////////////////////////////////////////CLIENTE INGLES//////////////////////////////////////////////////////////
	    Body_b = headers & "<p>Dear Consignee : </p><p>Here we present the current status of your shipment with the following information : </p>"
        if trim(HBLNumberi) = "" then
            Body_b = Body_b & "<p><b>AWB : </b>" & BLi & "<br>"        
        else
	        Body_b = Body_b & "<p><b>HAWB : </b>" & HBLNumberi & "<br>"        
        end if	

	    Body_b = Body_b & "<b>Consignee : </b>" & ConsignerIDi & "<br><b>Shipper : </b>" & ShipperIDi & "<br>"

        if SendCoLoader = 1 then
	        Body_b = Body_b & "<b>Coloader : </b>" & Coloaderi & "<br>"	    
        end if

        if SendCoLoader = 1 or SendShipper = 1 then
            if SendNotify = 1 then 'id_cliente_order notify
                Body_b = Body_b & "<b>Notify Party : </b>" & Notifyi & "<br>"                	
            end  if 
        end if 

        Body_b = Body_b & "<b>STATUS : </b><font color=green>" & Status & "</font><br></p>"

        if right(Countriesi,3) = "TLA" then

        Body_b = Body_b & "<b>Comments : </b><font color='green'>" & Comentario & "</font></p>"
        Body_b = Body_b & "<p style='text-align:justify'>If you need any additional information please contact our Operations Department e-mail: <a href='mailto:" & LCase(ContactoEmail) & "'>" & ContactoEmail & "</a> or by phone: " & Replace(ContactoTel,"Supervisora de Servicio al cliente","") & ".</p>"	    
	    Body_b = Body_b & "<p>Cordially,</p>#*firma*#"
	    Body_b = Body_b & "<p> <b>IMPORTANT:<br>Please do not reply to this email, it is sent from an automated system, there will be no response from this address. For assistance contact customer service department.</b></p>"
	    Body_b = Body_b & "<font color='white'>" & Session("Login") & " " & Request.Servervariables("REMOTE_ADDR") & "</font>"
		
        else

        Body_b = Body_b & "<b>Comments : </b><font color='green'>" & Comentario & "</font></p>"
	    Body_b = Body_b & "<p style='text-align:justify'>If you need any additional information please visit our tracking on web page <a href='http://#*home_page*#'>#*home_page*#</a> or you can also contact our Customer Service Department e-mail: <a href='mailto:" & LCase(ContactoEmail) & "'>" & ContactoEmail & "</a> Phone: " & Replace(ContactoTel,"Supervisora de Servicio al cliente"," Customer Service Manager") & ".</p>"
	    Body_b = Body_b & "<p>To request a username and password to access to the tracking please contact a customer service representative.</p>"
	    Body_b = Body_b & "<p>Cordially,</p>#*firma*#"
	    Body_b = Body_b & "<p> <b>IMPORTANT:<br>Please do not reply to this email, it is sent from an automated system, there will be no response from this address. For assistance contact customer service department.</b></p>"
	    Body_b = Body_b & "<font color='white'>" & Session("Login") & " " & Request.Servervariables("REMOTE_ADDR") & "</font>"
		
        end if


        'from = "tracking@aimargroup.com"    
    
        if Mode = 1 then 'Mode es 1 cuando viene de php
            response.write "(*" & CountListaMails & ")<br>\n"
        end if

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
                EmailStr = contactoXpais1(ListaMails(11,i), ListaMails(1,i), Countriesi)
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
                        'response.write "nuevo " & ListaMails(8,i) & " " & arr1(i) & " " & local_ & "<br>"
                        'response.write ( "(" & Mid(Countries, 1, 2) & ")<br>" )
                'if Mode = 1 then 'Mode es 1 cuando viene de php                    
                    'response.write (  ListaMails(1,i) & " " & ListaMails(8,i) & " " &  ListaMails(9,i) & "<br>" )			
                'end if	

                        if ListaMails(8,i) <> "Desarrollo" then 'Copia
                            if ListaMails(9,i) = "Si" then 'Copia
                                '//////////////////////////////////////////////////////////////////////////////////////////////////////
                                if local_ = 0 then
                                    if ListaMails(8,i) = "Agente" then                            
                                        if InStr(1, EmailStr, "aimargroup") = 0 then
                                            'comentado temporalmente solicitado por Ticket#2015062404000623 — Fwd: Re: NUEVO SISTEMA DE AUTO NOTIFICACIONES AIMAR Linea 6                                            
                                            'result = SendMail(BodyEnglish, EmailStr, subject, Countriesi)
                                            'eSendYesAge = eSendYesAge & " " & EmailStr & "<br>"                                    
                                        end if
                                    else
                                        'if Countries <> "GT" then
                                            'response.write "Paso Aqui 1<br>"
                                            if Left(Countriesi, 2) = "BZ" then
                                                result = SendMail(Body_b, EmailStr, subject, Countriesi)
                                            else
                                                result = SendMail(BodySpanish, EmailStr, subject, Countriesi)
                                            end if
                                        'end if
                                    end if
                                else
                                    result = -2
                                end if
                            end if
                            
                            'response.write ( EmailStr & "<br>" )

                            if result = -1 or result = -2  then                            
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

                if result = -1 or result = -2 then
                    Select Case ListaMails(8,i)            
                    Case "Agente","Consigneer","Shipper","Coloader","Notify"
				        BodyNoEnvio = BodyNoEnvio & "<font color=red>IMPORTANTE:</font> No se pudo enviar estatus al "
				        'if ListaMails(3,i) = 2 then BodyNoEnvio = BodyNoEnvio & "Contacto del "
				        BodyNoEnvio = BodyNoEnvio & ListaMails(8,i) & "<br>"
                        BodyNoEnvio = BodyNoEnvio & "<b>" & ListaMails(12,i) & " - " & ListaMails(0,i) & "</b> E-mail : " & EmailStr & "<br>"
				        if trim(EmailStr) = "" then
					        BodyNoEnvio = BodyNoEnvio & "no tiene cuenta de correo, favor de revisar y actualizar<br><br>"
				        else
                            if result = -1 then
					            BodyNoEnvio = BodyNoEnvio & "tiene error en cuenta de correo, favor de revisar y actualizar<br><br>"
                            else
                                BodyNoEnvio = BodyNoEnvio & "por seguridad desde localhost no envia a clientes<br><br>"
                            end if
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


            result = -4

		    'response.write("DESARROLLO<BR>")                
		    'A DESARROLLO 				
            'if Countries <> "GT" then		
            for i=0 to CountListaMails		
            
                'if Mode = 1 then 'Mode es 1 cuando viene de php                    
                '    response.write (  ListaMails(1,i) & " " & ListaMails(8,i) & " " &  ListaMails(9,i) & "<br>\n" )			
                'end if	
				
                if ListaMails(8,i) = "Desarrollo" and ListaMails(9,i) = "Si" then 'Copia

                    'response.write ListaMails(8,i) & " " & ListaMails(9,i) & "<br>" 
                    'response.write BodySpanish & "<br>****************<br>"                     
                    'response.write eSendYes  & "<br>****************<br>" 
                    'response.write eSendYesAge & "<br>****************<br>" 
                    'response.write Divisiones & "<br>****************<br>" 
                    
                    'response.write ListaMails(1,i) & "<br>"
                    'response.write subject & "<br>"
                    'response.write from & "<br>"                                        

                    result = SendMail(BodySpanish & eSendYes & eSendYesAge & Divisiones, ListaMails(1,i), subject, Countriesi)
                                            
                    'if result < 0 then
                    '    response.write IIF(Mode = 1,"","<span style='color:orange'>")
                    'else
                    '    response.write IIF(Mode = 1,"","<span style='color:green'>")
                    'end if
                    '
                    'select case result
                    'case -3
                    '    response.write "Reviser los parametros de configuracion de empresa." 
                    'case -2
                    '    response.write "Error en servidor de correos."
                    'case -1
                    '    response.write "Error al procesar correo."
                    'case else
                    '    response.write "Autonotificacion enviada correctamente." 
                    'end select
                    
                    'response.write IIF(Mode = 1,"","<span><br>")
                                                                                         
                end if
            Next

            if result = -4 then
                'response.write Divisiones & "<br><br>"
                'response.write BodyNoEnvio & "<br><br>"
            end if

            'end if

		    'response.write("RECHAZOS<BR>")
		    'RECHAZOS
		    if trim(BodyNoEnvio) <> "" then
				BodyNoEnvio = BodyNoEnvio & "A continuacion el mensaje original: <br><br>"
				BodyNoEnvio = BodyNoEnvio & "<hr><b>SUBJECT : " & subject & "</b><br><br><hr>" 
                                        
                if Mid(Countriesi, 1, 2) = "BZ" then                                                
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

                                'response.write "Paso Aqui 3<br>"
                                    result = SendMail(BodyNoEnvio, ListaMails(1,i), "NO SE ENVIO:" & subject, Countriesi)
                                else
                                    if local_ = 0 then
                                        result = SendMail(BodyNoEnvio, ListaMails(1,i), "NO SE ENVIO:" & subject, Countriesi)
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

    end if
    
    Set ListaMails = Nothing

    SendNotification = 1

End Function 




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





    Function GetaList1Values(AWBID,AWBType)  
        
        dim CountList1Values, aList1Values, SQLQuery, Conn, rs
               
	    CountList1Values = -1 
        
        OpenConn Conn

        SQLQuery = "select AWBNumber, HAWBNumber, ifnull(ConsignerID,0), ifnull(ShipperID,0), ifnull(AgentID,0), Countries, ifnull(id_coloader,0), ifnull(id_cliente_order,0), 1 as no, RoutingID from "
        
		if AWBType = 1 or AwbType = "1" then
		    SQLQuery = SQLQuery & "Awb"
		else 
            SQLQuery = SQLQuery & "Awbi" '2
		end if
        'esto debido a que hubo un caso que no trajo nada 
        SQLQuery = SQLQuery & " where AWBID=" & AWBID & " union select 'n/a','n/a',0,0,0,'',0,0,2 as no, 0 order by no asc"
		
        'response.write(SQLQuery & "<br>")

		set rs = Conn.Execute(SQLQuery)
		if Not rs.EOF then
		    aList1Values = rs.GetRows
			CountList1Values = rs.RecordCount - 1
		end if        

        'response.write("(" & AWBType  & ")(" & CountList1Values & ")<br>")

	    CloseOBJs rs, Conn
        GetaList1Values = aList1Values        

    End Function 



    Function GetSQLQuery(SendAgent, SendConsigner, SendShipper, aList1Values2, AwbType)  

        dim SendCoLoader, SendNotify, SQLQuery 

        if CheckNum(aList1Values2(6,0)) > 0 then 'coloader
            SendCoLoader = 1
        else
            SendCoLoader = 0
        end if

        if CheckNum(aList1Values2(7,0)) > 0 then 'notify
            SendNotify = 1
        else
            SendNotify = 0
        end if

        SQLQuery = "SELECT nombre, email, telefono, pais, area, impexp, carga, tranship, tipo_persona, copia, rechazo, contactoxpais, id_catalogo, id_contacto, case when tipo_persona = 'Desarrollo' then 10 when tipo_persona = 'Contacto' then 9 when tipo_persona = 'Soporte' then 8 when tipo_persona = 'Shipper' then 6 when tipo_persona = 'Coloader' then 5 when tipo_persona = 'Consigneer' then 4 when tipo_persona = 'Agente' then 3 else 0 end as sort FROM contactos_divisiones " & _ 
        "WHERE status = 'Activo' AND area ILIKE '%Aereo%' AND " & _ 
        "( (catalogo = 'USUARIO' AND pais ILIKE '%""" & aList1Values2(5,0) & """%') " & SQLQuery & ") "

        if AwbType = 1 or AwbType = "1" then 
            SQLQuery = SQLQuery & " AND impexp ILIKE '%Export%'"
        else
            SQLQuery = SQLQuery & " AND impexp ILIKE '%Import%'"  '2
        end if		

        if SendAgent = 1 then                        
            if SQLQuery <> "" then
                SQLQuery = SQLQuery & " UNION "
            end if
            SQLQuery = SQLQuery & EmailsExternos(3,CheckNum(aList1Values2(4,0)),"Agente",aList1Values2(5,0))
        end if
        
        if SendCoLoader = 1 and (SendShipper = 1 or SendConsigner = 1) then            
            if SQLQuery <> "" then
                SQLQuery = SQLQuery & " UNION "
            end if
            SQLQuery = SQLQuery & EmailsExternos(5,CheckNum(aList1Values2(6,0)),"Coloader","")
        else
            if SendShipper = 1 then
                if SQLQuery <> "" then
                    SQLQuery = SQLQuery & " UNION "
                end if
                SQLQuery = SQLQuery & EmailsExternos(6,CheckNum(aList1Values2(3,0)),"Shipper","")
            end if

		    if SendConsigner = 1 then
		        if SQLQuery <> "" then
                    SQLQuery = SQLQuery & " UNION "
                end if
                SQLQuery = SQLQuery & EmailsExternos(4,CheckNum(aList1Values2(2,0)),"Consigneer","")
            end if
        end if

        if SendShipper = 1 then
            
            if SendNotify = 1 then 'id_cliente_order notify
                if SQLQuery <> "" then
                    SQLQuery = SQLQuery & " UNION "
                end if
                SQLQuery = SQLQuery & EmailsExternos(7,CheckNum(aList1Values2(7,0)),"Notify","")
            end  if 

        end if

        if SQLQuery <> "" then
            SQLQuery = "select * from (" & SQLQuery & ") x order by sort" 
        end if
     
        GetSQLQuery = SQLQuery

    End Function 

%>

