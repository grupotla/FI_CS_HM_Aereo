<%@ Language=VBScript%>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<!-- #INCLUDE file="MD5.asp" -->
<%
Dim Conn, rs, OperatorID, OperatorCountry, OperatorLogin, OperatorPwd, OperatorLevel, Menu, ProductCount, SystemDate, pw_passwd_dias, pw_passwd_fecha, OperatorPais, exactus_pagos
Dim resultStr
OperatorLogin = replace(Request.Form("login"),"'","",1,-1)
OperatorPwd = replace(Request.Form("pwd"),"'","",1,-1)
OperatorID = 0
exactus_pagos = "0"

	'Borrando Posibles Sesiones Antiguas
	If Trim(OperatorLogin) <> "" And Trim(OperatorPWD) <> "" Then
		'OperatorPWD = MD5(OperatorPWD) 2016-05-24
		openConn2 Conn
            resultStr = "SELECT id_usuario, pais, pw_passwd_dias, pw_passwd_fecha, to_date(cast(pw_passwd_fecha + interval '1' day * pw_passwd_dias as varchar),'yyyy-mm-dd') as fec_res, CURRENT_DATE - to_date(cast(pw_passwd_fecha + interval '1' day * pw_passwd_dias as varchar),'yyyy-mm-dd') as days, CAST(exactus_pagos as text) FROM usuarios_empresas WHERE pw_activo=1 and pw_name='" & OperatorLogin & "' and (pw_passwd=md5('" & OperatorPWD & "') OR pw_passwd='" & OperatorPWD & "') LIMIT 1"
			response.write resultStr & "<br>"
            Set rs = Conn.Execute(resultStr)
            if Not rs.EOF Then
				OperatorID = Cint(rs(0))
                OperatorCountry = Trim(rs(1))
                pw_passwd_dias = rs(5)
                pw_passwd_fecha = rs(4)
                Session("OperatorPais") = rs(1)
                exactus_pagos = rs(6)
			end if
		CloseOBJs rs, Conn
		 
		if OperatorID > 0 then
            

            'response.write "(" & pw_passwd_dias & ")(" & pw_passwd_fecha & ")<br>******************************************"

            if pw_passwd_dias > 0 then
                Response.Redirect ("default.asp?MS=5")
            else 

			    openConn Conn
			    Set rs = Conn.Execute("select OperatorLevel, Email, FirstName, LastName, Countries, Sign from Operators where Active=True and OperatorID=" & OperatorID)
			    If Not rs.EOF Then
				    'Creando la Variable de Sesion
				    Session.SessionID
				    Session("Login") = OperatorLogin
				    Session("OperatorID") = OperatorID

                    if OperatorCountry <> "XX" then
                        Session("OperatorCountry") = OperatorCountry
				    else
                        Session("OperatorCountry") = ""
                    end if

				    Session("OperatorLevel") = rs(0)
				    Session("OperatorEmail") = rs(1)
				
                    if rs(2) = rs(3) then
				        Session("OperatorName") = rs(2)
                    else
                        Session("OperatorName") = rs(2) & " " & rs(3)
                    end if

                    Session("exactus_pagos") = exactus_pagos

				    if Trim(rs(4)) <> "" then
					    Session("Countries") = "(" & rs(4) & ")"
				    else
					    Session("Countries") = "('XX')"
				    end if
				    Session("Sign") = rs(5)
				    closeOBJ rs
				    'Seteando el Tiempo de Sesion
				    Set rs = Conn.Execute("select AdminTime, ClientName, SearchResults, HourDif, PBAValue, ClientURL from Miscellaneous")
					    Session.TimeOut = rs("AdminTime")
					    Session("ClientName") = rs("ClientName")
					    'Nivel de Categorias
					    Session("SearchResults") = rs("SearchResults")
					    SystemDate = Date + (CInt(rs("HourDif")) * 0.041666667)
					    Session("PBAValue") = rs("PBAValue")
					    Session("ClientURL") = rs("ClientURL")
					    Session("Date") = NameOfDay(WeekDay(SystemDate)) & " " & Day(SystemDate) & " de " & NameOfMonth(Month(SystemDate)) & " de " & Year(SystemDate)										
				    closeOBJs rs, Conn

Session("Pricing") = ""
                     
OpenConn3 Conn

resultStr = "SELECT string_agg(DISTINCT '''' || tpl_pais_fk || '''', ',') " & _ 
"FROM ti_pricing_list " & _ 
"INNER JOIN ti_pricing_ruta ON tpr_tpl_fk = tpl_pk AND tpr_tps_fk = 1 " & _ 
"INNER JOIN ti_pricing_articulo ON tpa_tpr_fk = tpr_pk AND tpa_tps_fk = 1 " & _  
"INNER JOIN ti_pricing_articulo_tarifa ON tpat_tpa_fk = tpa_pk AND tpa_tps_fk = 1 " & _  
"WHERE tpl_tps_fk = 1 "

Set rs = Conn.Execute(resultStr) '"SELECT string_agg(DISTINCT '''' || tpl_pais_fk || '''', ',') FROM ti_pricing_list WHERE tpl_tps_fk = '1'"
If Not rs.EOF Then 

    Session("Pricing") = rs(0)
    
end if
closeOBJs rs, Conn



                    resultStr = WsExactus_TIPO_DOC_CP("TIPO_DOC_CP",Session("Login"))



                    'if pw_passwd_dias > -10 then 'advertencia dias antes de vencer
                    '    Response.Redirect ("default.asp?MS=6&Dias=" & abs(pw_passwd_dias) ) 
                    'else 
                        Response.Redirect ("content.asp")
                    'end if

			    Else
				    closeOBJs rs, Conn
                    Response.Redirect ("default.asp?MS=2")
                End If

            End If

		Else
			Response.Redirect ("default.asp?MS=2")
		End if
	Else
		  Response.Redirect ("default.asp?MS=3")
	End If		
%>