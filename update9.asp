<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
Opt = 1 '0=despliega los datos de vendedores entre las 2 tablas para igualar nombres, 1=por cada BLID busca el id de la tabla vieja en base al nombre para cambiar el ID
	
select case Opt
case 0
	CountTableValues = -1
	Table = "Awb"
	OpenConn Conn
	Set rs = Conn.Execute("select distinct SalespersonID from " & Table)
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJs rs, Conn

	response.write "<table border=1>" & _
		"<tr><td>id</td><td>nombre</td><td>pais</td><td>ID</td><td>NOMBRE</td><td>PAIS</td></tr>"
	OpenConn2 Conn
	for i=0 to CountTableValues
			'Namex = Split(aTableValues(1,i),chr(13))
			'Names = Namex(0)
			'Names = Replace(aTableValues(1,i),chr(13),"<br>",1,-1)
			Set rs = Conn.Execute("select id_usuario, nombre, id_pais from usuarios where id_usuario=" & aTableValues(0,i))
			If Not rs.EOF Then
				'Set rs2 = Conn.Execute("select id_usuario, pw_gecos, pais from usuarios_empresas where pw_gecos ilike '" & Mid(rs(1),1,12) & "%'")
				Set rs2 = Conn.Execute("select id_usuario, pw_gecos, pais from usuarios_empresas where pw_gecos = '" & rs(1) & "'")
				If Not rs2.EOF Then
					response.write "<tr><td>" & rs(0) & "</td><td>" & rs(1) & "</td><td>" & rs(2) & "</td><td>" & rs2(0) & "</td><td>" & rs2(1) & "</td><td>" & rs2(2) & "</td></tr>"
				Else
					response.write  "<tr><td>" & rs(0) & "</td><td>" & rs(1) & "</td><td>" & rs(2) & "</td><td></td><td></td><td></td></tr>"
				End If
				CloseOBJ rs2
			else
				response.write  "<tr><td>" & aTableValues(0,i) & "</td><td></td><td></td><td></td><td></td><td></td></tr>"
			End If
			CloseOBJ rs
	next
	CloseOBJ Conn
	Set aTableValues=Nothing
	response.write "</table>"
case 1
	CountTableValues = -1
	Table = "Awb"
	OpenConn Conn
	Set rs = Conn.Execute("select SalespersonID, AwbID from " & Table)
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJs rs, Conn

	OpenConn2 Conn
	OpenConn Conn2
	for i=0 to CountTableValues
			'Namex = Split(aTableValues(1,i),chr(13))
			'Names = Namex(0)
			'Names = Replace(aTableValues(1,i),chr(13),"<br>",1,-1)
			Set rs = Conn.Execute("select id_usuario, nombre, id_pais from usuarios where id_usuario=" & aTableValues(0,i))
			If Not rs.EOF Then
				'Set rs2 = Conn.Execute("select id_usuario, pw_gecos, pais from usuarios_empresas where pw_gecos ilike '" & Mid(rs(1),1,12) & "%'")
				
				Set rs2 = Conn.Execute("select id_usuario, pw_gecos, pais from usuarios_empresas where pw_gecos = '" & rs(1) & "'")
				If Not rs2.EOF Then
					response.write "update " & Table & " set SalespersonID=" & rs2(0) & " where AwbID=" & aTableValues(1,i) & "<br>"
					Conn2.Execute("update " & Table & " set SalespersonID=" & rs2(0) & " where AwbID=" & aTableValues(1,i))
				End If
				CloseOBJ rs2
			End If
			CloseOBJ rs
	next
	CloseOBJ Conn2
	CloseOBJ Conn
	Set aTableValues=Nothing
end select
%>
listo