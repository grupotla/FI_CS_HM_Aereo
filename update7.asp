<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	CountTableValues = -1
	OpenConn Conn

	Set rs = Conn.Execute("select x from test")
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If

	for i=0 to CountTableValues
		Set rs = Conn.Execute("select Name, Expired, Address, Address2, Phone1, Phone2, CreatedDate, CreatedTime, AccountNo, IATANo, DefaultVal, Countries from Agents where Name = '" & Trim(aTableValues(0,i)) & "'")
		do while not rs.EOF
			if rs(1) = 0 then
				activo = true
			else
				activo = false
			end if
			response.Write "insert into agentes (agente, activo, direccion, telefono, fax, fecha_creacion, hora_creacion, AccountNo, IATANo, DefaultVal, Countries) values ('" & rs(0) & "', " & activo & ", '" & Trim(rs(2) & " " & rs(3)) & "', '" & rs(4) & "', '" & rs(5) & "', '" & rs(6) & "', " & rs(7) & ", '" & rs(8) & "', '" & rs(9) & "', " & rs(10) & ", '" & rs(11) & "');<br>"
			rs.MOveNext
		loop
		CloseOBJ rs
	next

	CloseOBJ Conn
	Set aTableValues=Nothing
%>
listo