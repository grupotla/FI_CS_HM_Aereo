<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	CountTableValues = -1
	OpenConn Conn
	Set rs = Conn.Execute("select distinct AgentID, '', '', '', '', AgentData from Awb order by AgentData")
	
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If

	for i=0 to CountTableValues
		Name = Split(aTableValues(5,i),chr(13))
		select case aTableValues(0,i)
		case 0
			Set rs = Conn.Execute("select Name, Expired, Address, Address2, Phone1, Phone2, CreatedDate, CreatedTime, AccountNo, IATANo, DefaultVal, Countries from Agents where Name like '%" & Name(0) & "%'")
		case else
			Set rs = Conn.Execute("select Name, Expired, Address, Address2, Phone1, Phone2, CreatedDate, CreatedTime, AccountNo, IATANo, DefaultVal, Countries from Agents where AgentID="& aTableValues(0,i))
		end select
		if not rs.EOF then
			aTableValues(1,i) = rs(0)
			aTableValues(2,i) = Trim(rs(2) & " " & rs(3))
			if rs(1) = 0 then
				Activo = true
			else
				Activo = false
			end if
			aTableValues(3,i) = "insert into agentes (agente, activo, direccion, telefono, fax, fecha_creacion, hora_creacion, accountno, iatano, defaultval, countries) values ('" & rs(0) & "', " & Activo & ", '" & Trim(rs(2) & " " & rs(3)) & "', '" & rs(4) & "', '" & rs(5) & "', '" & rs(6) & "', " & rs(7) & ", '" & rs(8) & "', '" & rs(9) & "', " & rs(10) & ", '" & rs(11) & "');"
			aTableValues(4,i) = "update agentes set agente = '" & rs(0) & "', activo=" & Activo & ", direccion='" & Trim (rs(2) & " " & rs(3)) & "', telefono='"& rs(4) & "', fax='" & rs(5) & "', fecha_creacion='" & rs(6) & "', hora_creacion=" & rs(7) & ", accountno='" & rs(8) & "', iatano='" & rs(9) & "', defaultval=" & rs(10) & ", countries='" & rs(11) & "' where agente_id="
			aTableValues(5,i) = " where AgentID=" & aTableValues(0,i) & ";"
		else
			aTableValues(5,i) = ""
		end if
		CloseOBJ rs
	next
	CloseOBJ Conn

	response.write "<table border=1>" & _
		"<tr><td>AgentID</td><td>Query</td><td>Query2</td><td>Query3</td><td>Direccion</td><td>Agent</td><td>AGENTE</td><td>DIRECCION</td><td>AGENTE_ID</td></tr>"
	OpenConn2 Conn
	for i=0 to CountTableValues
		if aTableValues(5,i) <> "" then
			'Namex = Split(aTableValues(1,i),chr(13))
			'Names = Namex(0)
			'Names = Replace(aTableValues(1,i),chr(13),"<br>",1,-1)
			'Set rs = Conn.Execute("select a.agente, a.direccion, a.agente_id from agentes a where a.agente ilike '" & Mid(aTableValues(1,i),1,7) & "%'")
			Set rs = Conn.Execute("select a.agente, a.direccion, a.agente_id from agentes a where a.agente ilike '" & Mid(aTableValues(1,i),1,14) & "%'")
			If Not rs.EOF Then
				'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & aTableValues(2,i) & "</td><td>" & rs(0) & "</td><td>" & rs(1) & "</td></tr>"
				'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & Names(0) & "</td><td>" & rs(0) & "</td><td>" & rs(1) & "</td></tr>"
				response.write "<tr><td>" & aTableValues(0,i) &  "</td><td>" & aTableValues(4,i) & rs(2) & ";</td><td>" & aTableValues(3,i) & "</td><td>update Agents set Name='" & rs(0) & "' " & aTableValues(5,i) & "</td><td>" & aTableValues(2,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & rs(0) & "</td><td>" & rs(1) & "</td><td>" & rs(2) & "</td></tr>"
			Else
				'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & aTableValues(2,i) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
				'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & Names(0) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
				response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(4,i) & "</td><td>" & aTableValues(3,i) & "</td><td>" & aTableValues(5,i) & "</td><td>" & aTableValues(2,i) & "</td><td>" & aTableValues(1,i) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
			End If
			CloseOBJ rs
		end if
	next
	CloseOBJ Conn
	Set aTableValues=Nothing
	response.write "</table>"
%>
listo