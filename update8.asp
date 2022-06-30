<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	CountTableValues = -1
	Table = "Awb"
	OpenConn Conn
	Set rs = Conn.Execute("select distinct AgentID, '', '', '', '', AgentData, AwbID from " & Table & " order by AgentData")
	
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
		end if
		CloseOBJ rs
	next
	CloseOBJ Conn

	response.write "<table border=1>" & _
		"<tr><td>Agent</td><td>id</td><td>AGENTE</td><td>ID</td><td>Query2</td></tr>"
	OpenConn2 Conn
	for i=0 to CountTableValues
		if aTableValues(5,i) <> "" and aTableValues(0,i) > 0  then
			'Namex = Split(aTableValues(1,i),chr(13))
			'Names = Namex(0)
			'Names = Replace(aTableValues(1,i),chr(13),"<br>",1,-1)
			Set rs = Conn.Execute("select a.agente, a.direccion, a.agente_id from agentes a where a.agente ilike '" & Mid(aTableValues(1,i),1,14) & "%'")
			If Not rs.EOF Then
				'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & aTableValues(2,i) & "</td><td>" & rs(0) & "</td><td>" & rs(1) & "</td></tr>"
				'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & Names(0) & "</td><td>" & rs(0) & "</td><td>" & rs(1) & "</td></tr>"
				response.write "<tr><td>" & rs(0) & "</td><td>" & rs(2) & "</td><td>" & aTableValues(1,i) & "</td><td>" & aTableValues(0,i) & "</td><td>update " & Table & " set AgentID=" & rs(2) & " where AwbID=" & aTableValues(6,i) & ";</td></tr>"
			Else
				'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & aTableValues(2,i) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
				'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & Names(0) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
				response.write "<tr><td></td><td></td><td>" & aTableValues(1,i) & "</td><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(6,i) & "</td></tr>"
			End If
			CloseOBJ rs
		end if
	next
	CloseOBJ Conn
	Set aTableValues=Nothing
	response.write "</table>"
%>
listo
