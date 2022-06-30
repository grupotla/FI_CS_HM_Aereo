<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	Table = "Awb"
	CountTableValues = -1
	OpenConn Conn
	Set rs = Conn.Execute("select AgentID, AgentData, AwbID from " & Table & " order by AgentData")
	
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	'Para asignar AgentsID a documentos que tengan AgentID=0
	response.write "<table border=1>"
	for i=0 to CountTableValues
		Name = Split(aTableValues(1,i),chr(13))
		select case aTableValues(0,i)
		case 0
			Set rs = Conn.Execute("select Name, AgentID from Agents where Name like '%" & Name(0) & "%'")
		case else
			Set rs = Conn.Execute("select Name, AgentID from Agents where AgentID=" & aTableValues(0,i))
		end select
		if not rs.EOF then
			aTableValues(1,i) = rs(0)
			response.write "<tr><td>" & Name(0) & "</td><td>" & rs(0) & "</td><td>update " & Table & " set AgentID=" & rs(1) & " where AwbID=" & aTableValues(2,i) & "</td></tr>"
		else
			response.write "<tr><td>" & Name(0) & "</td><td></td><td>" & aTableValues(2,i) & "</td></tr>"
		end if
		CloseOBJ rs
	next
	CloseOBJ Conn
	Set aTableValues=Nothing
	response.write "</table>"
%>
listo