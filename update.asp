<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	Server.ScriptTimeOut = 600
	OpenConn Conn
	OpenConn2 Conn2
	
	Conn.Execute("delete from ChargeItems where AWBID in (4206, 4209, 3177, 3367, 3357)")
	
	Conn.Execute("update ChargeItems set Expired=0")
	Conn.Execute("update ChargeItems set Local=0 where Local=-1")
	
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='0'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='-1'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='1'")
	Conn.Execute("update ChargeItems set CurrencyID='GTQ' where CurrencyID='2'")
	Conn.Execute("update ChargeItems set CurrencyID='EUR' where CurrencyID='3'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='4'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='5'")
	Conn.Execute("update ChargeItems set CurrencyID='LPS' where CurrencyID='6'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='7'")
	Conn.Execute("update ChargeItems set CurrencyID='NIO' where CurrencyID='8'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='9'")
	Conn.Execute("update ChargeItems set CurrencyID='CRC' where CurrencyID='10'")
	Conn.Execute("update ChargeItems set CurrencyID='EUR' where CurrencyID='11'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='12'")
	Conn.Execute("update ChargeItems set CurrencyID='BZD' where CurrencyID='13'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='14'")
	Conn.Execute("update ChargeItems set CurrencyID='GTQ' where CurrencyID='16'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='17'")
	Conn.Execute("update ChargeItems set CurrencyID='GTQ' where CurrencyID='18'")
	Conn.Execute("update ChargeItems set CurrencyID='USD' where CurrencyID='19'")
	
	'AGREGANDO LOS NOMBRES DE LOS RUBROS
	Set rs = Conn.Execute("select distinct ItemID from ChargeItems")
	do While not rs.EOF
		Set rs2 = Conn2.Execute("select desc_rubro_es from rubros where id_rubro=" & rs(0))
		if Not rs2.EOF then
			'response.write "update ChargeItems set ItemName='" & rs2(0) & "' where ItemID=" & rs(0) & "<br>"
			Conn.Execute("update ChargeItems set ItemName='" & rs2(0) & "' where ItemID=" & rs(0))
		end if
		CloseOBJ rs2
		rs.MoveNext
	loop
	CloseOBJ rs
	
	'ASIGNANDO POS, CREATEDTIME y CREATEDDATE DEL AWB A LOS RUBROS CORRESPONDIENTES EXPORT
	Set rs = Conn.Execute("select distinct a.AWBID, a.CreatedDate, a.CreatedTime, a.AdditionalChargeName3, a.AdditionalChargeName4, a.AdditionalChargeName5, a.AdditionalChargeName8, a.AdditionalChargeVal3, a.AdditionalChargeVal4, a.AdditionalChargeVal5, a.AdditionalChargeVal8, a.AdditionalChargeName1, a.AdditionalChargeName2, a.AdditionalChargeName6, a.AdditionalChargeName7, a.AdditionalChargeName9, a.AdditionalChargeName10, a.AdditionalChargeName11, a.AdditionalChargeName12, a.AdditionalChargeName13, a.AdditionalChargeName14, a.AdditionalChargeName15, a.AdditionalChargeVal1, a.AdditionalChargeVal2, a.AdditionalChargeVal6, a.AdditionalChargeVal7, a.AdditionalChargeVal9, a.AdditionalChargeVal10, a.AdditionalChargeVal11, a.AdditionalChargeVal12, a.AdditionalChargeVal13, a.AdditionalChargeVal14, a.AdditionalChargeVal15 from Awb a, ChargeItems b where b.DocTyp=0 and a.AWBID=b.AWBID")
	do While not rs.EOF
		j = 0
		Items = "14,15,11,12,13,31,38,115,116"
		
		Set rs2 = Conn.Execute("select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=0 and (ItemID in (14,15,11,12,13))")
		Do While Not rs2.EOF
			Conn.Execute("update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", AgentTyp=0 where DocTyp=0 and AWBID=" & rs(0) & " and ItemID=" & rs2(0))
			j=j+1
			rs2.MoveNext
		Loop
		CloseOBJ rs2

		Set rs2 = Conn.Execute("select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=0 and (ItemID in (31,38,115,116))")
		Do While Not rs2.EOF
			Conn.Execute("update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", AgentTyp=1 where DocTyp=0 and AWBID=" & rs(0) & " and ItemID=" & rs2(0))
			j=j+1
			rs2.MoveNext
		Loop
		CloseOBJ rs2
		
		For i=0 to 3
			if rs(3+i) <> "" then
				response.write "select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=0 and ItemName='" & rs(3+i) & "'<br>"
				Set rs2 = Conn.Execute("select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=0 and ItemName='" & rs(3+i) & "'")
				if rs2.EOF then
					Set rs3 = Conn2.Execute("select id_rubro from rubros where desc_rubro_es='" & rs(3+i) & "'")
					if Not rs3.EOF then
						Items = Items & "," & rs3(0)
						response.write "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(7+i))  & ", 0, 0, 0, '" & rs(3+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")<br>"
						Conn.Execute("insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(7+i))  & ", 0, 0, 0, '" & rs(3+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")")
					else
						response.write "NO SE ENCONTRO EL ID DEL RUBRO " & rs(3+i) & " EN LA MASTER<br>"
					end if
					CloseOBJ rs3
				else
					Items = Items & "," & rs2(0)
					response.write "update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=0 where DocTyp=0 and AWBID=" & rs(0) & " and ItemID=" & rs2(0) & "<br>" 
					Conn.Execute("update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=0 where DocTyp=0 and AWBID=" & rs(0) & " and ItemID=" & rs2(0))
				end if
				CloseOBJ rs2				
				j=j+1
			end if
		Next

		For i=0 to 10
			if rs(11+i) <> "" then
				response.write "select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=0 and ItemName='" & rs(11+i) & "'<br>"
				Set rs2 = Conn.Execute("select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=0 and ItemName='" & rs(11+i) & "'")
				if rs2.EOF then
					Set rs3 = Conn2.Execute("select id_rubro from rubros where desc_rubro_es='" & rs(11+i) & "'")
					if Not rs3.EOF then
						Items = Items & "," & rs3(0)
						response.write "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(22+i))  & ", 0, 1, 0, '" & rs(11+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")<br>"
						Conn.Execute("insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(22+i))  & ", 0, 1, 0, '" & rs(11+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")")
					else						
						response.write "NO SE ENCONTRO EL ID DEL RUBRO " & rs(11+i) & " EN LA MASTER<br>"
					end if
					CloseOBJ rs3
				else
					Items = Items & "," & rs2(0)
					response.write "update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=1 where DocTyp=0 and AWBID=" & rs(0) & " and ItemID=" & rs2(0) & "<br>" 
					Conn.Execute("update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=1 where DocTyp=0 and AWBID=" & rs(0) & " and ItemID=" & rs2(0))
				end if
				CloseOBJ rs2
				j=j+1
			end if
		Next
		response.write "delete from ChargeItems where AWBID=" & rs(0) & " and DocTyp=0 and ItemID Not In (" & Items & ")<br>" 
		Conn.Execute("delete from ChargeItems where AWBID=" & rs(0) & " and DocTyp=0 and ItemID Not In (" & Items & ")")
		rs.MoveNext
	loop
	CloseOBJ rs

	'ASIGNANDO POS, CREATEDTIME y CREATEDDATE DEL AWB A LOS RUBROS CORRESPONDIENTES IMPORT
	Set rs = Conn.Execute("select distinct a.AWBID, a.CreatedDate, a.CreatedTime, a.AdditionalChargeName3, a.AdditionalChargeName4, a.AdditionalChargeName5, a.AdditionalChargeName8, a.AdditionalChargeVal3, a.AdditionalChargeVal4, a.AdditionalChargeVal5, a.AdditionalChargeVal8, a.AdditionalChargeName1, a.AdditionalChargeName2, a.AdditionalChargeName6, a.AdditionalChargeName7, a.AdditionalChargeName9, a.AdditionalChargeName10, a.AdditionalChargeName11, a.AdditionalChargeName12, a.AdditionalChargeName13, a.AdditionalChargeName14, a.AdditionalChargeName15, a.AdditionalChargeVal1, a.AdditionalChargeVal2, a.AdditionalChargeVal6, a.AdditionalChargeVal7, a.AdditionalChargeVal9, a.AdditionalChargeVal10, a.AdditionalChargeVal11, a.AdditionalChargeVal12, a.AdditionalChargeVal13, a.AdditionalChargeVal14, a.AdditionalChargeVal15, a.OtherChargeName1, a.OtherChargeName2, a.OtherChargeName3, a.OtherChargeName4, a.OtherChargeName5, a.OtherChargeName6, a.OtherChargeVal1, a.OtherChargeVal2, a.OtherChargeVal3, a.OtherChargeVal4, a.OtherChargeVal5, a.OtherChargeVal6 from Awbi a, ChargeItems b where b.DocTyp=1 and a.AWBID=b.AWBID")
	do While not rs.EOF
		j = 0
		Items = "11,12,13,31,38,115"

		Set rs2 = Conn.Execute("select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and (ItemID in (11,12,13))")
		Do While Not rs2.EOF
			Conn.Execute("update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", AgentTyp=0 where DocTyp=1 and AWBID=" & rs(0) & " and ItemID=" & rs2(0))
			j=j+1
			rs2.MoveNext
		Loop
		CloseOBJ rs2

		Set rs2 = Conn.Execute("select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and (ItemID in (31,38,115))")
		Do While Not rs2.EOF
			Conn.Execute("update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", AgentTyp=1 where DocTyp=1 and AWBID=" & rs(0) & " and ItemID=" & rs2(0))
			j=j+1
			rs2.MoveNext
		Loop
		CloseOBJ rs2
		
		For i=0 to 3
			if rs(3+i) <> "" then
				response.write "select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and ItemName='" & rs(3+i) & "'<br>"
				Set rs2 = Conn.Execute("select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and ItemName='" & rs(3+i) & "'")
				if rs2.EOF then
					Set rs3 = Conn2.Execute("select id_rubro from rubros where desc_rubro_es='" & rs(3+i) & "'")
					if Not rs3.EOF then
						Items = Items & "," & rs3(0)
						response.write "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(7+i))  & ", 0, 0, 1, '" & rs(3+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")<br>"
						Conn.Execute("insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(7+i))  & ", 0, 0, 1, '" & rs(3+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")")
					else
						response.write "NO SE ENCONTRO EL ID DEL RUBRO " & rs(3+i) & " EN LA MASTER<br>"
					end if
					CloseOBJ rs3
				else
					Items = Items & "," & rs2(0)
					response.write "update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=0 where DocTyp=1 and AWBID=" & rs(0) & " and ItemID=" & rs2(0) & "<br>" 
					Conn.Execute("update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=0 where DocTyp=1 and AWBID=" & rs(0) & " and ItemID=" & rs2(0))
				end if
				CloseOBJ rs2				
				j=j+1
			end if
		Next

		For i=0 to 10
			if rs(11+i) <> "" then
				response.write "select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and ItemName='" & rs(11+i) & "'<br>"
				Set rs2 = Conn.Execute("select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and ItemName='" & rs(11+i) & "'")
				if rs2.EOF then
					Set rs3 = Conn2.Execute("select id_rubro from rubros where desc_rubro_es='" & rs(11+i) & "'")
					if Not rs3.EOF then
						Items = Items & "," & rs3(0)
						response.write "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(22+i))  & ", 0, 1, 1, '" & rs(11+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")<br>"
						Conn.Execute("insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(22+i))  & ", 0, 1, 1, '" & rs(11+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")")
					else						
						response.write "NO SE ENCONTRO EL ID DEL RUBRO " & rs(11+i) & " EN LA MASTER<br>"
					end if
					CloseOBJ rs3
				else
					Items = Items & "," & rs2(0)
					response.write "update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=1 where DocTyp=1 and AWBID=" & rs(0) & " and ItemID=" & rs2(0) & "<br>" 
					Conn.Execute("update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=1 where DocTyp=1 and AWBID=" & rs(0) & " and ItemID=" & rs2(0))
				end if
				CloseOBJ rs2
				j=j+1
			end if
		Next

		For i=0 to 5
			if rs(33+i) <> "" then
				response.write "select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and ItemName='" & rs(33+i) & "'<br>"
				Set rs2 = Conn.Execute("select ItemID from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and ItemName='" & rs(33+i) & "'")
				if rs2.EOF then
					Set rs3 = Conn2.Execute("select id_rubro from rubros where desc_rubro_es='" & rs(33+i) & "'")
					if Not rs3.EOF then
						Items = Items & "," & rs3(0)
						response.write "insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(39+i))  & ", 0, 2, 1, '" & rs(33+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")<br>"
						Conn.Execute("insert into ChargeItems (AWBID, CurrencyID, ItemID, Value, Local, AgentTyp, DocTyp, ItemName, CreatedDate, CreatedTime, Pos) values (" & rs(0)  & ", 'USD', " & rs3(0) & ", " & CheckNum(rs(39+i))  & ", 0, 2, 1, '" & rs(33+i) & "', '" & ConvertDate(rs(1),2) & "', " & rs(2)+j & ", " & i+1 & ")")
					else						
						response.write "NO SE ENCONTRO EL ID DEL RUBRO " & rs(33+i) & " EN LA MASTER<br>"
					end if
					CloseOBJ rs3
				else
					Items = Items & "," & rs2(0)
					response.write "update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=2 where DocTyp=1 and AWBID=" & rs(0) & " and ItemID=" & rs2(0) & "<br>" 
					Conn.Execute("update ChargeItems set CreatedDate='" & ConvertDate(rs(1),2) & "', CreatedTime=" & rs(2)+j & ", Pos=" & i+1 & ", AgentTyp=2 where DocTyp=1 and AWBID=" & rs(0) & " and ItemID=" & rs2(0))
				end if
				CloseOBJ rs2
				j=j+1
			end if
		Next

		response.write "delete from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and ItemID Not In (" & Items & ")<br>" 
		Conn.Execute("delete from ChargeItems where AWBID=" & rs(0) & " and DocTyp=1 and ItemID Not In (" & Items & ")")
		rs.MoveNext
	loop
	CloseOBJ rs

	'REVISAR QUE LOS REPORTES NO CAMBIEN

	CloseOBJ Conn2
	CloseOBJ Conn

	'response.write "<table border=1>"
	'for i=0 to CountTableValues
    '	response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & aTableValues(2,i) & "</td></tr>"
		'response.write "<tr><td>" & aTableValues(0,i) & "</td></tr>"
	'next
	'Set aTableValues=Nothing
	'response.write "</table>"
%>
listo