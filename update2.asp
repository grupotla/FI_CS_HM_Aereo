<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	CountTableValues = -1
	Table = "Awbi"
	OpenConn Conn
	Set rs = Conn.Execute("select AWBID, ChargeableWeights, AWBNumber from " & Table)
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJ rs

	for i=0 to CountTableValues
		if aTableValues(1,i) <> "" then
			LenWeights = -1
			Weights = Split(aTableValues(1,i),chr(13))
			LenWeights = Ubound(Weights)
			if LenWeights > 0 then LenWeights = LenWeights-1
			for j=0 to LenWeights
				TotWeights = Weights(j)
			next
			Conn.Execute("update " & Table & " set TotWeightChargeable='" & TotWeights & "' where AWBID=" & aTableValues(0,i))
			response.write aTableValues(0,i) & "-" & aTableValues(1,i) & "-" & TotWeights & "<br>"
			TotWeights = 0
			Set Weights=Nothing
		else
			Conn.Execute("update " & Table & " set TotWeightChargeable='0' where AWBID=" & aTableValues(0,i))
			response.write aTableValues(0,i) & "-" & aTableValues(1,i) & "-0<br>"
		end if
	next
	CloseOBJ Conn
	Set aTableValues=Nothing
%>
listo