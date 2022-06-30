<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	Table = "Awb"
	CountTableValues = -1

	Dim Camps(23)
	Camps(2)="AdditionalChargeName1"
	Camps(3)="AdditionalChargeName2"
	Camps(4)="AdditionalChargeName6"
	Camps(5)="AdditionalChargeName7"
	Camps(6)="AdditionalChargeName9"
	Camps(7)="AdditionalChargeName10"
	Camps(8)="AdditionalChargeName11"
	Camps(9)="AdditionalChargeName12"
	Camps(10)="AdditionalChargeName13"
	Camps(11)="AdditionalChargeName14"
	Camps(12)="AdditionalChargeName15"
	Camps(13)="AdditionalChargeVal1"
	Camps(14)="AdditionalChargeVal2"
	Camps(15)="AdditionalChargeVal6"
	Camps(16)="AdditionalChargeVal7"
	Camps(17)="AdditionalChargeVal9"
	Camps(18)="AdditionalChargeVal10"
	Camps(19)="AdditionalChargeVal11"
	Camps(20)="AdditionalChargeVal12"
	Camps(21)="AdditionalChargeVal13"
	Camps(22)="AdditionalChargeVal14"
	Camps(23)="AdditionalChargeVal15"

	OpenConn Conn
	Set rs = Conn.Execute("select AWBID, AWBNumber, AdditionalChargeName1, AdditionalChargeName2, AdditionalChargeName6, AdditionalChargeName7, " & _
		"AdditionalChargeName9, AdditionalChargeName10, AdditionalChargeName11, AdditionalChargeName12, AdditionalChargeName13, " & _
		"AdditionalChargeName14, AdditionalChargeName15, " & _
		"AdditionalChargeVal1, AdditionalChargeVal2, AdditionalChargeVal6, " & _
		"AdditionalChargeVal7, AdditionalChargeVal9, AdditionalChargeVal10, AdditionalChargeVal11, AdditionalChargeVal12, " & _
		"AdditionalChargeVal13, AdditionalChargeVal14, AdditionalChargeVal15 from " & Table)
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJ rs
	response.write "<table border=1>"
	for i=0 to CountTableValues
		for j=2 to 12
			if aTableValues(j+11,i)<>"" then
				if aTableValues(j,i)="PICK UP" or aTableValues(j,i)="PICK-UP" or aTableValues(j,i)="PICKUP" or aTableValues(j,i)="PICUP" then
					Conn.Execute("update " & Table & " Set PickUp='" & aTableValues(j+11,i) & "', " & Camps(j) & "='', " & Camps(j+11) & "='' where AWBID=" & aTableValues(0,i))
					response.write "update " & Table & " Set PickUp='" & aTableValues(j+11,i) & "', " & Camps(j) & "='', " & Camps(j+11) & "='' where AWBID=" & aTableValues(0,i) & "<br><br>"
				end if
				if aTableValues(j,i)="SEED FEE" or aTableValues(j,i)="SED FEE" or aTableValues(j,i)="SED FILING FEE" then
					Conn.Execute("update " & Table & " Set SedFilingFee='" & aTableValues(j+11,i) & "', " & Camps(j) & "='', " & Camps(j+11) & "='' where AWBID=" & aTableValues(0,i))
					response.write "update " & Table & " Set SedFilingFee='" & aTableValues(j+11,i) & "', " & Camps(j) & "='', " & Camps(j+11) & "='' where AWBID=" & aTableValues(0,i) & "<br><br>"
				end if
			end if				
		next
	next
	CloseOBJ Conn
	
	Set aTableValues=Nothing
	response.write "</table>"
%>