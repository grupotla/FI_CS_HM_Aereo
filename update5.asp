<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	Table = "Awb"
	CountTableValues = -1

	Dim Camps(35)
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
	Camps(13)="OtherChargeName1"
	Camps(14)="OtherChargeName2"
	Camps(15)="OtherChargeName3"
	Camps(16)="OtherChargeName4"
	Camps(17)="OtherChargeName5"
	Camps(18)="OtherChargeName6"
	Camps(19)="AdditionalChargeVal1"
	Camps(20)="AdditionalChargeVal2"
	Camps(21)="AdditionalChargeVal6"
	Camps(22)="AdditionalChargeVal7"
	Camps(23)="AdditionalChargeVal9"
	Camps(24)="AdditionalChargeVal10"
	Camps(25)="AdditionalChargeVal11"
	Camps(26)="AdditionalChargeVal12"
	Camps(27)="AdditionalChargeVal13"
	Camps(28)="AdditionalChargeVal14"
	Camps(29)="AdditionalChargeVal15"
	Camps(30)="OtherChargeVal1"
	Camps(31)="OtherChargeVal2"
	Camps(32)="OtherChargeVal3"
	Camps(33)="OtherChargeVal4"
	Camps(34)="OtherChargeVal5"
	Camps(35)="OtherChargeVal6"

	OpenConn Conn
	Set rs = Conn.Execute("select AWBID, AWBNumber, AdditionalChargeName1, AdditionalChargeName2, AdditionalChargeName6, AdditionalChargeName7, " & _
		"AdditionalChargeName9, AdditionalChargeName10, AdditionalChargeName11, AdditionalChargeName12, AdditionalChargeName13, " & _
		"AdditionalChargeName14, AdditionalChargeName15, OtherChargeName1, OtherChargeName2, OtherChargeName3, OtherChargeName4, " & _
		"OtherChargeName5, OtherChargeName6, AdditionalChargeVal1, AdditionalChargeVal2, AdditionalChargeVal6, " & _
		"AdditionalChargeVal7, AdditionalChargeVal9, AdditionalChargeVal10, AdditionalChargeVal11, AdditionalChargeVal12, " & _
		"AdditionalChargeVal13, AdditionalChargeVal14, AdditionalChargeVal15, OtherChargeVal1, OtherChargeVal2, OtherChargeVal3, " & _
		"OtherChargeVal4, OtherChargeVal5, OtherChargeVal6 from " & Table)
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJ rs
	response.write "<table border=1>"
	for i=0 to CountTableValues
		for j=2 to 18
			if aTableValues(j+17,i)<>"" then
				if aTableValues(j,i)="PICK UP" or aTableValues(j,i)="PICK-UP" or aTableValues(j,i)="PICKUP" or aTableValues(j,i)="PICUP" then
					Conn.Execute("update " & Table & " Set PickUp='" & aTableValues(j+17,i) & "', " & Camps(j) & "='', " & Camps(j+17) & "='' where AWBID=" & aTableValues(0,i))
					response.write "update " & Table & " Set PickUp='" & aTableValues(j+17,i) & "', " & Camps(j) & "='', " & Camps(j+17) & "='' where AWBID=" & aTableValues(0,i) & "<br><br>"
				end if
				'if aTableValues(j,i)="FUEL SURCHARGE" then
				'	Conn.Execute("update " & Table & " Set FuelSurcharge='" & aTableValues(j+17,i) & "', " & Camps(j) & "='', " & Camps(j+17) & "='' where AWBID=" & aTableValues(0,i))
				'	response.write "update " & Table & " Set FuelSurcharge='" & aTableValues(j+17,i) & "', " & Camps(j) & "='', " & Camps(j+17) & "='' where AWBID=" & aTableValues(0,i) & "<br><br>"
				'end if
				if aTableValues(j,i)="SEED FEE" or aTableValues(j,i)="SED FEE" or aTableValues(j,i)="SED FILING FEE" then
					Conn.Execute("update " & Table & " Set SedFilingFee='" & aTableValues(j+17,i) & "', " & Camps(j) & "='', " & Camps(j+17) & "='' where AWBID=" & aTableValues(0,i))
					response.write "update " & Table & " Set SedFilingFee='" & aTableValues(j+17,i) & "', " & Camps(j) & "='', " & Camps(j+17) & "='' where AWBID=" & aTableValues(0,i) & "<br><br>"
				end if
				'if aTableValues(j,i)="SECURITY FEE" then
				'	Conn.Execute("update " & Table & " Set SecurityFee='" & aTableValues(j+17,i) & "', " & Camps(j) & "='', " & Camps(j+17) & "='' where AWBID=" & aTableValues(0,i))
				'	response.write "update " & Table & " Set SecurityFee='" & aTableValues(j+17,i) & "', " & Camps(j) & "='', " & Camps(j+17) & "='' where AWBID=" & aTableValues(0,i) & "<br><br>"
				'end if
			end if				
		next
	next
	CloseOBJ Conn
	
	Set aTableValues=Nothing
	response.write "</table>"
%>