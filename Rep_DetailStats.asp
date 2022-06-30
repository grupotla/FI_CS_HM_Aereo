<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim AWBNumber, CarrierID, AgentID, SalespersonID, ShipperID, AirportDepID, AirportDesID, AwbType, MM, YY, Conn, rs, aListValues, CountListValues
Dim i, Option1, Option2, Option3, Option4, Option5, Option6, Option7, Option8, QuerySelect, HTMLTitle, HTMLCode, MoreOptions
Dim ShipperName, ConsignerName, HTMLTit

	 CountListValues = -1
	 AwbType = CheckNum(Request.Form("AwbType"))
	 MM = CheckNum(Request.Form("MM"))
	 YY = CheckNum(Request.Form("YY"))
	 QuerySelect = "select a.AWBID, a.CreatedDate, a.CreatedTime, a.AWBNumber, a.HAWBNumber, c.Name, a.ShipperData, a.ConsignerData, b.AirportCode, a.AduanaValue, a.TotWeight, a.TotPrepaid, a.TotCollect"
	 if AwbType = 1 then
		QuerySelect = QuerySelect & " from Awb a, Airports b, Carriers c "
		HTMLTit = "Resultados de Guias EXPORT "
	 else
		QuerySelect = QuerySelect & " from Awbi a, Airports b, Carriers c "
		HTMLTit = "Resultados de Guias IMPORT "
	 end if
	 
	 Option1 = " a.CarrierID=c.CarrierID and a.AirportDesID=b.AirportID and a.Countries in " & Session("Countries") & _
	 		   " and a.CreatedDate>='" & YY & "-" & TwoDigits(MM) & "-01' and a.CreatedDate<'" & SetFilterMonth(MM, YY) & "'"
	 HTMLTit = HTMLTit & " de " & NameOfMonth(MM) & " " & YY
	 HTMLTitle ="<tr><td class=titlelist align=center>&nbsp;</td><td class=titlelist align=center><b>Fecha</b></td>" & _ 
	 			"<td class=titlelist align=center width=100><b>Guia Master</b></td><td class=titlelist align=center><b>Guia House</b></td>>" & _
				"<td class=titlelist align=center><b>Transportista</b></td><td class=titlelist align=center><b>Embarcador</b></td>" & _
				"<td class=titlelist align=center><b>Consignatario</b></td><td class=titlelist align=center><b>Destino</b>" & _
				"</td><td class=titlelist align=center><b>Valor Declarado</b></td><td class=titlelist align=center><b>Peso Total</b></td>" & _
				"<td class=titlelist align=center><b>Total PREPAGO</b></td><td class=titlelist align=center><b>Total COLLECT</b></td></tr>"
	 
	 AWBNumber = Request.Form("AWBNumber")
	 CarrierID = CheckNum(Request.Form("CarrierID"))
	 AgentID = CheckNum(Request.Form("AgentID"))
	 SalespersonID = CheckNum(Request.Form("SalespersonID"))
	 ShipperID = CheckNum(Request.Form("ShipperID"))
	 AirportDepID = CheckNum(Request.Form("AirportDepID"))
	 AirportDesID = CheckNum(Request.Form("AirportDesID"))
	 
	 if AWBNumber <> "" then
		Option2 = " a.AWBNumber like '%" & AWBNumber & "%' "					
	 end if
	 if CarrierID <> 0 then
			Option3 = " a.CarrierID=" & CarrierID & " "
	 end if
	 if AgentID <> 0 then
			Option4 = " a.AgentID=" & AgentID & " "
	 end if
	 if SalespersonID <> 0 then
			Option5 = " a.SalespersonID=" & SalespersonID & " "
	 end if
	 if ShipperID <> 0 then
			Option6 = " a.ShipperID=" & ShipperID & " "
	 end if
	 if AirportDepID <> 0 then
			Option7 = " a.AirportDepID=" & AirportDepID & " "
	 end if
	 if AirportDesID <> 0 then
			Option8 = " a.AirportDesID=" & AirportDesID & " "
	 end if

	MoreOptions = 0
	CreateSearchQuery QuerySelect, Option1, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option2, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option3, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option4, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option5, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option6, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option7, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option8, MoreOptions, " and "
	QuerySelect = QuerySelect & " order by AWBID"
	HTMLCode = ""

	OpenConn Conn
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
		aListValues = rs.GetRows
		CountListValues = rs.RecordCount-1
	end if
	CloseOBJs rs, Conn

	for i=0 to CountListValues
		ShipperName = Split(aListValues(6,i), chr(13))
		ConsignerName = Split(aListValues(7,i), chr(13))
		HTMLCode = HTMLCode & "<tr><td class=label align=right>" & (i+1) & "</td>" & _
				"<td class=label align=right>" & aListValues(1,i) & "</td>" & _
				"<td class=label align=right><a href=InsertData.asp?GID=1&OID=" & aListValues(0,i) & "&CD=" & aListValues(1,i) & "&CT=" & aListValues(2,i) & "&AT=" & AwbType & ">" & aListValues(3,i) & "</a></td>" & _
				"<td class=label align=right>&nbsp;<a href=InsertData.asp?GID=1&OID=" & aListValues(0,i) & "&CD=" & aListValues(1,i) & "&CT=" & aListValues(2,i) & "&AT=" & AwbType & ">" & aListValues(4,i) & "</a></td>" & _
				"<td class=label align=right>" & aListValues(5,i) & "</td>" & _
				"<td class=label align=right>" & ShipperName(0) & "</td>" & _
				"<td class=label align=right>" & ConsignerName(0) & "</td>" & _
				"<td class=label align=right>" & aListValues(8,i) & "</td>" & _
				"<td class=label align=right>" & aListValues(9,i) & "</td>"
		select case aListValues(10,i)
		case "", "NaN"
			HTMLCode = HTMLCode & "<td class=label align=right>0</td>"
		case else
			HTMLCode = HTMLCode & "<td class=label align=right>" & Round(aListValues(10,i),2) & "</td>"
		end select


		select case aListValues(11,i)
		case "", "NaN"
			HTMLCode = HTMLCode & "<td class=label align=right>0</td>"
		case "AS AGREED"
			HTMLCode = HTMLCode & "<td class=label align=right>AS&nbsp;AGREED</td>"
		case else
			HTMLCode = HTMLCode & "<td class=label align=right>" & Round(aListValues(11,i),2) & "</td>"
		end select

		select case aListValues(12,i)
		case "", "NaN"
			HTMLCode = HTMLCode & "<td class=label align=right>0</td></tr>"
		case "AS AGREED"
			HTMLCode = HTMLCode & "<td class=label align=right>AS&nbsp;AGREED</td>"
		case else
			HTMLCode = HTMLCode & "<td class=label align=right>" & Round(aListValues(12,i),2) & "</td></tr>"
		end select
	next
%>
<HTML><HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function DisplayDetailStats (MM, YY) {
	document.forma.MM.value = MM;
	document.forma.YY.value = YY;
	document.forma.submit();
}
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<TR>
		<TD width=40% colspan=2 class=label align=right valign=top>
			<TABLE cellspacing=1 cellpadding=3 align=center border=1>
			 <tr><td class=titlelist3 align=center colspan=12><%=HTMLTit%></td><tr>
			 <%=HTMLTitle%>
			 <%=HTMLCode%>
			</TABLE>
		</TD>
	  </TR>
	  <TR>
	  <TD width=40% colspan=2 class=label align=right valign=top>
		<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
		<TR>
		<TD class=label align=left>
		<a class=label onclick=JavaScript:history.back(); href=# target=_self><u><< Regresar</u></a>
		</TD>
		</TR>
		</TABLE>
	  </TD>
	  </TR>
	</TABLE>
</BODY>
</HTML>
<%
	Set aListValues = Nothing
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>

