<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0"
Dim Conn, rs, Action, TaxValue, PBAValue, AdminTime, ClientName, SearchResults, ClientURL, HourDif, SystemDate
	 Action = CheckNum(Request.Form("Action"))
	 OpenConn Conn
	 If Action = 1 then
		'Obteniendo Datos de Setup
		TaxValue = CheckNum(Request.Form("TaxValue"))
		PBAValue = CheckNum(Request.Form("PBAValue"))
		AdminTime = CheckNum(Request.Form("AdminTime"))
		ClientName = PurgeData(Request.Form("ClientName"))
		SearchResults = CheckNum(Request.Form("SearchResults"))
		ClientURL = PurgeData(Request.Form("ClientURL"))	
		HourDif = CheckNum(Request.Form("HourDif"))
		Conn.Execute("delete from Miscellaneous")
		Conn.Execute("insert into Miscellaneous (TaxValue, PBAValue, AdminTime, ClientName, SearchResults, ClientURL, HourDif) Values (" & TaxValue & ", " & PBAValue & ", " & AdminTime & ", '" & ClientName & "', " & SearchResults & ", '" & ClientURL & "', " & HourDif & ")")
	 end if
 				 Set rs = Conn.Execute("select TaxValue, PBAValue, AdminTime, ClientName, SearchResults, ClientURL, HourDif from Miscellaneous")
				   if Not rs.EOF then
					 		TaxValue = rs(0)
							PBAValue = rs(1)
							AdminTime = rs(2)
							ClientName = rs(3)
					 		SearchResults = rs(4)
							ClientURL = rs(5)
							HourDif = rs(6)
					 else						 
					 		TaxValue = 0
							PBAValue = 0
							AdminTime = 0
							ClientName = ""
					 		SearchResults = 10
							ClientURL = ""
							HourDif = 0
					 end if
					closeObj rs
 				closeObj Conn
				Session.TimeOut = AdminTime
				Session("ClientName") = ClientName
				'Nivel de Categorias
				Session("SearchResults") = SearchResults
				SystemDate = Date + (CInt(HourDif) * 0.041666667)
				Session("PBAValue") = PBAValue
				Session("ClientURL") = ClientURL
				Session("Date") = NameOfDay(WeekDay(SystemDate)) & " " & Day(SystemDate) & " de " & NameOfMonth(Month(SystemDate)) & " de " & Year(SystemDate)
%>

<HTML><HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar() {
		if (!valTxt(document.forma.ClientName, 1, 5)){return (false)};
		if (!valTxt(document.forma.ClientURL, 1, 5)){return (false)};
		if (!valTxt(document.forma.TaxValue, 1, 5)){return (false)};
		if (!valTxt(document.forma.PBAValue, 1, 5)){return (false)};
		if (!valTxt(document.forma.AdminTime, 1, 5)){return (false)};
		if (!valTxt(document.forma.SearchResults, 1, 5)){return (false)};
		if (!valTxt(document.forma.HourDif, 1, 5)){return (false)};
	    document.forma.Action.value = 1;
		document.forma.submit();			 
	 }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<FORM name=forma action="setup.asp" method="post" target=_self>
	<INPUT name=Action type=hidden value=0>
	<TABLE border=0 cellspacing=0 cellpadding=2 width=100%>
	  <TR>
		<TD class=label align=right width=45%><b><i><u>Datos Empresa:</u></i></b></TD>
		<TD class=label align=left width=55%>&nbsp;</TD>
		</TR>
		<TR>
		<TD class=label align=right width=45%><b>Nombre:</b></TD>
		<TD class=label align=left width=55%><INPUT name=ClientName id="Nombre de la Empresa" type=text value="<%=ClientName%>" size=60 maxLength=50 class=label></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=45%><b>URL o IP del Cliente:</b></TD>
		<TD class=label align=left width=55%><INPUT name=ClientURL id="URL o IP de la Empresa" type=text value="<%=ClientURL%>" size=20 maxLength=50 class=label></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=45%><b><i><u>Configuraciones:</u></i></b></TD>
		<TD class=label align=left width=55%>&nbsp;</TD>
		</TR>
		<TR>
		<TD class=label align=right width=45%><b>Tax:</b></TD>
		<TD class=label align=left width=55%><INPUT name=TaxValue id="Tax" type=text value="<%=TaxValue%>" size=20 maxLength=50 class=label onKeyUp="res(this,numb);"></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=45%><b>PBA:</b></TD>
		<TD class=label align=left width=55%><INPUT name=PBAValue id="PBA" type=text value="<%=PBAValue%>" size=20 maxLength=50 class=label onKeyUp="res(this,numb);"></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=45%><b>Timeout de Sesión:</b></TD>
		<TD class=label align=left width=55%><INPUT name=AdminTime id="Timeout de Sesión" type=text value="<%=AdminTime%>" size=5 maxLength=4 class=label onKeyUp="res(this,numb);"></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=45%><b>Resultados por B&uacute;squeda:</b></TD>
		<TD class=label align=left width=55%><INPUT name=SearchResults id="Resultados por B&uacute;squeda" type=text value="<%=SearchResults%>" size=5 maxLength=4 class=label onKeyUp="res(this,numb);"></TD>
		<TR>
		<TR>
		<TD class=label align=right width=45%><b>Horario del sistema:</b></TD>
		<TD class=label align=left width=55%><%=Now%></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=45%><b>Diferencia de Horario con el sistema:</b></TD>
		<TD class=label align=left width=55%><INPUT name=HourDif id="Diferencia de Horario con el sistema" type=text value="<%=HourDif%>" size=5 maxLength=2 maxLength=50 class=label onKeyUp="res(this,numb);"></TD>
	  </TR>
	</TABLE>
	<TABLE border=0 cellspacing=0 cellpadding=2 width=95%>
	  <TR>
		<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar()" value="&nbsp;&nbsp;Configurar&nbsp;&nbsp;" class=label></TD>
	  </TR>
	</TABLE>
	</FORM>
</BODY>
</HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>