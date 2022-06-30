<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID
	GroupID = CheckNum(Request("GID"))
%>

<HTML><HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validate(){
	 	document.forma.submit();
	}
	 	 
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="Javascript:self.focus();">
	<FORM name="forma" action="Search_ResultsAWBData.asp" method="post" target=_self>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<br>
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<%select case GroupID
		Case 23 'Notify	%>
		<TR>
		<TD class=label align=center colspan="2"><b>Notificar</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR> 
		<% Case 7, 21, 24 'Consigneer	%>
		<TR>
		<TD class=label align=center colspan="2"><b>Destinatario</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR> 

		<% Case 22 'Coloader %>
		<TR>
		<TD class=label align=center colspan="2"><b>CoLoader</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR> 


		<% Case 8 'Agentes %>
		<TR>
		<TD class=label align=center colspan="2"><b>Agente</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR> 
		<% Case 10 'Embarcadores %>
		<TR>
		<TD class=label align=center colspan="2"><b>Embarcadores</b></TD>
		</TR>
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR> 
		<% Case 11 'Commodities %>
		<TR>
		<TD class=label align=center colspan="2"><b>Productos</b></TD>
		</TR>
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="NameES" value="" size=30></TD>
		</TR> 
		<% Case 17 'Routings %>
		<TR>
		<TD class=label align=center><b>Routings</b></TD>
		</TR>
		<TR>
		<TD class=label align=center><b>Routing:</b>&nbsp;
		<INPUT TYPE=text class=label name="Routing" value="<%=Request("routing")%>" size=30>
		<INPUT name="AT" type=hidden value="<%=CheckNum(Request("AT"))%>">
		</TD>
		</TR> 
		<%Case 18, 19 'Rubros %>
		<TR>
		<TD class=label align=center colspan="2"><b>Rubros</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left>
		<INPUT TYPE=text class=label name="Name" value="" size=30>
		<INPUT name="N" type=hidden value="<%=Request("N")%>">
		<INPUT name="NID" type=hidden value="<%=Request("NID")%>">
		</TD>
		</TR> 
		<% Case 20 'Proveedores %>
		<TR>
			<%select Case Request("ST")%>
			<%case 0%>
			<TD class=label align=center colspan="2"><b>Lineas Aereas</b></TD>
			<%case 1%>
			<TD class=label align=center colspan="2"><b>Agentes</b></TD>
			<%case 2%>
			<TD class=label align=center colspan="2"><b>Navieras</b></TD>
			<%case 3%>		
			<TD class=label align=center colspan="2"><b>Proveedores</b></TD>
			<%end select%>
		</TR> 
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left>
		<INPUT TYPE=text class=label name="Name" value="" size=30>
		<INPUT name="N" type=hidden value="<%=Request("N")%>">
		<INPUT name="ST" type=hidden value="<%=Request("ST")%>">
		</TD>
		</TR> 		
		<%
		end select
		%>
		</TABLE>
		<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
				<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validate();" value="&nbsp;&nbsp;Buscar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD>
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
