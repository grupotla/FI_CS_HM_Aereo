<%
Checking "0|1"
%>
<HTML>
<HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
  		document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=430 align=center>
		<TR>
		<TD colspan=2 class=label align=right valign=top>
			<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
			<%If HTMLCode <> "" then%>
			<tr>
			<td class="label">Existen datos similares, por favor revise la siguiente informaci&oacute;n:
			  <ul>
			  <li>Si sus datos ya se encuentran en la lista no es necesario ingresarla nuevamente, si desea <b>actualizarlos</b> haga click sobre el nombre correspondiente.</li>
			  <li>Si sus datos son nuevos presione el bot&oacute;n al final de la lista <%if SearchOption = 1 then%> y luego <b>"asigne"</b> la informacion al documento<%end if%></li>
			  </ul>
			  </td>
			</tr>
			<%=HTMLCode%>
			<tr>
			<%End If%>
			<td class="label" align="center">
				<FORM name="forma" action="InsertData.asp" method="post" target=_self>
				<%=VirtualForm%>
				<%If HTMLCode <> "" then%>
				<INPUT name=enviar type=button onClick="JavaScript:validar(1);" value='Agregar <%=UCase(request.Form("Name"))%>' class=label>
				<%end if%>
				</FORM>				
			</td> 
			</tr>
			</TABLE>
		</TD>
	  </TR>	  
	  </TABLE>
</BODY>
<%if HTMLCode = "" then%>
<SCRIPT LANGUAGE="JavaScript">
	validar(1);
</SCRIPT>
<%end if%>
</HTML>