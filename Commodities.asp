<%
Checking "0|1"
'Dim NameES, NameEN, TypeVal, CommodityCode
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	NameES = aTableValues(4, 0)
	NameEN = aTableValues(5, 0)
	TypeVal = aTableValues(6, 0)
	CommodityCode = aTableValues(7, 0)
	Arancel_GT = aTableValues(8, 0)
	Arancel_SV = aTableValues(9, 0)
	Arancel_HN = aTableValues(10, 0)
	Arancel_NI = aTableValues(11, 0)
	Arancel_CR = aTableValues(12, 0)
	Arancel_PA = aTableValues(13, 0)
	Arancel_BZ = aTableValues(14, 0)
end if
%>
<HTML>
<HEAD><TITLE>AWB - Aimar - Administraci�n</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
  		if (Action != 3) {
			if (!valTxt(document.forma.NameES, 3, 5)){return (false)};
			if (!valSelec(document.forma.TypeVal)){return (false)};
			//if (!valTxt(document.forma.CommodityCode, 4, 10)){return (false)};
	    }
		document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
		<TR><TD class=label align=right><b>C�digo:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creaci�n:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
        <TR><TD class=label align=right><b>Activa:</b></TD><TD class=label align=left><INPUT name=Expired TYPE=checkbox class=label <%If Expired = False Then response.write " checked"  End If%>></TD></TR>
		<TR><TD class=label align=right><b>Nombre Espa&ntilde;ol:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="NameES" id="Nombre Espa�ol" value="<%=NameES%>" maxlength="60"></TD></TR> 
        <TR><TD class=label align=right><b>Nombre Ingl&eacute;s:</TD><TD class=label align=left><INPUT TYPE=text class=label name="NameEN" value="<%=NameEN%>" maxlength="60"></TD></TR>
		<TR><TD class=label align=right><b>C�digo SCR:</TD><TD class=label align=left><INPUT TYPE=text class=label name="CommodityCode" id="C�digo SCR" value="<%=CommodityCode%>" maxlength="10"></TD></TR>
		<TR><TD class=label align=right><b>Tipo:</TD><TD class=label align=left>
		<select class=label name="TypeVal" id="Tipo Producto">
			<option	value="-1">Seleccionar
			<option	value="1">PERECEDERO
			<option	value="2">CARGA SECA
			<option	value="3">PELIGROSA
		</select>
		</TD></TR>
		<TR><TD class=label align=right><b>C�digo Arancel GT:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Arancel_GT" id="C�digo Arancelario Guatemala" value="<%=Arancel_GT%>" maxlength="20"></TD></TR>
		<TR><TD class=label align=right><b>C�digo Arancel SV:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Arancel_SV" id="C�digo Arancelario El Salvador" value="<%=Arancel_SV%>" maxlength="20"></TD></TR>
		<TR><TD class=label align=right><b>C�digo Arancel HN:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Arancel_HN" id="C�digo Arancelario Honduras" value="<%=Arancel_HN%>" maxlength="20"></TD></TR>
		<TR><TD class=label align=right><b>C�digo Arancel NI:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Arancel_NI" id="C�digo Arancelario Nicaragua" value="<%=Arancel_NI%>" maxlength="20"></TD></TR>
		<TR><TD class=label align=right><b>C�digo Arancel CR:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Arancel_CR" id="C�digo Arancelario Costa Rica" value="<%=Arancel_CR%>" maxlength="20"></TD></TR>
		<TR><TD class=label align=right><b>C�digo Arancel PA:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Arancel_PA" id="C�digo Arancelario Panama" value="<%=Arancel_PA%>" maxlength="20"></TD></TR>
		<TR><TD class=label align=right><b>C�digo Arancel BZ:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Arancel_BZ" id="C�digo Arancelario Belice" value="<%=Arancel_BZ%>" maxlength="20"></TD></TR>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
				  <TR>
							<%if CountTableValues = -1 then%>
									 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(4)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label></TD>
							<%else%>
									 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
									 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(3)" value="&nbsp;&nbsp;Eliminar&nbsp;&nbsp;" class=label></TD>
							<%end if%>
					</TR>
			</TABLE>
		<TD>
		</TR>
	</FORM>
	</TABLE>
<script>
selecciona('forma.TypeVal','<%=TypeVal%>');
</script>
</BODY>
</HTML>
