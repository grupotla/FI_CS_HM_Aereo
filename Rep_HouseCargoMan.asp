<%
Checking "0|1|2"
'Dim TaxNo, Address, Phone1, Phone2, AccountNo, Attn, Expired
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	AWBNumber = aTableValues(4, 0)
	Comment2 = aTableValues(5, 0)
    ManifestNumber = aTableValues(6, 0)
        flg_master = aTableValues(7, 0)
        flg_totals = aTableValues(8, 0)
end if
%>
<HTML>
<HEAD><TITLE>AWB - Aimar - Administraci�n</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
	    if (!valTxt(document.forma.Comment2, 1, 5)){return (false)};
		document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
	function abrir(Label){
	var DateSend, Subject;
		DateSend = document.forma(Label).value;
		Subject = '';	
		window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
	}
	
	function validate(){
	 	document.forma.submit();
	}
_editor_url = "Javascripts/";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
  document.write('<scr' + 'ipt src="' +_editor_url + 'editor.js"');
  document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }
</script>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="Expired" type=hidden value="on">
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
	<INPUT name="AT" type=hidden value="<%=AwbType%>">
		<TR><TD class=label align=right><b>Master AWB:</b></TD><TD class=label align=left><%=AWBNumber%></TD></TR>
		<TR><TD class=label align=right><b>Tipo:</b></TD><TD class=label align=left><%if AwbType = 1 then%>EXPORT<%else%>IMPORT<%end if%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creaci�n:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
		<TR><TD class=label align=right><b>C�digo:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR><TD class=label align=right><b>Comentario:</b></TD><TD class=label align=left><Textarea name="Comment2" id="Comentario" cols="30" rows="5"><%=Comment2%></Textarea></TD></TR> 
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
                 <% if flg_master = "0" and flg_totals = "0" then '2017-12-08 %>
				 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="Javascript:window.open('Reports.asp?Action=2&OID=<%=ObjectID%>&AT=<%=AwbType%>','HouseCargoMan','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=650,height=600');return false;" value="&nbsp;&nbsp;Previsualizar&nbsp;Mail&nbsp;&nbsp;" class=label></TD>
                 <% else %>
                 <TD class=label align=center colspan=2><INPUT name=enviar type=button  value="&nbsp;&nbsp;Previsualizar&nbsp;Mail&nbsp;&nbsp;" class=label disabled></TD>                  
                 <%end if%>

				 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</FORM>
	</TABLE>
</BODY>
<script language="javascript1.2">
editor_generate('Comment2');
</SCRIPT>

</HTML>