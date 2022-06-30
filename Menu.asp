<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"

iPerfilOpcion = PerfilOpcion()

'iPerfilOpcion = PerfilOpcion2(23,Session("OperatorID"))

%>
<HTML><HEAD><TITLE>Sistema Aereo</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<!--<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>-->
<script>
var Menu1 = "";
var MenuColor1 = "";
var Menu2 = "";
var MenuColor2 = "";
var TitleMenu = "";

  var isNav = false;
  var isIE = false;
  var col1 = "";
  var styleObj = "";

  function showTitle(msg) { 
  
  }

  if (parseInt(navigator.appVersion) >= 4) {
     if(navigator.appName == "Netscape" ) {
		isNav = true;
     }
     else {
		isIE = true;
		col1 = "all.";
		styleObj = ".style";
     }	
  } else {
		styleObj = ".style";
  }
  //alert(navigator.appName);
  //alert(parseInt(navigator.appVersion));
  //alert(col1);
  //alert(styleObj);
  //alert(isIE);

  function getObject( obj ){
  var theObj;
  var x = typeof obj;
	//alert(obj);
	//alert(x);
	//alert(eval("document." + col1 + obj + styleObj));
	
	if (x == "string" ){
		if (parseInt(navigator.appVersion) < 5) {
			theObj = eval("document." + col1 + obj + styleObj );
		} else {
			theObj = eval(document.getElementById(obj).style);
		}
	}
	else
	{
		theObj = obj;
	}
	//alert(theObj);
	return theObj;
	
  }

	function showMenu( Main, MainColor, Level ){
	var ColorMenu;
	var ColorMain;
	var Menu;
	var MenuColor;	
	if (Level == 1) {
		 Menu = Menu1;
		 MenuColor = MenuColor1;
 		 ColorMenu = "#000066";
		 ColorMain = "#006699";
		 if (Menu2 != "") {
		 		var objectMenu2 = getObject( Menu2 );
				var objectMenuColor2 = getObject( MenuColor2 );
				objectMenu2.visibility = "hidden";
  		  		objectMenuColor2.background = "#000066";
				if (TitleMenu != "") {
					 var objectTitleMenu = getObject( TitleMenu );
					 objectTitleMenu.visibility = "hidden";
				}
		 }
	};
	getObject( Main );
	var objectMain = getObject( Main );
	var objectMainColor = getObject( MainColor );
	if (Menu == "") {
		 objectMain.visibility = "visible";
		 objectMainColor.background = ColorMenu;
		 Menu = Main;
		 MenuColor = MainColor;		 
	} else {
  		if (Menu != Main) {
  		 	var objectMenu = getObject( Menu );
				var objectMenuColor = getObject( MenuColor );
				objectMenu.visibility = "hidden";
  		  objectMain.visibility = "visible";
				objectMenuColor.background = ColorMain;
  		  objectMainColor.background = ColorMenu;
				Menu = Main;				
				MenuColor = MainColor;
			} else {
						 if (Menu == Main) {
						 		var objectMenu = getObject( Menu );
								var objectMenuColor = getObject( MenuColor );
								objectMenu.visibility = "hidden";
								objectMenuColor.background = ColorMain;								
								Menu = "";
								MenuColor = "";
							}
			}
	}
	if (Level == 1) {
		 Menu1 = Menu;
		 MenuColor1 = MenuColor;
	};
	if (Level == 2) {
		 Menu2 = Menu;
		 MenuColor2 = MenuColor;
	};
};





/*In your javascript*/
var prevItem = null;
function activateItem(t) {
    if (prevItem != null) {
        prevItem.className = prevItem.className.replace(/activeItem/, "");
   }
   t.className += " activeItem";
   prevItem = t;
}


</script>
<LINK REL="stylesheet" type="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">

<!--  Session("Login") & " " & Session("OperatorID") & " " & Session("OperatorEmail") & " " &  -->

<TABLE cellSpacing=0 cellPadding=0 width="100%" class=menu>
	<TR><TD colspan="10">
	<TABLE border=0 cellSpacing=0 cellPadding=2 width="100%" class=menu>
	<TR>
		<TD class=titlea vAlign=center align=left width="50%">&nbsp;&nbsp;Sistema Aereo<BR>&nbsp;&nbsp;<SPAN class=label><FONT color=#ffffff><%=Session("Date")%><!--Administrador AWB 2.0--></FONT></SPAN></TD>
        <TD class=titlea vAlign=center align=left width="50%">
            <span style="float:right">
            <%=Session("OperatorName")%> :: 
            
        <%
        dim tempo, lentempo, i
        tempo = Split(Replace(Replace(Replace(Session("Countries"),"(",""),")",""),"'",""), ",")
        lentempo = ubound(tempo)
        response.write "<select style='background-color:transparent;color:white;border:0px'><option> Oficinas </option>" 
        for i = 0 to lentempo
            response.write "<option>" & tempo(i) & "</option>"
        next
        response.write "</select>" 
        %>
                    
            <%'=Replace(Replace(Replace(Session("Countries"),"(",""),")",""),"'","")%>        <br />

            <%="IP " & Request.ServerVariables("REMOTE_ADDR")%></span>
        </TD>
	    <!--<TD valign=right align=middle width="10%"><IMG height=1 src="img/transparente.gif" width=1></TD>--> 
	</TR>
	</TABLE>
	<TABLE cellSpacing=0 cellPadding=0 width="100%" border="1">
	<TR>
		<TD class=border><IMG height=1 src="img/transparente.gif" width=1></TD>
	</TR>
	</TABLE>
	</TD></TR>
	<TR>
    <TD class=inactiveMain valign=center align=left>
		<table>
		<TR>
		<TD id=TDSetup0 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup0', 'TDSetup0', 1);">AWB</A>&nbsp;&nbsp;</TD>

        

		<TD class=separator vAlign=center align=left <%=IIf(iArr2.Item("23Log") = "1","","style='display:none;'")%>>|</TD>
		<TD id=TDSetup11 class=menu vAlign=center align=left>&nbsp;&nbsp;<A <%=IIf(iArr2.Item("23Log") = "1","","onclick='activateItem(this);return false;' style='display:none;'")%> class=activeMain href="javascript:showMenu('mainSetup11', 'TDSetup11', 1);">Batch</A>&nbsp;&nbsp;</TD>

        

        <%if Session("OperatorLevel")=0 or Session("OperatorLevel")=1 then%> 	
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup1 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup1', 'TDSetup1', 1)">Transportistas</A>&nbsp;&nbsp;</TD>
        <TD class=separator vAlign=center align=left>|</TD>
        <TD id=TDSetup10 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup10', 'TDSetup10', 1)">Catalogos</A>&nbsp;&nbsp;</TD>
        
		<!--
        <TD class=separator vAlign=center align=left>|</TD>
		<TD class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="http://10.10.1.20/catalogo_admin/login.php" target="_blank">Destinatarios</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="http://10.10.1.20/catalogo_admin/login.php" target="_blank">Embarcadores</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="http://10.10.1.20/catalogo_admin/login.php" target="_blank">Agentes</A>&nbsp;&nbsp;</TD>
		-->
        <!--
		<TD id=TDSetup2 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup2', 'TDSetup2', 1)">Destinatarios</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup5 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup5', 'TDSetup5', 1)">Embarcadores</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup3 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup3', 'TDSetup3', 1)">Agentes</A>&nbsp;&nbsp;</TD>
		-->
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup4 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup4', 'TDSetup4', 1)">Aeropuertos</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup6 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup6', 'TDSetup6', 1)">Productos</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup7 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup7', 'TDSetup7', 1)">Monedas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup8 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup8', 'TDSetup8', 1)">Rangos</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<%end if
		if Session("OperatorLevel")=0 then%> 	
		<!-- <TD id=TDSetup9 class=menu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup9', 'TDSetup9', 1)">&nbsp;&nbsp;&nbsp;Configuración&nbsp;&nbsp;&nbsp;</A></TD> -->
		<!-- <TD class=menu vAlign=center width="100%" align=left>&nbsp;&nbsp;</TD> -->
		<%end if%>
		</TR>
		</table>
		</TD>
		<!--<TD class=menu vAlign=center width="100%" align=left>&nbsp;&nbsp;</TD>-->
		<TD width="1" class=separator vAlign=middle align=right>|</TD>
		<TD  width="50" id=MisDatos class=menu vAlign=center align=middle>&nbsp;&nbsp;<A class=activeMain onClick="javascript:activateItem(this);showTitle('mainMyData');showMenu('mainMisDatos', 'MisDatos', 1)" href="MyData.asp" target=principal><B>&nbsp;&nbsp;&nbsp;Mis&nbsp;datos&nbsp;&nbsp;&nbsp;</B></A></TD>
    <%if Session("OperatorLevel")=0 or Session("OperatorLevel")=1 then %>
		<TD width="1" class=separator vAlign=middle align=right>|</TD>
    	<TD width="50" id=Operators class=menu vAlign=center align=middle>
				<A class=activeMain href="javascript:showMenu('mainOperators', 'Operators', 1)">&nbsp;&nbsp;&nbsp;Administradores&nbsp;&nbsp;&nbsp;</A></A>
		</TD>
    <%end if%>
		<TD width="1" class=separator vAlign=middle align=right>|</TD>
    <TD width="50" class=menu vAlign=center align=middle>
				<A class=activeMain href="javascript:if(%20confirm('Esta%20seguro%20que%20desea%20salir')%20)%20document.location%20='LogOff.asp';">&nbsp;&nbsp;Salir&nbsp;&nbsp;</A>
		</TD>
</TR>
<!--</TABLE>-->

<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup0 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoAWB');"  href="InsertData.asp?GID=1&AT=1&awb_frame2=3" target=principal>Nuevo&nbsp;Export</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoAWB');"  href="InsertData.asp?GID=1&AT=2&awb_frame2=3" target=principal>Nuevo&nbsp;Import</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarAWB');" href="Search_Admin.asp?GID=1" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ReporteConfReserv');" href="Search_Admin.asp?GID=4" target=principal>Confirmaci&oacute;n&nbsp;de&nbsp;Reserva</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ReporteHMan');" href="Search_Admin.asp?GID=6" target=principal>House&nbsp;Cargo&nbsp;Manifiesto</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ReporteArribo');" href="Search_Admin.asp?GID=15" target=principal>Arribo/Salida</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Costos');" href="Search_Admin.asp?GID=17" target=principal>Costos</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Estadisticas');" href="Search_Admin.asp?GID=16" target=principal>Estad&iacute;sticas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
        <TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Rastreo');" href="Search_Admin.asp?GID=18" target=principal>Rastreo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>        
        <TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A <%=IIf(iArr2.Item("21Log") = "1","","onclick='activateItem(this);return false;' style='display:none;'")%> class=activeMain href="InsertData.asp?GID=21&AT=1" target=principal>Guias</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left <%=IIf(iArr2.Item("21Log") = "1","","style='display:none;'")%>>|</TD>
        <TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Mediciones');" href="Search_Admin.asp?GID=22" target=principal>Mediciones</A>&nbsp;&nbsp;</TD>        
		<TD class=submenu vAlign=center width="1000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<%if Session("OperatorLevel")=0 or Session("OperatorLevel")=1 then %>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup1 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="5%" align=right>|</TD>
		<!-- <TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoTransportista');"  href="InsertData.asp?GID=2" target=principal>Nuevo</A>&nbsp;&nbsp;</TD> 
		<TD class=separator vAlign=center align=left>|</TD> -->
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarTransportista');" href="Search_Admin.asp?GID=2" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('AsignarRango');" href="InsertData.asp?GID=5" target=principal>Asignar&nbsp;Rango</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarRango');" href="Search_Admin.asp?GID=5" target=principal>Editar&nbsp;Rango</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('AsignarSalidas');" href="InsertData.asp?GID=3" target=principal>Asignar&nbsp;Salidas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarSalidas');" href="Search_Admin.asp?GID=3" target=principal>Editar&nbsp;Salidas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('AsignarTarifas');" href="CarrierRates.asp" target=principal>Asignar&nbsp;Destinos&nbsp;y&nbsp;Tarifas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="1000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup2 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="16%" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoDestinatario');"  href="InsertData.asp?GID=7" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarDestinatario');" href="Search_Admin.asp?GID=7" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="100%" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup3 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="35%" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoAgente');"  href="InsertData.asp?GID=8" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarAgente');" href="Search_Admin.asp?GID=8" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="100%" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup4 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="40%" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoAeropuerto');"  href="InsertData.asp?GID=9" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarAeropuerto');" href="Search_Admin.asp?GID=9" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup5 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="26%" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoEmbarcador');"  href="InsertData.asp?GID=10" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarEmbarcador');" href="Search_Admin.asp?GID=10" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup6 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="48%" align=right>|</TD>		
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoProducto');"  href="InsertData.asp?GID=11" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarProducto');" href="Search_Admin.asp?GID=11" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup7 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="56%" align=right>|</TD>		
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevaMoneda');"  href="InsertData.asp?GID=12" target=principal>Nueva</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarMoneda');" href="Search_Admin.asp?GID=12" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>

<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup8 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="63%" align=right>|</TD>		
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoRango');"  href="InsertData.asp?GID=13" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarRango');" href="Search_Admin.asp?GID=13" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>


<% if Session("OperatorLevel")=0 then%> 	
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup9 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="69%" align=right>|</TD>		
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Varios');"  href="Setup.asp" target=principal>Varios</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoImpuesto');" href="InsertData.asp?GID=14" target=principal>Nuevo&nbsp;Impuesto</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarImpuesto');" href="Search_Admin.asp?GID=14" target=principal>Editar&nbsp;Impuesto</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="1000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<% end if %>
<%
end if
%>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainMisDatos style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0>
<TR>	
		<TD class=separator vAlign=center align=left>&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>

<%if Session("OperatorLevel")=0 or Session("OperatorLevel")=1 then%>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainOperators style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>
		<TD class=separator vAlign=center width="88%" align=right>|</TD>		
		<TD id=TDOperator1 class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="javascript:activateItem(this);showTitle('NuevoEditor');" href="OPerators.asp" target="principal">Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDOperator2 class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="javascript:activateItem(this);showTitle('EditarEditor');" href="Search_Operators.asp" target="principal">Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<% end if%>
<!--Menu de encargado -->
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainEditores style="LEFT: 25px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto">
<TABLE cellSpacing=0 cellPadding=0 bgcolor=#ff7900>

  <TR>
    <TD class=inactiveMain align=left>&nbsp;<A class=activeMain 
      onmouseover="JavaScript:showMain('mainEditores')" 
      onmouseout="JavaScript:hideMain('mainEditores')" 
      href="#" target=principal 
      target=principal>Nuevo</A> &nbsp;</TD></TR>
  <TR>
    <TD class=inactiveMain align=left>&nbsp;<A class=activeMain 
      onmouseover="JavaScript:showMain('mainEditores')" 
      onmouseout="JavaScript:hideMain('mainEditores')" 
      href="#" target=principal 
      target=principal>Editar</A> &nbsp;</TD></TR>
  <TR>
    <TD class=inactiveMain align=left>&nbsp;<A class=activeMain 
      onmouseover="JavaScript:showMain('mainEditores')" 
      onmouseout="JavaScript:hideMain('mainEditores')" 
      href="#" target=principal 
      target=principal>Eliminar</A> &nbsp;</TD></TR>
</TABLE>
</DIV>
</TD></TR>




<!-- catalogos -->
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup10 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>			
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="http://10.10.1.20/catalogo_admin/login.php" target="principal">Destinatarios</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="http://10.10.1.20/catalogo_admin/login.php" target="principal">Embarcadores</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain href="http://10.10.1.20/catalogo_admin/login.php" target="principal">Agentes</A>&nbsp;&nbsp;</TD>		
</TR>
</TABLE>
</DIV>
</TD></TR>



<!-- Batch Rubros 2017-11-06 -->
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup11 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain <%=IIf(iArr2.Item("23Upd") = "1","onclick='activateItem(this)'","onclick='this.display.style.color=red;return false;' style='color:gray'")%> href="http://<%=IIf(1 = 1, "172.16.0.193/yii/batchrubros", "10.10.1.20/BatchRubros")%>/index.php?r=batch/admin&sis=aereo&usr=<%=Session("Login")%>&ide=<%=Session("OperatorID")%>" target=principal>Administrar</A>&nbsp;&nbsp;</TD>

		<TD class=separator vAlign=center align=left>|</TD>

		<TD class=submenu vAlign=center align=left nowrap>&nbsp;&nbsp;<A class=activeMain <%=IIf(iArr2.Item("23Ins") = "1","onclick='activateItem(this)'","onclick='return false;' style='color:gray'")%> href="http://<%=IIf(1 = 1, "172.16.0.193/yii/batchrubros", "10.10.1.20/BatchRubros")%>/index.php?r=batch/create&tipo_batch=comision&sis=aereo&usr=<%=Session("Login")%>&ide=<%=Session("OperatorID")%>" target=principal>Crear Batch Comisiones</A>&nbsp;&nbsp;</TD>

		<TD class=separator vAlign=center align=left>|</TD>

		<TD class=submenu vAlign=center align=left nowrap>&nbsp;&nbsp;<A class=activeMain <%=IIf(iArr2.Item("23Ins") = "1","onclick='activateItem(this)'","onclick='return false;' style='color:gray'")%> href="http://<%=IIf(1 = 1, "172.16.0.193/yii/batchrubros", "10.10.1.20/BatchRubros")%>/index.php?r=batch/create&tipo_batch=rubros&sis=aereo&usr=<%=Session("Login")%>&ide=<%=Session("OperatorID")%>" target=principal>Crear Batch Rubros</A>&nbsp;&nbsp;</TD>

		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>






</TABLE>

<DIV id=NuevoAWB style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo AWB</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarAWB style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar AWB</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=ReporteHMan style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;House Cargo Manifiesto</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=ReporteArribo style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Notificaci&oacute;n de Arribo</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Costos style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Costos</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Estadisticas style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Estad&iacute;sticas</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=ReporteConfReserv style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Confirmaci&oacute;n de Reservaci&oacute;n a L&iacute;nea A&eacute;rea</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoEmbarcador style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Embarcador</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarEmbarcador style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Embarcador</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoDestinatario style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Destinatario</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarDestinatario style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Destinatario</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoAgente style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Agente</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarAgente style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Agente</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoAeropuerto style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Aeropuerto</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarAeropuerto style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Aeropuerto</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoTransportista style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Transportista</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarTransportista style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Transportista</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=AsignarRango style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Asignar Rango</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarRango style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Rango</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=AsignarSalidas style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Asignar Salidas</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarSalidas style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Salidas</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=AsignarDestinos style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Asignar Destino</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=AsignarTarifas style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Asignar Destinos y Tarifas</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoProducto style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Producto</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarProducto style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Producto</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevaMoneda style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nueva Moneda</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarMoneda style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Moneda</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoRango style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Rango</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Varios style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Varios</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoImpuesto style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Impuesto</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarImpuesto style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Impuesto</TD>
	</TR>
</TABLE>
</DIV>


<DIV id=TextBlank style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0>
  <TR>
		<TD class=title vAlign=center width="100%" align=center>&nbsp;</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoEditor style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Administrador</TD>
	</TR>
</TABLE>
</DIV>


<DIV id=EditarEditor style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Configurar Administrador</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=mainMyData style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Mis Datos</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Rastreo style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Rastreo</TD>
	</TR>
</TABLE>
</DIV>

</BODY></HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
