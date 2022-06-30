<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim Conn, rs, AWBID, CLientID
Dim separator, aTableValues, CountTableValues, i, QuerySelect

	CountTableValues = -1
	AWBID=CheckNum(request("AWBID"))
	ClientID=CheckNum(request("CID"))

    'response.write( "(" & request("CID") & ")")	


    OpenConn Conn

    'response.write( "(" & request("AWBType") & ")")	
    
    'if request("pais") = "GT" and request("AWBType") = 2 then
    'este segmento de codigo ya no se ejecutara desde aca en tracking, se creo un script para ejecutar desde consola GetEstatusCombex.php 2017-04-10 / 11
    if request("pais") = "Guatemala" and request("AWBType") = 2 then 'se agrego para que no entre 2016-12-16

        QuerySelect = "SELECT * FROM Tracking WHERE AWBID = " & AWBID & " AND ClientID = " & ClientID & " AND BLStatusName IN ('FECHA DE INGRESO','FECHA UBICACION')"
	    'response.write QuerySelect & "<br>"	
	    Set rs = Conn.Execute(QuerySelect)
        if rs.EOF then
            Dim ResultText(7)    
            ResultText(0) = "GUIA"
            ResultText(1) = "FECHA DE INGRESO"
            ResultText(2) = "FECHA UBICACION"
            ResultText(3) = "INI REV SAT"
            ResultText(4) = "FECHA FACTURADA"
            ResultText(5) = "EGRESO CARGA"
    
            Dim ImpExp
            if request("AWBType") = 1 then
                ImpExp = "E"
            else
                ImpExp = "I"
            end if

            Dim ResultCode, Tim, Fec

            ResultCode = GetEstatusCombex(request("HAWBNumber"),ImpExp)        
          
            'response.write request("HAWBNumber") & " " & ImpExp & "<br>"	
            'response.write ResultCode & "<br>"	

            If InStr(1,ResultCode, "|") > 0 Then
                ResultCode = Split(ResultCode,"|")
	            for i = 1 TO 5 
                    if ResultCode(i) <> "" then
                        Tim = Split(ResultCode(i)," ")
                        Fec = Tim(0)
                        QuerySelect = "INSERT INTO Tracking (CreatedDate, CreatedTime, AWBID, ClientID, Comment, BLStatusName, DocTyp, OperatorID) " & _
                        "VALUES ('" & Right(Fec,4) & "-" & Mid(Fec,4,2) & "-" & Left(Fec,2) & "','" & Replace(Tim(1),":","") & "00', " & AWBID & ", " & ClientID & ",concat('Sistema Automatico ',now()),'" & ResultText(i) & "', " & request("AWBType") & ", 166)"
                        'response.write QuerySelect & "<br>"	
	                    'Conn.Execute(QuerySelect)
	                end if
	            next
            else
                if ResultCode = 200 then
                    response.write "Respuesta del servidor pero sin informacion"        
                else
                    response.write "Respuesta del servidor " & ResultCode
                end if
            end if

	    end if
	    CloseOBJ rs

    end if

	'response.write BLID & "<br>"
	QuerySelect = "select a.TrackingID, a.CreatedDate, a.CreatedTime, a.Comment, b.FirstName, b.LastName, a.OperatorID, a.BLStatus, a.BLStatusName, a.DocTyp from Tracking a, Operators b " & _
		" where a.AWBID=" & AWBID & " and a.OperatorID=b.OperatorID AND DocTyp = " & request("AWBType") & " order by TrackingID Desc"
	'response.write QuerySelect & "<br>"
                	
    response.write Request.Servervariables("REMOTE_ADDR")

	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount - 1
	end if
	CloseOBJs rs, Conn
%>
<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=2 cellpadding=2 width=100% align=center>
	<TR>
		<TD class=titlelist><b>Fecha</b></TD>
		<TD class=titlelist><b>Estado</b></TD>
		<TD class=titlelist><b>Comentario</b></TD>
		<TD class=titlelist><b>Usuario</b></TD>
	</TR>       
	<%for i=0 to CountTableValues%>
	<TR>
		<TD class=label valign="top"><%=aTableValues(1,i)%><br><%=FormatHour(aTableValues(2,i))%></TD>
		<TD class=label valign="top"><%=aTableValues(8,i)%></TD>
		<TD class=label valign="top"><%=aTableValues(3,i)%></TD>
		<TD class=label valign="top"><%=aTableValues(4,i) & " " & aTableValues(5,i)%>
		<%if aTableValues(6,i) = Session("OperatorID") then%>
		<br><a href="InsertData.asp?OID=<%=aTableValues(0,i)%>&GID=18&CD=<%=ConvertDate(aTableValues(1,i),2)%>&CT=<%=aTableValues(2,i)%>&AT=<%=aTableValues(9,i)%>&Action=98" target="_parent">Editar&nbsp;Comentario</a>
		<%end if%>
		</TD>		
	</TR>
	<TR>
		<TD class=submenu colspan="4"></TD>
	</TR>
	<%next%>  
	</TABLE>
<%Set aTableValues=Nothing%>	
</BODY>
</HTML>