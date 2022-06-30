<script LANGUAGE="VBscript" RUNAT="Server">
Const rsOpenStatic = 3
Const rsLockReadOnly = 1
Const rsClientSide = 3
Const spacer = "&nbsp;&nbsp;"

Const HTMLBase = "http://www.aimargroup.com/"
Const HTMLBaseSSL = "https://www.aimargroup.com/"
Const ConnectionSite = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=DBSite;PWD=siteaimar;DATABASE=db_site"
Const ConnectionLand = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=DbTerrestre_View;PWD=t3rr32tr3a1mar;DATABASE=db_terrestre"
Const ConnectionAir = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=DbAereo_View;PWD=a3r30a1mar;DATABASE=db_aereo"
Const ConnectionMaster = "Driver={PostgreSQL ANSI(x64)};SERVER=10.10.1.20;UID=user_master;PWD=20l0c0n2UlTa2;DATABASE=master-aimar;Fetch=50000;"'PORT=5432
Const ConnectionWMS = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=user_VieWMS;PWD=v13W0nly2;DATABASE=WMS_AIMAR"
Const ConnectionCustomer = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=view_customer;PWD=v3RCu2t0m3R;DATABASE=customer"


'Const HTMLBase = "http://localhost:4040/site/"
'Const HTMLBaseSSL = "http://localhost:4040/site/"
'Const ConnectionSite = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=DBSite;PWD=siteaimar;DATABASE=db_site"
'Const ConnectionLand = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=DbTerrestre_View;PWD=t3rr32tr3a1mar;DATABASE=db_terrestre"
'Const ConnectionAir = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=DbAereo_View;PWD=a3r30a1mar;DATABASE=db_aereo"
'Const ConnectionMaster = "Driver={PostgreSQL 64-Bit ODBC Drivers};SERVER=10.10.1.20;UID=user_master;PWD=20l0c0n2UlTa2;DATABASE=master-aimar;Fetch=50000;"'PORT=5432
'Const ConnectionWMS = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=user_VieWMS;PWD=v13W0nly2;DATABASE=WMS_AIMAR"
'Const ConnectionCustomer = "Driver={MySQL ODBC 5.1 Driver};SERVER=10.10.1.18;UID=view_customer;PWD=v3RCu2t0m3R;DATABASE=customer"


Const PtrnCountries = "'GT'|'SV'|'HN'|'NI'|'CR'|'PA'|'BZ'|'PE'|'EC'|'CO'|'GT2'|'SV2'|'HN2'|'NI2'|'CR2'|'PA2'|'BZ2'"

Sub openConnSite(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (ConnectionSite)
End Sub

Sub openConnMaster(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (ConnectionMaster)
End Sub

Sub openConnWMS(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (ConnectionWMS)
End Sub

Sub openConnLand(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (ConnectionLand)
End Sub

Sub openConnAir(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (ConnectionAir)
End Sub

Sub openConnCustomer(objConn)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open (ConnectionCustomer)
End Sub

Sub openConnOcean(objConn, DB)
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.ConnectionTimeout = 10 ' Seconds
    objConn.CommandTimeout = 10 ' Seconds
    objConn.CursorLocation = rsClientSide
	objConn.Open ("Driver={PostgreSQL ANSI(x64)};SERVER=10.10.1.20;UID=user_vtas;PWD=c0N2ulTas2vtAs;PROTOCOL=2;DATABASE=" & DB & ";Fetch=50000;")
    'objConn.Open ("Driver={PostgreSQL 64-Bit ODBC Drivers};SERVER=10.10.1.20;UID=user_vtas;PWD=c0N2ulTas2vtAs;PROTOCOL=2;DATABASE=" & DB & ";Fetch=50000;")
End Sub
 
Sub closeOBJ(theOBJ)
    On Error Resume Next
    theOBJ.Close
    Set theOBJ = Nothing
End Sub

Sub openTable(Conn, szTable, rs)
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open szTable, Conn, 2, 3, 2
End Sub

Sub SaveData(rs, Action, DataToInsert)
Dim CountElements, i
	CountElements = UBound(DataToInsert, 1) - 1
	if Action =1 then
    	 rs.AddNew
	end if
    For i = 0 To CountElements
        rs(DataToInsert(i)) = DataToInsert(i + 1)
		i = i + 1
    Next
    rs.Update
End Sub

Sub closeOBJs(theOBJ1, theOBJ2)
    On Error Resume Next
    theOBJ1.Close
    theOBJ2.Close
    Set theOBJ1 = Nothing
    Set theOBJ2 = Nothing
End Sub

Function FRegExp(patrn, string1, string2, Action)
Dim regEx ' Create variable.
   If IsNull(string1) Then
    string1 = ""
   End If
   
   If IsNull(string2) Then
    string2 = ""
   End If

   Set regEx = New RegExp ' Create a regular expression.
        regEx.Pattern = patrn   ' Set pattern.
        regEx.IgnoreCase = True   ' Set case insensitivity.
        regEx.Global = True   ' Set global applicability.
   Select Case Action
   Case 1
        Set FRegExp = regEx.Execute(string1)
   Case 2
        FRegExp = regEx.test(string1)
   Case 3
        regEx.Global = False 'Not Set global applicability.
        FRegExp = regEx.Replace(string1, string2)   ' Make replacement.
   Case 4
        regEx.Global = True ' Set global applicability.
        FRegExp = regEx.Replace(string1, string2)   ' Make replacement.
   End Select	 
End Function

function printDec(text,dec)

    printDec = formatnumber(int(CDbl(text) * 100) / 100, dec)

End Function

Function CheckNum(Data)
CheckNum = 0
If InStr(1, Data, " ") = 0 Then
    If FRegExp("[0-9]*", Data, "", 2) Then
         On Error Resume Next
		CheckNum = CDbl(Data)
    End If
End If
End Function

Function CheckTxt(Data)
CheckTxt = ""
If InStr(1, Data, " ") = 0 Then
    If FRegExp("[a-zA-Z0-9]", Data, "", 2) Then
         CheckTxt = CStr(Data)
    End If
End If
End Function

Sub InsertData(rs, Elements)
Dim CountElements, CantElements, Val, Val2
CountElements = UBound(Elements, 1) - 1
'        rs.AddNew
        For i = 0 To CountElements
        Val = Elements(i)
        i = i + 1
        Val2 = Elements(i)
        Next
 '       rs.Update
End Sub

Sub GetTableData(GroupID, ByRef TableName, ByRef ObjectName, ByRef QuerySelect)
    Select Case GroupID 'Tipo de Grupo = 1 Categoria, 2 Producto, 3 Mensaje, 4 User, 5 Noticia
        Case 1, 2
            'QuerySelect = "select CID, CreatedDate, CreatedTime, Expired, Title, Content, Image, Image2, Typ, Pos, Ord, Grp, Countries, SubTitle, Template, Tag, LanguageID from "
			QuerySelect = "select * from "
			TableName = "Content"
            ObjectName = "CID"
		Case 3
            QuerySelect = "select SurveyID, CreatedDate, CreatedTime, Expired, Name, Countries from "
			TableName = "Surveys"
            ObjectName = "SurveyID"
		Case 4, 6
			QuerySelect = "select Name from Surveys where SurveyID="
			TableName = "Questions"
            ObjectName = "QuestionID"
		Case 5, 7
			QuerySelect = "select Name from Surveys where SurveyID="
			TableName = "Answers"
            ObjectName = "AnswerID"
		Case 8
            QuerySelect = "select CampaignID, CreatedDate, CreatedTime, Expired, Name, Inventory, TotImpress, StartDate, FinishDate, Exclusive, Priority, Countries, TotClicks from "
			TableName = "Campaigns"
            ObjectName = "CampaignID"
		Case 9
            QuerySelect = "select CampaignID, Expired, Name from "
			TableName = "Campaigns"
            ObjectName = "CampaignID"
		Case 10
            QuerySelect = "select ZoneID, CreatedDate, CreatedTime, Expired, Name, Description, Height, Width from "
			TableName = "Zones"
            ObjectName = "ZoneID"
		Case 12
			QuerySelect = "select acceso_id, fecha_creacion, hora_creacion, usr, pwd, es_shipper, es_consigneer, es_agent, es_interno, id_cliente, agente_id, id_usuario from "
       	    TableName = "accesos"
			Typ = CheckNum(Request("Typ"))
			Select Case Typ
			Case 1
    	        ObjectName = "id_cliente"
			Case 2
				ObjectName = "agente_id"
			Case 3
				ObjectName = "id_usuario"
			End Select
		Case 13
            QuerySelect = "select ServiceID, CreatedDate, CreatedTime, Expired, Name, URL, Ord, Typ, Height, Width from "
			TableName = "Services"
            ObjectName = "ServiceID"
		Case 14
			QuerySelect = "select id_grupo, fecha_creacion, hora_creacion, id_estatus, nombre_grupo from "
            TableName = "grupos"
            ObjectName = "id_grupo"
		Case 15
			QuerySelect = "select LanguageID, CreatedDate, CreatedTime, Expired, Name, ISOCode, Def from "
            TableName = "Languages"
            ObjectName = "LanguageID"
		Case 16
			QuerySelect = "select * from "
            TableName = "Links"
            ObjectName = "LinkID"
    End Select
End Sub

Function SetOn(Val, Action)
	select case Action
	Case 1
		if Val = "on" then
			SetOn = 0
		else
			SetOn = 1
		end if	
	Case 2
		if Val = "on" then
			SetOn = 1
		else
			SetOn = 0
		end if	
	End Select
End Function

Sub SaveInfo(Conn, rs, Action, GroupID, CreatedDate, CreatedTime)
Dim Grp, rs2, Usr2, Servs, CountServs, AccessID

If Action = 1 Then
    rs.AddNew
	select case GroupID
	case 12, 14
		rs("fecha_creacion") = CreatedDate 
	case else
		rs("CreatedDate") = CreatedDate 
	end select
End If

Select Case GroupID
Case 1,2 'Menu y Contenido
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	rs("Image") = Request.Form("Image")
	rs("Image2") = Request.Form("Image2")
	rs("Typ") = CheckNum(Request.Form("Typ"))
	rs("Pos") = CheckNum(Request.Form("Pos"))
	rs("Ord") = CheckNum(Request.Form("Ord"))
	Grp = CheckNum(Request.Form("Grp"))
	if Grp < 0 then
		rs("Grp") = 0
	else
		rs("Grp") = Grp
	end if
	rs("Countries") = Request.Form("Countries")
	rs("Template") = Request.Form("Template")

	Set rs2 = Conn.Execute("select ISOCode from Languages order by LanguageID")
	do while Not rs2.EOF
		rs("Title"&rs2(0)) = Request.Form("Title"&rs2(0))
		rs("Content"&rs2(0)) = Request.Form("Content"&rs2(0))
		rs("SubTitle"&rs2(0)) = Request.Form("SubTitle"&rs2(0))
		rs("Tag"&rs2(0)) = Request.Form("Tag"&rs2(0))
		rs("MetaTag"&rs2(0)) = Request.Form("MetaTag"&rs2(0))
		rs2.MoveNext
	loop
	CloseOBJ rs2
Case 3 'Encuestas
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	rs("Name") = Request.Form("Name")
	rs("Countries") = Request.Form("Countries")
Case 6 'Preguntas
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	rs("SurveyID") = CheckNum(Request.Form("SID"))
	rs("Question") = Request.Form("Question")
	rs("QType") = CheckNum(Request.Form("QType"))
Case 7 'Respuestas
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	rs("QuestionID") = Request.Form("QID")
	rs("Answer") = Request.Form("Answer")
Case 8 'Campaigns
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	rs("Name") = Request.Form("Name")
	rs("Inventory") = CheckNum(Request.Form("Inventory"))
	rs("TotImpress") = CheckNum(Request.Form("TotImpress"))
	rs("StartDate") = ConvertDate(Request.Form("StartDate"),1)
	rs("FinishDate") = ConvertDate(Request.Form("FinishDate"),1)
	rs("Exclusive") = CheckNum(Request.Form("Exclusive"))
	rs("Priority") = CheckNum(Request.Form("Priority"))
	rs("Countries") = Request.Form("Countries")
Case 9 'Segmentation
	Dim Zones, CountZones
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	SetAsignations ObjectID, "Segmentations", "CampaignID", "ZoneID", "Zones"
Case 10 'Zonas
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	rs("Name") = Request.Form("Name")
	rs("Description") = Request.Form("Description")
	rs("Height") = CheckNum(Request.Form("Height"))
	rs("Width") = CheckNum(Request.Form("Width"))
Case 12
	Usr2 = PurgeData(Request.Form("Usr"))
	AccessID = CheckNum(Request.Form("AccessID"))
	Set rs2=Conn.Execute("select Usr, acceso_id from accesos where Usr='" & Usr2 & "'")
	if rs2.EOF then
		rs("hora_creacion") = CreatedTime 
		rs("id_cliente") = CheckNum(Request.Form("UID"))
		rs("agente_id") = CheckNum(Request.Form("AID"))
		rs("id_usuario") = CheckNum(Request.Form("EID"))
		rs("usr") = Usr2
		rs("pwd") = PurgeData(Request.Form("Pwd"))
		rs("es_shipper") =  SetOn(Request.Form("isShipper"),2)
		rs("es_consigneer") =  SetOn(Request.Form("isConsigneer"),2)	
		rs("es_agent") =  SetOn(Request.Form("isAgent"),2)	
		rs("es_interno") =  SetOn(Request.Form("isEmployee"),2)	
	else
		if CheckNum(rs2(1))=AccessID then
			rs("hora_creacion") = CreatedTime 
			rs("id_cliente") = CheckNum(Request.Form("UID"))
			rs("agente_id") = CheckNum(Request.Form("AID"))
			rs("id_usuario") = CheckNum(Request.Form("EID"))
			rs("pwd") = PurgeData(Request.Form("Pwd"))
			rs("es_shipper") =  SetOn(Request.Form("isShipper"),2)
			rs("es_consigneer") =  SetOn(Request.Form("isConsigneer"),2)	
			rs("es_agent") =  SetOn(Request.Form("isAgent"),2)
			rs("es_interno") =  SetOn(Request.Form("isEmployee"),2)
		else
			Action = 0
			JavaMsg = "El Usuario ya esta asignado a otro cliente"
		end if
	end if
	CloseOBJ rs2	
Case 13 'Zonas
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	rs("Name") = Request.Form("Name")
	rs("URL") = Request.Form("URL")
	rs("Ord") = CheckNum(Request.Form("Ord"))
	rs("Typ") = CheckNum(Request.Form("Typ"))
	rs("Height") = CheckNum(Request.Form("Height"))
	rs("Width") = CheckNum(Request.Form("Width"))
Case 14
	rs("hora_creacion") = CreatedTime 
	rs("id_estatus") =  CheckNum(Request.Form("Expired"))
	rs("nombre_grupo") =  Request.Form("Name")
Case 15
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	rs("Name") = Request.Form("Name")
	rs("ISOCode") = UCase(Request.Form("ISOCode"))
	rs("Def") = SetOn(Request.Form("Def"),2)
Case 16
	rs("CreatedTime") = CreatedTime
	rs("Expired") = SetOn(Request.Form("Expired"),1)
	rs("Name") = Request.Form("Name")
	Set rs2 = Conn.Execute("select ISOCode from Languages order by LanguageID")
	do while Not rs2.EOF
		rs("Title"&rs2(0)) = Request.Form("Title"&rs2(0))
		rs2.MoveNext
	loop
	CloseOBJ rs2
End Select
rs.Update
End Sub

Sub SetAsignations (ObjectID, TableName, ObjectName1, ObjectName2, RequestName)
	Dim List, CountList, Conn
	openConnSite Conn
	Conn.Execute("delete from " & TableName & " where " & ObjectName1 & "=" & ObjectID)
	List = Split(Request.Form(RequestName),"|")
	CountList=UBound(List)
	for i=0 to CountList
		Conn.Execute("insert into " & TableName & " (" & ObjectName1 & ", " & ObjectName2 & ") values (" & ObjectID & ", " & List(i) & ")")
	next
	CloseOBJ Conn
End Sub

Function ConvertDate (Data, Format)
Dim ConvertDay, ConvertMonth, ConvertYear
'response.write Month(Data) & "/" & Day(Data) & "/" & Year(Data) & "<br />"
if Data <> "" then
	 select case Format
	 case 0 'formato dd/mm/yyyy
		 ConvertDate = TwoDigits(Day(Data))& "/" & TwoDigits(Month(Data)) & "/" & Year(Data)		 		 		
	 case 1 'formato dd/mm/yyyy
	   if Day(Data) < 13 then		 		
				 ConvertDate = TwoDigits(Month(Data)) & "/" & TwoDigits(Day(Data)) & "/" & Year(Data)		 		 		
		 else
				 ConvertDate = TwoDigits(Day(Data)) & "/" & TwoDigits(Month(Data)) & "/" & Year(Data)		 		
		 end if
	 case 2 'formato yyyy/mm/dd para strings
		 ConvertDate = Year(Data) & "/" & TwoDigits(Month(Data)) & "/" & TwoDigits(Day(Data))
	 case 3 'formato yyyy/mm/dd para dates
		 if Day(Data) < 13 then
		 		ConvertDate = Year(Data) & "/" & TwoDigits(Month(Data)) & "/" & TwoDigits(Day(Data))
		 else 
				ConvertDate = Year(Data) & "/" & TwoDigits(Day(Data)) & "/" & TwoDigits(Month(Data))				
		 end if	
	 case 4 'formato ingles mm/dd/yyyy
		 if Day(Data) < 13 then
				ConvertDate = Day(Data) & "/" & Month(Data) & "/" & Year(Data) 
		 else 
		 		ConvertDate = Month(Data) & "/" & Day(Data) & "/" & Year(Data) 
		 end if	
		 
	 end select
end if
End Function


Function ConvertDate2 (Data, Format)
Dim ConvertDay, ConvertMonth, ConvertYear
'response.write Month(Data) & "/" & Day(Data) & "/" & Year(Data) & "<br />"
select case Format
case 1 'formato dd/mm/yyyy
	   if Day(Data) < 13 then		 		
				ConvertDate = Day(Data) & "/" & Month(Data) & "/" & Year(Data)		 		
		 else
		 		 ConvertDate = Month(Data) & "/" & Day(Data) & "/" & Year(Data) 
		 end if
case 2 'formato yyyy/mm/dd para strings
		 if Day(Data) < 13 then
		 		ConvertDate = Year(Data) & "/" & Day(Data) & "/" & Month(Data)
		 else 
				ConvertDate = Year(Data) & "/" & Month(Data) & "/" & Day(Data)
		 end if	
case 3 'formato yyyy/mm/dd para dates
		 ConvertDate = Year(Data) & "/" & Month(Data) & "/" & Day(Data)
end select			
End Function

Function CreateSearchQuery(QuerySelect, OptionX, ByRef MoreOptions)
	if OptionX <> "" then
			 'if GroupID <> 1 and MoreOptions = 0 then
			 if MoreOptions = 0 then
			 		QuerySelect = QuerySelect & " where "
			 else
					QuerySelect = QuerySelect & " and "
			 end if			 
			 QuerySelect = QuerySelect & OptionX
			 MoreOptions = 1
	end if
end Function

Function DisplayBirthDate(BirthDate, LN)
Dim i, BirthDay, BirthMonth, BirthYear, HTMLSelect, selected
Dim SelectOption

Dim Matchh, Match, Matches
if isDate(BirthDate) then
	 BirthDate = ConvertDate(BirthDate,1)
	 Set Match = FRegExp("([0-9]*)\/([0-9]*)\/([0-9]*)", BirthDate, "", 1)
	 For Each Matchh In Match
     Set Matches = Match(0)
     BirthDay = CInt(Matches.SubMatches(0))
     BirthMonth = CInt(Matches.SubMatches(1))
     BirthYear = CInt(Matches.SubMatches(2))
   Next
else
		BirthDay = ""
		BirthMonth = ""
		BirthYear = ""
end if
'Desplegando los Dias
HTMLSelect = "<select class=label name=BirthDay><option value='00'>" & TranslateName(LN, "Día") & "</option>"
for i = 1 to 31
		selected = ">"
		if i = BirthDay then
			 selected = " selected>"
		end if
		if i < 10 then
			 SelectOption = "0" & i
		else
			 SelectOption = i				  
		end if
		HTMLSelect = HTMLSelect & "<option value='" & SelectOption & "' " & selected & SelectOption & "</option>"
next
HTMLSelect = HTMLSelect & "</select>"

'Desplegando los Meses
HTMLSelect = HTMLSelect & "<select class=label name=BirthMonth><option value='00'>" & TranslateName(LN, "Mes") & "</option>"
for i = 1 to 12
		selected = ">"
		if i = BirthMonth then
			 selected = " selected>"
		end if
		if i < 10 then
			 SelectOption = "0" & i
		else
			 SelectOption = i				  
		end if
		HTMLSelect = HTMLSelect & "<option value='" & SelectOption & "' " & selected & SelectOption & "</option>"
next
'Desplegando el Año
HTMLSelect = HTMLSelect & "</select>" & _
					  "<INPUT name=BirthYear maxlength=4 TYPE=text size=4 class=label value=" & BirthYear & ">"
DisplayBirthDate = HTMLSelect 
end Function

Function PurgeData(byval AllData)
Dim Data
 Data = replace(AllData,"&","&amp;",1,-1)
 'Data = replace(Data,"?","",1,-1)
 Data = replace(Data,">","&gt;",1,-1)
 Data = replace(Data,"<","&lt;",1,-1)
 Data = replace(Data,"'","",1,-1)
 Data = replace(Data,";","",1,-1)
 Data = replace(Data,"%","",1,-1)
 'Data = replace(Data,"'","",1,-1)
 'Data = replace(Data,chr(34),"",1,-1)
 'Data = replace(Data,"=","",1,-1)
 'Data = replace(Data,"|","",1,-1)
 'Data = replace(Data,"^","",1,-1)
 'Data = replace(Data,"$","",1,-1)
 'Data = replace(Data,"Ñ","N",1,-1)
 'Data = replace(Data,"ñ","n",1,-1)
 'Data = replace(Data,"À","A",1,-1)
 'Data = replace(Data,"Á","A",1,-1)
 'Data = replace(Data,"Â","A",1,-1)
 'Data = replace(Data,"Ã","A",1,-1)
 'Data = replace(Data,"Ä","A",1,-1)
 'Data = replace(Data,"Å","A",1,-1)
 'Data = replace(Data,"Æ","",1,-1)
 'Data = replace(Data,"Ç","",1,-1)
 'Data = replace(Data,"È","E",1,-1)
 'Data = replace(Data,"É","E",1,-1)
 'Data = replace(Data,"Ê","E",1,-1)
 'Data = replace(Data,"Ë","E",1,-1)
 'Data = replace(Data,"Ì","I",1,-1)
 'Data = replace(Data,"Í","I",1,-1)
 'Data = replace(Data,"Î","I",1,-1)
 'Data = replace(Data,"Ï","I",1,-1)
 'Data = replace(Data,"Ð","",1,-1)
 'Data = replace(Data,"Ñ","N",1,-1)
 'Data = replace(Data,"Ò","O",1,-1)
 'Data = replace(Data,"Ó","O",1,-1)
 'Data = replace(Data,"Ô","O",1,-1)
 'Data = replace(Data,"Õ","O",1,-1)
 'Data = replace(Data,"Ö","O",1,-1)
 'Data = replace(Data,"×","",1,-1)
 'Data = replace(Data,"Ø","",1,-1)
 'Data = replace(Data,"Ù","U",1,-1)
 'Data = replace(Data,"Ú","U",1,-1)
 'Data = replace(Data,"Û","U",1,-1)
 'Data = replace(Data,"Ü","U",1,-1)
 'Data = replace(Data,"Ý","",1,-1)
 'Data = replace(Data,"Þ","",1,-1)
 'Data = replace(Data,"ß","",1,-1)
 'Data = replace(Data,"à","a",1,-1)
 'Data = replace(Data,"á","a",1,-1)
 'Data = replace(Data,"â","a",1,-1)
 'Data = replace(Data,"ã","a",1,-1)
 'Data = replace(Data,"ä","a",1,-1)
 'Data = replace(Data,"å","a",1,-1)
 'Data = replace(Data,"æ","",1,-1)
 'Data = replace(Data,"ç","",1,-1)
 'Data = replace(Data,"è","e",1,-1)
 'Data = replace(Data,"é","e",1,-1)
 'Data = replace(Data,"ê","e",1,-1)
 'Data = replace(Data,"ë","e",1,-1)
 'Data = replace(Data,"ì","i",1,-1)
 'Data = replace(Data,"í","i",1,-1)
 'Data = replace(Data,"î","i",1,-1)
 'Data = replace(Data,"ï","i",1,-1)
 'Data = replace(Data,"ð","",1,-1)
 'Data = replace(Data,"ñ","n",1,-1)
 'Data = replace(Data,"ò","o",1,-1)
 'Data = replace(Data,"ó","o",1,-1)
 'Data = replace(Data,"ô","o",1,-1)
 'Data = replace(Data,"õ","o",1,-1)
 'Data = replace(Data,"ö","o",1,-1)
 'Data = replace(Data,"÷","",1,-1)
 'Data = replace(Data,"ø","",1,-1)
 'Data = replace(Data,"ù","u",1,-1)
 'Data = replace(Data,"ú","u",1,-1)
 'Data = replace(Data,"û","u",1,-1)
 'Data = replace(Data,"ü","u",1,-1)
 'Data = replace(Data,"ý","",1,-1)
 'Data = replace(Data,"þ","",1,-1)
 'Data = replace(Data,"ÿ","",1,-1) 
 PurgeData = Data
End Function

Sub Checking (OL)
	'Revisando Permisos de los Usuarios
	If Not FRegExp(OL, Session("OperatorLevel"), "", 2) Then
		Response.Redirect "redirect.asp?MS=4"
	end if
End Sub

Function TwoDigits(Val)
	if Val <= 9 then
		 TwoDigits = "0" & Val
	else 
		 TwoDigits = Val
	end if
End Function
	
Sub FormatTime (ByRef CreatedDate, ByRef CreatedTime) 
		If Not isDate(CreatedDate) or Not isNumeric(CreatedTime) then
			 CreatedDate = Date 
			 CreatedTime = Time
			 CreatedTime = Hour(CreatedTime) & TwoDigits(Minute(CreatedTime)) & TwoDigits(Second(CreatedTime))
			 'CreatedDate = Year(CreatedDate) & "/" & Month(CreatedDate) & "/" & Day(CreatedDate) 
		end if
		CreatedDate = ConvertDate(CreatedDate,2)
		'CreatedDate = FormatDateTime(Year(CreatedDate) & "/" & Month(CreatedDate) & "/" & day(CreatedDate))
end sub

Function FormatHour(TheTime)
Dim LenTime
	LenTime = Len(TheTime)
	select Case LenTime
	Case 5
		FormatHour = Left(TheTime,1) & ":" & Mid(TheTime,2,2) & ":" & Right(TheTime,2)
	Case 6
		FormatHour = Left(TheTime,2) & ":" & Mid(TheTime,3,2) & ":" & Right(TheTime,2)
	End Select
End Function

Function NameOfDay (DayValue)
select case DayValue
case 1
		 NameOfDay = "Domingo"
case 2
		 NameOfDay = "Lunes"
case 3
		 NameOfDay = "Martes"
case 4
		 NameOfDay = "Miercoles"
case 5
		 NameOfDay = "Jueves"
case 6
		 NameOfDay = "Viernes"
case 7
		 NameOfDay = "Sabado"
end select
End Function

Function NameOfMonth (MonthValue)
select case MonthValue
case 1
		 NameOfMonth = "Enero"
case 2
		 NameOfMonth = "Febrero"
case 3
		 NameOfMonth = "Marzo"
case 4
		 NameOfMonth = "Abril"
case 5
		 NameOfMonth = "Mayo"
case 6
		 NameOfMonth = "Junio"
case 7
		 NameOfMonth = "Julio"
case 8
		 NameOfMonth = "Agosto"
case 9
		 NameOfMonth = "Septiembre"
case 10
		 NameOfMonth = "Octubre"
case 11
		 NameOfMonth = "Noviembre"
case 12
		 NameOfMonth = "Diciembre"
end select
End Function


Sub DisplaySearchAdminResults (HTMLCode) 
Dim Conn, rs, j, HTMLCode2, HTMLCode3, ListColor
Dim x
select case GroupID 
case 12, 14
	openConnMaster Conn
case else
	openConnSite Conn
end select
'Buscando los archivos que coinciden con el query de Busqueda
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
		'Obteniendo la cantidad de resultados por busqueda
		rs.PageSize = Session("SearchResults")
		'Saltando a la pagina seleccionada
  	  	rs.AbsolutePage = AbsolutePage
		PageCount = rs.PageCount
		
		
		'response.write x) & "<br />"
		'Desplegando los resultados de la pagina
		for i=1 to rs.PageSize
		    CD = ConvertDate(rs(2),2)'Day(rs(3)) & "/" & Month(rs(3)) & "/" & Year(rs(3))
			
			for j = 2 to elements
				 Select case j
				 case 2 'cuando es fecha se le da formato español
						HTMLCode3 = CD & "</a></td>"
				 case 3 'cuando es titulo, nombre o descripcion, se le da formato acotado
						HTMLCode3 = Mid(rs(j),1,50) & "...</a></td>"
				 case elements 'cuando es titulo, nombre o descripcion, se le da formato acotado
						if GroupID=14 or GroupID=12 then
							if rs(elements)=1 then
								HTMLCode3 = "Activo</a></td>"
							else
								HTMLCode3 = "Inactivo</a></td>"
							end if
						else 
							if rs(elements)=0 then
								HTMLCode3 = "Activo</a></td>"
							else
								HTMLCode3 = "Inactivo</a></td>"
							end if
						end if
				 case else
				 		HTMLCode3 = rs(j) & "</a></td>"
				end select
				
				select case GroupID
				case 6, 11
					HTMLCode2 = HTMLCode2 & "<td class=list><a class=labellist href=Stats.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & ">" & HTMLCode3
				case 12
					HTMLCode2 = HTMLCode2 & "<td class=list><a class=labellist href=InsertData.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & "&Typ=" & Typ & ">" & HTMLCode3				
				case else
					HTMLCode2 = HTMLCode2 & "<td class=list><a class=labellist href=InsertData.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & ">" & HTMLCode3				
				End Select
				
				'if GroupID <> 6 and GroupID <>11 then 
				'	HTMLCode2 = HTMLCode2 & "<td class=list><a class=labellist href=InsertData.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & ">" & HTMLCode3
				'else
				'	HTMLCode2 = HTMLCode2 & "<td class=list><a class=labellist href=Stats.asp?OID=" & rs(0) & "&GID=" & GroupID & "&CD=" & CD & "&CT=" & rs(1) & ">" & HTMLCode3
				'end if
			next
			HTMLCode = HTMLCode & "<tr>" & HTMLCode2 & "</tr>"
			HTMLCode2 = ""
			rs.MoveNext
			If rs.EOF Then Exit For 
		  next 
	else
			JavaMsg = "No Hay Resultados para esta búsqueda"
	end if
CloseOBJs rs, Conn
End Sub

Function GetNameUser (UserID, UserType)
Dim Conn, rs, QuerySelect
		select case UserType
		case 0
				 QuerySelect = "select Name from Users where UserID=" & UserID				 
		case 1
				 QuerySelect = "select FirstName, LastName from Operators where OperatorID=" & UserID
		end select
		
		openConnSite Conn
		Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			 select case UserType
			 case 0
				 GetNameUser = rs(0)
		   case 1
				 GetNameUser = rs(0) & " " & rs(1)
			 end select
		end if
		CloseOBJs rs, Conn
End Function

Function FormatClock (TimeVal)
FormatClock = ""
Select Case Len(TimeVal)
Case 5
		 FormatClock = " " & Mid(TimeVal, 1, 1) & ":" & Mid(TimeVal, 2, 2) & ":" & Mid(TimeVal, 4, 2)
Case 6
		 FormatClock = " " & Mid(TimeVal, 1, 2) & ":" & Mid(TimeVal, 3, 2) & ":" & Mid(TimeVal, 5, 2) 
end select
End Function

Function DisplayLogo(Country)
	   select case Country
	   Case "GT", "SV", "HN", "NI", "CR", "PA", "BZ", "CO", "PE", "EC"
		DisplayLogo = "<img src='img/aimar.jpg' border='0' width='208' height='62'/>"
	   Case "GT2", "SV2", "HN2", "NI2", "CR2", "PA2", "BZ2"
	   	DisplayLogo = "<img src='img/craft.bmp' border='0' width='208' height='62'/>"
	   Case Else
	   	DisplayLogo = "<img src='img/aimar.jpg' border='0' width='208' height='62'/>"
	   End Select
End Function

Function DisplayCountries(Country)
Dim MatchCountries, Match, HTML
	Set MatchCountries = FRegExp(PtrnCountries, Session("Countries"),  "", 1)
	For Each Match in MatchCountries
	   select case Match.value
	   Case "'GT'"
	   		HTML = HTML & "<option value='GT'"
			if Country = "GT" then HTML = HTML & " selected" end if
			HTML = HTML & ">Guatemala-Aimar</option>" 
	   Case "'SV'"
	   		HTML = HTML & "<option value='SV'"
			if Country = "SV" then HTML = HTML & " selected" end if
			HTML = HTML & ">El Salvador-Aimar</option>" 
	   Case "'HN'"
	   		HTML = HTML & "<option value='HN'"
			if Country = "HN" then HTML = HTML & " selected" end if
			HTML = HTML & ">Honduras-Aimar</option>" 
	   Case "'NI'"
	   		HTML = HTML & "<option value='NI'"
			if Country = "NI" then HTML = HTML & " selected" end if
			HTML = HTML & ">Nicaragua-Aimar</option>" 
	   Case "'CR'"
	   		HTML = HTML & "<option value='CR'"
			if Country = "CR" then HTML = HTML & " selected" end if
			HTML = HTML & ">Costa Rica-Aimar</option>" 
	   Case "'PA'"
	   		HTML = HTML & "<option value='PA'"
			if Country = "PA" then HTML = HTML & " selected" end if
			HTML = HTML & ">Panama-Aimar</option>" 
	   Case "'BZ'"
	   		HTML = HTML & "<option value='BZ'"
			if Country = "BZ" then HTML = HTML & " selected" end if
			HTML = HTML & ">Belice-Aimar</option>" 
	   Case "'CO'"
	   		HTML = HTML & "<option value='CO'"
			if Country = "CO" then HTML = HTML & " selected" end if
			HTML = HTML & ">Colombia-Aimar</option>" 
	   Case "'EC'"
	   		HTML = HTML & "<option value='EC'"
			if Country = "EC" then HTML = HTML & " selected" end if
			HTML = HTML & ">Ecuador-Aimar</option>" 
	   Case "'PE'"
	   		HTML = HTML & "<option value='PE'"
			if Country = "PE" then HTML = HTML & " selected" end if
			HTML = HTML & ">Peru-Aimar</option>" 
	   Case "'GT2'"
	   		HTML = HTML & "<option value='GT2'"
			if Country = "GT2" then HTML = HTML & " selected" end if
			HTML = HTML & ">Guatemala-Craft</option>" 
	   Case "'SV2'"
	   		HTML = HTML & "<option value='SV2'"
			if Country = "SV2" then HTML = HTML & " selected" end if
			HTML = HTML & ">El Salvador-Craft</option>" 
	   Case "'HN2'"
	   		HTML = HTML & "<option value='HN2'"
			if Country = "HN2" then HTML = HTML & " selected" end if
			HTML = HTML & ">Honduras-Craft</option>" 
	   Case "'NI2'"
	   		HTML = HTML & "<option value='NI2'"
			if Country = "NI2" then HTML = HTML & " selected" end if
			HTML = HTML & ">Nicaragua-Craft</option>" 
	   Case "'CR2'"
	   		HTML = HTML & "<option value='CR2'"
			if Country = "CR2" then HTML = HTML & " selected" end if
			HTML = HTML & ">Costa Rica-Craft</option>" 
	   Case "'PA2'"
	   		HTML = HTML & "<option value='PA2'"
			if Country = "PA2" then HTML = HTML & " selected" end if
			HTML = HTML & ">Panama-Craft</option>" 
	   Case "'BZ2'"
	   		HTML = HTML & "<option value='BZ2'"
			if Country = "BZ2" then HTML = HTML & " selected" end if
			HTML = HTML & ">Belice-Craft</option>" 
	   end select
	Next
	response.Write HTML
End Function

Function ListCountries (CountriesChecked, CountriesAssigned)
Dim MatchCountries, Match, GT, SV, HN, NI, CR, PA, BZ, CO, PE, EC, GT2, SV2, HN2, NI2, CR2, PA2, BZ2
Dim HTML, TRTDB, BTDTDI, I1, I2, TDTR	
	Set MatchCountries = FRegExp(PtrnCountries, CountriesChecked, "", 1)
	For Each Match in MatchCountries
	   select case Match.value
	   Case "'GT'"
	   		GT = "checked"			
	   Case "'SV'"
	   		SV = "checked"
	   Case "'HN'"
	   		HN = "checked"
	   Case "'NI'"
	   		NI = "checked"
	   Case "'CR'"
	   		CR = "checked"
	   Case "'PA'"
	   		PA = "checked"
	   Case "'BZ'"
	   		BZ = "checked"
	   Case "'CO'"
	   		CO = "checked"
	   Case "'EC'"
	   		EC = "checked"
	   Case "'PE'"
	   		PE = "checked"
	   Case "'GT2'"
	   		GT2 = "checked"			
	   Case "'SV2'"
	   		SV2 = "checked"
	   Case "'HN2'"
	   		HN2 = "checked"
	   Case "'NI2'"
	   		NI2 = "checked"
	   Case "'CR2'"
	   		CR2 = "checked"
	   Case "'PA2'"
	   		PA2 = "checked"
	   Case "'BZ2'"
	   		BZ2 = "checked"
	   end select
   Next
   
   
   TRTDB = "<TR><td class=label align=left>"
   BTDTDI = "</TD><td class=label align=left><INPUT name='"
   I1 = "' value='"
   I2 = "' type=checkbox class=label "
   TDTR = "></TD></TR>"
   HTML = "<TABLE align=center>"	   
   Set MatchCountries = FRegExp(PtrnCountries, CountriesAssigned, "", 1)
   For Each Match in MatchCountries
	   select case Match.value
	   Case "'GT'"
			HTML = HTML & TRTDB & "Guatemala-Aimar:" & BTDTDI & "GT" & I1 & "GT" & I2 & GT & TDTR
	   Case "'SV'"
			HTML = HTML & TRTDB & "El Salvador-Aimar:" & BTDTDI & "SV" & I1 & "SV" & I2 & SV & TDTR
	   Case "'HN'"
			HTML = HTML & TRTDB & "Honduras-Aimar:" & BTDTDI & "HN" & I1 & "HN" & I2 &  HN & TDTR
	   Case "'NI'"
			HTML = HTML & TRTDB & "Nicaragua-Aimar:" & BTDTDI & "NI" & I1 & "NI" & I2 & NI & TDTR
	   Case "'CR'"
			HTML = HTML & TRTDB & "Costa Rica-Aimar:" & BTDTDI & "CR" & I1 & "CR" & I2 & CR & TDTR
	   Case "'PA'"
			HTML = HTML & TRTDB & "Panama-Aimar:" & BTDTDI & "PA" & I1 & "PA" & I2 & PA & TDTR
	   Case "'BZ'"
			HTML = HTML & TRTDB & "Belice-Aimar:" & BTDTDI & "BZ" & I1 & "BZ" & I2 & BZ & TDTR
	   Case "'CO'"
			HTML = HTML & TRTDB & "Colombia-Aimar:" & BTDTDI & "CO" & I1 & "CO" & I2 & CO & TDTR
	   Case "'PE'"
			HTML = HTML & TRTDB & "Peru-Aimar:" & BTDTDI & "PE" & I1 & "PE" & I2 & PE & TDTR
	   Case "'EC'"
			HTML = HTML & TRTDB & "Ecuador-Aimar:" & BTDTDI & "EC" & I1 & "EC" & I2 & EC & TDTR
	   Case "'GT2'"
			HTML = HTML & TRTDB & "Guatemala-Craft:" & BTDTDI & "GT2" & I1 & "GT2" & I2 & GT2 & TDTR
	   Case "'SV2'"
			HTML = HTML & TRTDB & "El Salvador-Craft:" & BTDTDI & "SV2" & I1 & "SV2" & I2 & SV2 & TDTR
	   Case "'HN2'"
			HTML = HTML & TRTDB & "Honduras-Craft:" & BTDTDI & "HN2" & I1 & "HN2" & I2 & HN2 & TDTR
	   Case "'NI2'"
			HTML = HTML & TRTDB & "Nicaragua-Craft:" & BTDTDI & "NI2" & I1 & "NI2" & I2 & NI2 & TDTR
	   Case "'CR2'"
			HTML = HTML & TRTDB & "Costa Rica-Craft:" & BTDTDI & "CR2" & I1 & "CR2" & I2 & CR2 & TDTR
	   Case "'PA2'"
			HTML = HTML & TRTDB & "Panama-Craft:" & BTDTDI & "PA2" & I1 & "PA2" & I2 & PA2 & TDTR
	   Case "'BZ2'"
			HTML = HTML & TRTDB & "Belice-Craft:" & BTDTDI & "BZ2" & I1 & "BZ2" & I2 & BZ2 & TDTR
	   end select
	Next
   response.Write HTML & "</TABLE>"
end Function

Function DisplayGroups (ObjectID, Action, Grp, Pos, spacer, Count, Lang) 
Dim rs, aGroupValues, CountGroupValues, Menu, ntr, i, SQLQuery, spac
	ntr = chr(10) & chr(13)
	CountGroupValues = -1

	if Action = 0 then
		SQLQuery = "select CID, Title" & Lang & " from Content where Typ=0 and Expired=0 and Grp=" & Grp & " and Pos=" & Pos & " and CID <> " & ObjectID & " order by Pos, Ord, Title" & Lang
	else
		SQLQuery = "select CID, Title" & Lang & " from Content where Typ=0 and Expired=0 and Grp=" & Grp & " and Pos=" & Pos & " order by Pos, Ord, Title" & Lang
	end if
	
	Count = Count + 1
    For i = 0 To Count
        spac = spac & spacer
    Next
	
	Set rs = Conn.Execute(SQLQuery)
	If Not rs.EOF Then
    	aGroupValues = rs.GetRows
        CountGroupValues = rs.RecordCount - 1
    End If
	CloseOBJ rs
	
	if CountGroupValues >= 0 then
		for i = 0 to CountGroupValues
				Menu = Menu & "<option value='" & aGroupValues(0,i) & "'>" & spac & aGroupValues(1,i) & "</option>" & ntr
				Menu = Menu & DisplayGroups(aGroupValues(0,i), Action, aGroupValues(0,i), 0, spacer, Count, Lang) 
		next
	end if
	Count = Count - 1
	DisplayGroups = Menu
End Function

Function DisplayLangs (Conn, Action) 
	Set rs = Conn.Execute("select LanguageID, Name, ISOCode, Def from Languages where Expired=0 order by Name")
	select Case Action
	case 0
		for i=0 to rs.RecordCount-1
			if rs(3)=1 then
				response.write "<option value=" & rs(0) & " selected>" & rs(1) & "</option>"
			else
				response.write "<option value=" & rs(0) & ">" & rs(1) & "</option>"
			end if
			rs.MoveNext
		next
	case 1
		for i=0 to rs.RecordCount-1
			if rs(3)=1 then
				response.write "<option value=" & rs(2) & " selected>" & rs(1) & "</option>"
			else
				response.write "<option value=" & rs(2) & ">" & rs(1) & "</option>"
			end if
			rs.MoveNext
		next
	end select
End Function


Function TranslateCompany(Country)
   select case Country
   Case "GT"
		TranslateCompany = "Aimar GT"
   Case "SV"
		TranslateCompany = "Aimar SV"
   Case "HN"
		TranslateCompany = "Aimar HN"
   Case "NI"
		TranslateCompany = "Aimar NI"
   Case "CR"
		TranslateCompany = "Aimar CR"
   Case "PA"
		TranslateCompany = "Aimar PA"
   Case "BZ"
		TranslateCompany = "Aimar BZ"
   Case "CO"
		TranslateCompany = "Aimar CO"
   Case "PE"
		TranslateCompany = "Aimar PE"
   Case "EC"
		TranslateCompany = "Aimar EC"
   Case "GT2"
		TranslateCompany = "Craft GT"
   Case "SV2"
		TranslateCompany = "Craft SV"
   Case "HN2"
		TranslateCompany = "Craft HN"
   Case "NI2"
		TranslateCompany = "Craft NI"
   Case "CR2"
		TranslateCompany = "Craft CR"
   Case "PA2"
		TranslateCompany = "Craft PA"
   Case "BZ2"
		TranslateCompany = "Craft BZ"
   end select
End Function

Function GetImagefromType(ImgType, ImgName)
Dim Img
	'response.write StrComp(Trim(ImgType), "gif", 1) & "<br />"
	ImgType = replace(ImgType,chr(13),"",1,-1)
	ImgType = replace(ImgType,chr(10),"",1,-1)
	select case LCase(ImgType) 
	case "jpg", "gif", "bmp", "jpeg", "jfif", "tif", "tiff", "jpe"
			 Img = ImgName
	case "swf"
			 Img = "flashfile.gif"
	case "doc"
			 Img = "docfile.gif"
	case "xls"
			 Img = "xlsfile.gif"			 
	case "ppt"
			 Img = "pptfile.gif"
	case "pdf"
			 Img = "pdffile.gif"
	case "html", "htm"
			 Img = "htmlfile.gif"
	case "exe"
			 Img = "exefile.gif"
	case "txt"
			 Img = "textfile.gif"			 
	case "zip"
			 Img = "zipfile.gif"
	case else
			 Img = "otherfile.gif"		 
	end select
	GetImageFromType = Img
end function

Function GetImageData(Data, Action) 'As MatchCollection
Dim Matchh, Match, Matches, RegExpr
select case Action
case 0
		 RegExpr = "(\.)([A-Za-z0-9 \-]*$)"
case 1
		 RegExpr = "(\\)([A-Za-z0-9 _\-\.]*$)" 
end select

GetImageData = ""
Set Match = FRegExp(RegExpr, Data,  "", 1)
For Each Matchh in Match
      Set Matches = Match(0)'
      GetImageData = Trim(Matches.SubMatches(1)) '& vbCrLf      
Next
End Function

Sub GetSession
Dim Country
	
	if Request.Cookies("Country") = "" then
		SetSession "GT"
	else 
		Country = Request("C")
		if Country <> "" and Country <> Request.Cookies("Country") then
			SetSession Country
		end if
	end if

	SetLang Request("L")

End Sub

Sub SetSession (Country)
Dim rs
	Select Case Country
	Case "BZ"
		Response.Cookies("Country") = "BZ"
	Case "GT"
		Response.Cookies("Country") = "GT"
	Case "SV"
		Response.Cookies("Country") = "SV"
	Case "HN"
		Response.Cookies("Country") = "HN"
	Case "NI"
		Response.Cookies("Country") = "NI"
	Case "CR"
		Response.Cookies("Country") = "CR"
	Case "PA"
		Response.Cookies("Country") = "PA"
	Case "CO"
		Response.Cookies("Country") = "CO"
	Case "PE"
		Response.Cookies("Country") = "PE"
	Case "EC"
		Response.Cookies("Country") = "EC"
	Case Else
		Response.Cookies("Country") = "GT"
	End Select

	'Configurando parametros standar globales
	if Request.Cookies("ClientURL") = "" then
		Set rs = Conn.Execute("select SearchResults, ClientURL, GrpDisplayOrder, CntDisplayOrder, ImagePath, HitCounter, HourDif from Miscellaneous")
		if Not rs.EOF then
			Response.Cookies("SearchResults") = rs(0)
			Response.Cookies("ClientURL") = rs(1)
			Response.Cookies("GrpDisplayOrder") = rs(2)
			Response.Cookies("CntDisplayOrder") = rs(3)
			Response.Cookies("ImagePath") = rs(4)
			Response.Cookies("HitCounter") = rs(5) + 1
			Response.Cookies("HourDif") = rs(6)
		end if
		CloseOBJ rs
		Conn.Execute("update Miscellaneous set HitCounter=" & CheckNum(Request.Cookies("HitCounter")))
	end if
End Sub

Sub SetLang (Lang)
Dim rs, Langs
	'Configurando parametros standar globales
	if Request.Cookies("Lang") = "" then
		'Response.Cookies("CountLangs") = 0

		Set rs = Conn.Execute("select Name, ISOCode, Def from Languages where Expired=0 order by Name")
		Do While Not rs.EOF
			Langs = Langs & "<option value='"& rs(1) &"'>" & rs(0) & "</option>"
			if rs(2)=1 then 'Obteniendo el Lenguaje Default
				Response.Cookies("Lang") = rs(1)
			end if
			rs.MoveNext
		Loop
		Response.Cookies("CountLangs") = rs.RecordCount-1
		CloseOBJ rs
		Response.Cookies("Langs") = Langs
	else
		if Lang <> "" then
			Response.Cookies("Lang") = Lang
		end if
	end if
End Sub

Function SetOrder (Val, Lang)
	if CheckNum(Val) = 0 then 'Configurando el Orden de presentacion de los grupos
		SetOrder = "Title" & Lang
	else
		SetOrder = "Ord"
	end if
End Function

Sub SetMenu (Pos)
Dim rs, aMenuValues, CountMenuValues, i, Menu, com, Order, StyleC, StyleG, Lang
	CountMenuValues = -1
	com = chr(34)

	Lang = Request.Cookies("Lang")
	Order = SetOrder(Request.Cookies("GrpDisplayOrder"), Lang)

	StyleC = "G-" & CID & "-" & Pos
	StyleG = "G-" & Grp & "-" & Pos
	'Configurando el Menu
	if Request.Cookies(StyleC) = "" and Lang <> "" then
		'response.Write "Primer Ingreso de Contenido<br />"
		Set rs = Conn.Execute("select CID, Title" & Lang & ", Grp, SubTitle" & Lang & " from Content where Typ=0 and Expired=0 and Pos=" & Pos & " and Grp=" & CID & " and Countries like " & com & "'%" & Request.Cookies("Country") & "%'" & com & " order by " & Order)

        'response.write "select CID, Title" & Lang & ", Grp, SubTitle" & Lang & " from Content where Typ=0 and Expired=0 and Pos=" & Pos & " and Grp=" & CID & " and Countries like " & com & "'%" & Request.Cookies("Country") & "%'" & com & " order by " & Order

		if rs.EOF then
			Response.Cookies(StyleC) = "X"

			if Request.Cookies(StyleG) = "" then
				'Response.write "Se creara el Grupo<br />"
				CloseOBJ rs
				Set rs = Conn.Execute("select CID, Title" & Lang & ", Grp, SubTitle" & Lang & " from Content where Typ=0 and Expired=0 and Pos=" & Pos & " and Grp=" & Grp & " and Countries like " & com & "'%" & Request.Cookies("Country") & "%'" & com & " order by " & Order)

				if rs.EOF then
					Response.Cookies(StyleG) = "X"
					'response.write "se busca el default<br />"
					
					if Request.Cookies("G-0-" & Pos) = "" then
						'response.write "Se crea el Default<br />"
						CloseOBJ rs
						Set rs = Conn.Execute("select CID, Title" & Lang & ", Grp, SubTitle" & Lang & " from Content where Typ=0 and Expired=0 and Pos=" & Pos & " and Grp=0 and Countries like " & com & "'%" & Request.Cookies("Country") & "%'" & com & " order by " & Order)
						StyleG = "G-0-" & Pos
					end if
				end if
			end if
		end if
		
		if Not rs.EOF then
			aMenuValues = rs.GetRows
			CountMenuValues = rs.RecordCount-1
		end if		
		CloseOBJ rs

		if CountMenuValues >= 0 then
			if Pos = 0 then 'Vertical
				For i=0 to CountMenuValues
				'	Menu = Menu & "<tr>" & _
				'	"<td colspan='2'><a href='/" & LCase(Replace(Replace(aMenuValues(3, i),".","-",1,-1)," ","-",1,-1)) & "/' class='link1m'>" & _
				'	"<div class='link1m' onmouseover=" & com & "this.className='link1over';" & com & " onmouseout=" & com & "this.className='link1m';" & com & ">&nbsp;&bull;&nbsp;" & _
				'	aMenuValues(1, i) & "</div></a></td></tr>"
				Menu = Menu & "<tr><td colspan='2' class='link1m' onmouseover=" & com & "this.className='link1over';" & com & " onmouseout=" & com & "this.className='link1m';" & com & " onclick=" & com & "location.href='" & HTMLBase & LCase(Replace(Replace(aMenuValues(3, i),".","-",1,-1)," ","-",1,-1)) & "/';" & com & ">&nbsp;&bull;&nbsp;" & _
					aMenuValues(1, i) & "</td></tr><tr><td colspan='2' bgcolor='#FFFFFF'><img src='img/spacer.gif' width='1' height='1' alt='' /></td></tr>"
				next
			else 'Horizontal
				For i=0 to CountMenuValues
				Menu = Menu & "<td align='center' class='link1h' onmouseover=" & com & "this.className='link1hover';" & com & " onmouseout=" & com & "this.className='link1h';" & com & " onclick=" & com & "location.href='" & HTMLBase & LCase(Replace(Replace(aMenuValues(3, i),".","-",1,-1)," ","-",1,-1)) & "/';" & com & ">" & _
					aMenuValues(1, i) & "</td><td bgcolor='#B6AFAF'><img src='img/spacer.gif' width='1' height='1' alt='' /></td>"					
				next
			end if
			
			if Request.Cookies(StyleC) = "X" then
				Response.Cookies(StyleG) = Menu
			else
				Response.Cookies(StyleC) = Menu
			end if
			'response.Write styleC & "<br />" & styleG & "<br />"
		end if
	end if	
End Sub

Function DisplayMenu (Pos)
	if Request.Cookies("G-" & CID & "-" & Pos) = "X" then

		if Request.Cookies("G-" & Grp & "-" & Pos) = "X" then
			'response.Write "G-" & Grp & "-" & Pos & "DEFAULT<br />"
			Response.Write Request.Cookies("G-0-" & Pos)
		else
			'response.Write "G-" & Grp & "-" & Pos & "<br />"
			Response.Write Request.Cookies("G-" & Grp & "-" & Pos)
		end if

	else
		'response.Write "G-" & CID & "-" & Pos & "<br />"
		Response.Write Request.Cookies("G-" & CID & "-" & Pos)
	end if				
End Function

Sub SendMail(Message, eMails, Subject, FromAddress)
Dim Mailer
Dim iConf
Dim Pass

    Select Case FromAddress
    Case "sales-gt@aimargroup.com"
        Pass = "sal3s@gta!"
	Case "sales-sv@aimargroup.com"
        Pass = "sal3s@sva!"
	Case "sales-hn@aimargroup.com"
        Pass = "sal3s@hna!"
	Case "sales-nic@aimargroup.com"
        Pass = "sal3s@nica!"
	Case "sales-cr@aimargroup.com"
        Pass = "sal3s@cra!"
	Case "sales-bz@aimargroup.com"
        Pass = "sal3s@bza!"
	Case "sales-pa@aimargroup.com"
        Pass = "sal3s@pa!"
    Case Else
        FromAddress = "tracking@aimargroup.com"
        Pass = "Lq14@6A8"
    End Select

    Set Mailer = CreateObject("CDO.Message")
	Set iConf = Mailer.Configuration
	With iConf.Fields
	    .item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
		.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.aimargroup.com"
		.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(25)
		.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = FromAddress
		.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Pass
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 0
		.Update
	End With
    Mailer.From = FromAddress
    Mailer.To = eMails
    Mailer.Subject = Subject
    Mailer.HTMLBody = Message
    Mailer.Send
	Set Mailer = Nothing
	Set iConf = Nothing
End Sub

'Sub SendMail(Message, eMails, Subject, FromName, FromAddress)
'Dim Mailer
'Dim iConf
'Const cdoSendUsingPickup = 1
'Const strPickup = "C:\Inetpub\mailroot\Pickup"
'
'Set Mailer = CreateObject("CDO.Message")
'	
'	Set iConf = Mailer.Configuration
'	With iConf.Fields
'	  	.item("http://schemas.microsoft.com/cdo/configuration/sendusing")= cdoSendUsingPickup
'	    .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strPickup
'	    .Update
'	End With
'
'    Mailer.From = FromAddress
'    Mailer.To = eMails
'    Mailer.Subject = Subject
'    Mailer.HTMLBody = Message
'    Mailer.Send
'Set Mailer = Nothing
'End Sub


'Sub SendMail(Message, eMails, Subject, FromName, FromAddress)
'Dim Mailer
'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'		Mailer.Charset = 2 'Utilizamos UNICODE
'		Mailer.ContentType = "text/html"
'		Mailer.FromName = FromName
'		Mailer.FromAddress = FromAddress 
'		Mailer.Subject = Subject
'		Mailer.BodyText = Message
'		Mailer.RemoteHost = IP_SMTP'"mail-fwd.oemgrp.com"
'		Mailer.AddRecipient "", eMails
'		Mailer.SendMail
'Set Mailer = Nothing

'Set Mailer = Server.CreateObject("CDONTS.NewMail")
'Set Mailer = Server.CreateObject("CDO.Message")
'    Mailer.From = FromName
'    Mailer.To = eMails
'    Mailer.Subject = Subject
'    Mailer.HTMLBody = Message
'    Mailer.Send
'Set Mailer = Nothing
'End Sub

Sub GetData(Action)
Dim SQLQuery, rs, com, Lang
	com = chr(34)
	Lang = Request.Cookies("Lang")

	if Lang <> "" then
		Select Case Action
		Case 0
			SQLQuery = "select Title" & Lang & ", SubTitle" & Lang & ", Content" & Lang & ", Image, Image2, Template, Tag" & Lang & ", MetaTag" & Lang & " from Content where CID=(select max(CID) from Content where Typ=0 and Template=1)"
		Case 1
			SQLQuery = "select Title" & Lang & ", SubTitle" & Lang & ", Content" & Lang & ", Image, Image2, Template, Tag" & Lang & ", MetaTag" & Lang & " from Content where CID=" & CID
		Case 2
			SQLQuery = "select Title" & Lang & ", SubTitle" & Lang & ", Content" & Lang & ", Image, Image2, Template, Tag" & Lang & ", MetaTag" & Lang & " from Content where Grp=" & CID & " and Countries like " & com & "'%" & Request.Cookies("Country") & "%'" & com & " order by " & SetOrder(Request.Cookies("GrpDisplayOrder"), Lang)
		End Select
		'response.Write SQLQuery
		
		Set rs = Conn.Execute(SQLQuery)
		if Not rs.EOF then
			Title = rs(0) 
			SubTitle = rs(1)
			Content = rs(2) 
			Image = rs(3) 
			Image2 = rs(4) 
			Template = rs(5)
			Tag = rs(6)
			MetaTag = rs(7)
		else
			Title = ""
			SubTitle = ""
			Content = ""
			Image = ""
			Image2 = ""
			Template = 0 
			Tag = ""
			MetaTag = ""
		end if
		CloseOBJ rs
	end if
End Sub

Sub GetListData()
Dim SQLQuery, rs, com, Lang
	com = chr(34)
	Lang = Request.Cookies("Lang")
	if Lang <> "" then
		if CID = 0 then
			SQLQuery = "select CID, Title" & Lang & ", Content" & Lang & ", Image, Grp, SubTitle" & Lang & " from Content where Typ=1 and Grp=(select max(CID) from Content where Typ=0 and Template=1) and Countries like " & com & "'%" & Request.Cookies("Country") & "%'" & com & " order by " & SetOrder(Request.Cookies("CntDisplayOrder"), Lang)
		else
			SQLQuery = "select CID, Title" & Lang & ", Content" & Lang & ", Image, Grp, SubTitle" & Lang & " from Content where Typ=1 and Grp=" & CID & " and Countries like " & com & "'%" & Request.Cookies("Country") & "%'" & com & " order by " & SetOrder(Request.Cookies("CntDisplayOrder"), Lang)
		end if
	
		Set rs = Conn.Execute(SQLQuery)
		if Not rs.EOF then
			ListData = rs.GetRows
			CountListData = rs.RecordCount - 1
		end if
		CloseOBJ rs
	end if	
End Sub

Function ParseCountries
Dim MatchCountries, Match, SQLQuery
Set MatchCountries = FRegExp(PtrnCountries, Session("Countries"),  "", 1)
	For Each Match in MatchCountries
		if i = 0 then
				SQLQuery = " a.Countries like ""%" & Match.value & "%"" "
		else
				SQLQuery = SQLQuery & "or a.Countries like ""%" & Match.value & "%"" "
		end if
		i=1 
	Next
	ParseCountries = "(" & SQLQuery & ")"
End Function

Function DisplaySurvey (OID)
Dim aTableValues, CountTableValues, aTable2Values, CountTable2Values, Vals, rs, i, j
Dim union, Qtype, SQLQuery1, SQLQuery2, HTMLFinal, MaxResult
	CountTableValues = -1
	
	if OID = 0 then 'Recibe Respuestas
		SQLQuery1 = "select a.QuestionID, a.Question, a.QType from Questions a where a.Expired=0 and a.SurveyID=(select max(b.SurveyID) from Surveys b where b.Expired=0 and b.Countries like '%" & Request.Cookies("Country") & "%')"
		SQLQuery2 = ""
		HTMLFinal = "<tr><td 'colspan=2' align='center'><input class='input1' type='submit' value='ingresar'></td></tr></form></table>"
	else 'Presenta Respuestas de encuesta indicada en ObjectID
		SQLQuery1 = "select a.QuestionID, a.Question from Questions a where a.SurveyID=" & OID
		SQLQuery2 = " order by Counter Desc"
		HTMLFinal = "</form></table>"
	end if
	
	Set rs=Conn.Execute(SQLQuery1)
	If Not rs.EOF Then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	End If

	union = ""
	
	if CountTableValues>=0 then
		
		DisplaySurvey = "<form method=post action=setsurvey.asp name=formsurvey target=_blank><table cellspacing=0 cellpadding=0 align=center>"
		
		for i=0 to CountTableValues
			Vals = Vals & union & aTableValues(0,i)
			union = ", "
			CountSurvey = CountSurvey+1
		next
		Set rs=Conn.Execute("select QuestionID, AnswerID, Answer, Counter from Answers where Expired=0 and QuestionID in (" & Vals & ")" & SQLQuery2)
		If Not rs.EOF Then
			aTable2Values = rs.GetRows
			CountTable2Values = rs.RecordCount-1

			MaxResult = aTable2Values(3,0)'Para la graficacion
			if MaxResult <=0 then
				MaxResult = 1
			end if
		End If
		CloseOBJ rs

		for i=0 to CountTableValues
			'response.write aTableValues(1,i) & "<br />"
			DisplaySurvey = DisplaySurvey & "<tr><td class='contenido1' colspan='2' align='left'><b>" & aTableValues(1,i) & "</b></td></tr>"
			
			if OID = 0 then
				if aTableValues(2,i)=1 then
					QType = "checkbox"
				else
					QType = "radio"
				end if
				for j=0 to CountTable2Values
					if aTableValues(0,i) = aTable2Values(0,j) then
						DisplaySurvey = DisplaySurvey & "<tr><td class='contenido1' width='1%' align='left'><input class='link1' type='" & QType & "' name='" & aTable2Values(0,j) & " value='" & aTable2Values(1,j) & "'></td><td class='contenido1' align='left'>" & aTable2Values(2,j) & "</td></tr>"
						CountSurvey = CountSurvey+1
					end if
				next
			else
				for j=0 to CountTable2Values
					if aTableValues(0,i) = aTable2Values(0,j) then
						W = aTable2Values(3,j)*250/MaxResult
						DisplaySurvey = DisplaySurvey & "<tr><td class='contenido1' width='20%' align='left'>" & aTable2Values(2,j) & "</td><td class='contenido1' align='left'><img src='img/begin.jpg' border='0' /><img src='img/middle.jpg' border='0' width='" & W & "' height='14 '/><img src='img/end.jpg' border='0' />" & aTable2Values(3,j) & "</td></tr>"
					end if
				next
			end if
		next
		Set aTableValues = Nothing
		Set aTable2Values = Nothing
		
		DisplaySurvey = DisplaySurvey & HTMLFinal
	else 
		DisplaySurvey = ""
	end if
End Function

Function TranslateCountry (Country)
Select Case Country
	Case "AF" TranslateCountry = "AFGANISTAN"
	Case "AL" TranslateCountry = "ALBANIA"
	Case "DZ" TranslateCountry = "ARGELIA"
	Case "AD" TranslateCountry = "ANDORRA"
	Case "AO" TranslateCountry = "ANGOLA"
	Case "AQ" TranslateCountry = "ANTÁRTIDA"
	Case "AG" TranslateCountry = "ANTIGUA Y BARBUDA"
	Case "AR" TranslateCountry = "ARGENTINA"
	Case "AM" TranslateCountry = "ARMENIA"
	Case "AW" TranslateCountry = "ARUBA"
	Case "AU" TranslateCountry = "AUSTRALIA"
	Case "AT" TranslateCountry = "AUSTRIA"
	Case "AZ" TranslateCountry = "AZERBAIJÁN"
	Case "BS" TranslateCountry = "BAHAMAS"
	Case "BH" TranslateCountry = "BAHRAYN"
	Case "BD" TranslateCountry = "BANGLADESH"
	Case "BB" TranslateCountry = "BARBADOS ISLAS DE BARLOVENTO"
	Case "BY" TranslateCountry = "BELARUS"
	Case "BE" TranslateCountry = "BELGICA"
	Case "BZ" TranslateCountry = "BELICE"
	Case "BJ" TranslateCountry = "BENÍN"
	Case "BM" TranslateCountry = "BERMUDAS"
	Case "BT" TranslateCountry = "BUTÁN"
	Case "BO" TranslateCountry = "BOLIVIA"
	Case "BA" TranslateCountry = "BOSNIA Y HERZEGOVINA"
	Case "BW" TranslateCountry = "BOTSWANA ESTADO DE  AFRICA AUSTRAL"
	Case "BV" TranslateCountry = "BOUVET ISLA NORUETA DEL ATLANTICO SUR"
	Case "BR" TranslateCountry = "BRASIL"
	Case "IO" TranslateCountry = "INDIAS BRITANICAS TERRITORIO DEL OCENO INDICO"
	Case "BN" TranslateCountry = "BRUNEI DARUSSALAM"
	Case "BG" TranslateCountry = "BULGARIA"
	Case "BF" TranslateCountry = "BURKINA FASO"
	Case "BI" TranslateCountry = "BURUNDI"
	Case "KH" TranslateCountry = "CAMBOYA"
	Case "CM" TranslateCountry = "CAMERÚN"
	Case "CA" TranslateCountry = "CANADÁ"
	Case "CV" TranslateCountry = "CAPE VERDE"
	Case "KY" TranslateCountry = "ISLAS CAYMAN"
	Case "CF" TranslateCountry = "REPÚBLICA DE AFRICA CENTRAL"
	Case "TD" TranslateCountry = "CHAD"
	Case "CL" TranslateCountry = "CHILE"
	Case "CN" TranslateCountry = "CHINA"
	Case "CX" TranslateCountry = "ISLA DE NAVIDAD"
	Case "CC" TranslateCountry = "ISLAS DE LOS COCOS"
	Case "CO" TranslateCountry = "COLOMBIA"
	Case "KM" TranslateCountry = "COMORES"
	Case "CG" TranslateCountry = "CONGO"
	Case "CD" TranslateCountry = "REPÚBLICA DEMOCRÁTICA DEL CONGO"
	Case "CK" TranslateCountry = "ISLAS DE COOK"
	Case "CR" TranslateCountry = "COSTA RICA"
	Case "CI" TranslateCountry = "CÔTE D'IVOIRE"
	Case "HR" TranslateCountry = "CROACIA"
	Case "CU" TranslateCountry = "CUBA"
	Case "CY" TranslateCountry = "CHIPRE"
	Case "CZ" TranslateCountry = "REPÚBLICA CHECA"
	Case "DK" TranslateCountry = "DINAMARCA"
	Case "DJ" TranslateCountry = "DJIBOUTI"
	Case "DM" TranslateCountry = "DOMINICA"
	Case "DO" TranslateCountry = "REPÚBLICA DOMINICANA"
	Case "EC" TranslateCountry = "ECUADOR"
	Case "EG" TranslateCountry = "EGYPTO"
	Case "SV" TranslateCountry = "EL SALVADOR"
	Case "VA" TranslateCountry = "EL VATICANO"
	Case "GQ" TranslateCountry = "GUINEA ECUATORIAL"
	Case "ER" TranslateCountry = "ERITREA"
	Case "EE" TranslateCountry = "ESTONIA"
	Case "ET" TranslateCountry = "ETIOPÍA"
	Case "FK" TranslateCountry = "ISLAS DE FALKAND MALVINAS"
	Case "FO" TranslateCountry = "ISLAS DE FAROE"
	Case "FJ" TranslateCountry = "FIDJI"
	Case "FI" TranslateCountry = "FINLANDIA"
	Case "FR" TranslateCountry = "FRANCIA"
	Case "GF" TranslateCountry = "GUINEA FRANCESA"
	Case "PF" TranslateCountry = "POLINESIA FRANCESA"
	Case "GA" TranslateCountry = "GABÓN"
	Case "GM" TranslateCountry = "GAMBIA"
	Case "GE" TranslateCountry = "GEORGIA"
	Case "DE" TranslateCountry = "GERMANIA"
	Case "GH" TranslateCountry = "GHÁNA"
	Case "GI" TranslateCountry = "GIBRALTAR"
	Case "GR" TranslateCountry = "GRECIA"
	Case "GL" TranslateCountry = "GROENLANDIA"
	Case "GD" TranslateCountry = "GRANADA"
	Case "GP" TranslateCountry = "GUADALUPE"
	Case "GU" TranslateCountry = "GUAM"
	Case "GT" TranslateCountry = "GUATEMALA"
	Case "GN" TranslateCountry = "GUINEA"
	Case "GW" TranslateCountry = "GUINEA PORTUGUESA"
	Case "GY" TranslateCountry = "GUYANA"
	Case "HT" TranslateCountry = "HAITÍ"
	Case "HN" TranslateCountry = "HONDURAS"
	Case "HK" TranslateCountry = "HONG KONG"
	Case "HU" TranslateCountry = "HUNGRIA"
	Case "IS" TranslateCountry = "ISLANDIA"
	Case "IN" TranslateCountry = "INDIA"
	Case "ID" TranslateCountry = "INDONESIA"
	Case "IR" TranslateCountry = "IRAN"
	Case "IQ" TranslateCountry = "IRAQ"
	Case "IE" TranslateCountry = "IRLANDIA"
	Case "IL" TranslateCountry = "ISRAEL"
	Case "IT" TranslateCountry = "ITALIA"
	Case "JM" TranslateCountry = "JAMAICA"
	Case "JP" TranslateCountry = "JAPÓN"
	Case "JO" TranslateCountry = "JORDANIA"
	Case "KZ" TranslateCountry = "KASAJISTÁN"
	Case "KE" TranslateCountry = "KENYA"
	Case "KI" TranslateCountry = "KIRIBATI"
	Case "KP" TranslateCountry = "REPÚBLICAS DEMOCRÁTICAS DE COREA"
	Case "KR" TranslateCountry = "REPÚBLICA DE COREA"
	Case "KW" TranslateCountry = "KUWAIT"
	Case "KG" TranslateCountry = "KIRGUIZISTÁN"
	Case "LA" TranslateCountry = "LAOS"
	Case "LV" TranslateCountry = "ESTADO RUSO DE LATVIA"
	Case "LB" TranslateCountry = "LÍBANO"
	Case "LS" TranslateCountry = "LESOTHO O BASUTOLANDIA"
	Case "LR" TranslateCountry = "LIBERIA"
	Case "LY" TranslateCountry = "LIBIA ARABE JAMAHIRYA"
	Case "LI" TranslateCountry = "LIECHTENSTEIN"
	Case "LT" TranslateCountry = "LITUNIA"
	Case "LU" TranslateCountry = "LUXEMBURGO"
	Case "MO" TranslateCountry = "MACAO"
	Case "MK" TranslateCountry = "MACEDONIA, ANTIGUA REPÚBLICA DE YUGOESLAVIA"
	Case "MG" TranslateCountry = "MADAGASCAR"
	Case "MW" TranslateCountry = "MALAWI"
	Case "MY" TranslateCountry = "MALASYA"
	Case "MV" TranslateCountry = "MALDIVAS"
	Case "ML" TranslateCountry = "MALÍ"
	Case "MT" TranslateCountry = "MALTA"
	Case "MH" TranslateCountry = "ISLAS MARSHALL"
	Case "MQ" TranslateCountry = "ISLA DE MARTINICA"
	Case "MR" TranslateCountry = "MAURITANIA"
	Case "MU" TranslateCountry = "ISLA MAURICIO"
	Case "YT" TranslateCountry = "MAYOTTÉ"
	Case "MX" TranslateCountry = "MÉXICO"
	Case "MD" TranslateCountry = "REPÚBLICA DE MOLDOVIA"
	Case "MC" TranslateCountry = "MÓNACO"
	Case "MN" TranslateCountry = "MONGOLIA"
	Case "MS" TranslateCountry = "MONTSERRAT"
	Case "MA" TranslateCountry = "MARRUECOS"
	Case "MZ" TranslateCountry = "MOZAMBIQUE"
	Case "MM" TranslateCountry = "BIRMANIA"
	Case "NA" TranslateCountry = "NAMIBIA"
	Case "NR" TranslateCountry = "NAURU"
	Case "NP" TranslateCountry = "NEPAL"
	Case "NL" TranslateCountry = "PAISES BAJOS"
	Case "AN" TranslateCountry = "ANTILLAS DE LOS PAISES BAJOS"
	Case "NC" TranslateCountry = "NUEVA CALEDONIA"
	Case "NZ" TranslateCountry = "NUEVA ZELANDA"
	Case "NI" TranslateCountry = "NICARAGUA"
	Case "NE" TranslateCountry = "NÍGER"
	Case "NG" TranslateCountry = "NIGERIA"
	Case "NU" TranslateCountry = "SAVAGE ISLA DEL PACÍFICO"
	Case "NF" TranslateCountry = "ISLA NORFOLK"
	Case "MP" TranslateCountry = "ISLAS MARIANAS DEL NORTE"
	Case "NO" TranslateCountry = "NORUEGA"
	Case "OM" TranslateCountry = "OMÁN"
	Case "PK" TranslateCountry = "PAKISTÁN"
	Case "PW" TranslateCountry = "PALAOS"
	Case "PS" TranslateCountry = "PALESTINA"
	Case "PA" TranslateCountry = "PANAMA"
	Case "PG" TranslateCountry = "NUEVA GUINEA - PAPUASIA"
	Case "PY" TranslateCountry = "PARAGUAY"
	Case "PE" TranslateCountry = "PERU"
	Case "PH" TranslateCountry = "FILIPINAS"
	Case "PN" TranslateCountry = "ISLA PITCAIRN"
	Case "PL" TranslateCountry = "POLONIA"
	Case "PT" TranslateCountry = "PORTUGAL"
	Case "PR" TranslateCountry = "PUERTO RICO"
	Case "QA" TranslateCountry = "QATAR"
	Case "RE" TranslateCountry = "RÉUNION"
	Case "RO" TranslateCountry = "RUMANIA"
	Case "RU" TranslateCountry = "FEDERACIÓN RUSIA"
	Case "RW" TranslateCountry = "RUANDA"
	Case "SH" TranslateCountry = "SANTA HELENA"
	Case "KN" TranslateCountry = "SANTA KITTS Y NEVIS "
	Case "LC" TranslateCountry = "SANTA LUCIA"
	Case "PM" TranslateCountry = "SANTO PIER Y MIKELON"
	Case "VC" TranslateCountry = "SAN VICENTE Y LAS GRANADINAS"
	Case "WS" TranslateCountry = "SAMOA"
	Case "SM" TranslateCountry = "SAN MARINO"
	Case "ST" TranslateCountry = "TOMO DE SAO Y PRÍNCIPE"
	Case "SA" TranslateCountry = "ARABIA SAUDITA"
	Case "SN" TranslateCountry = "SENEGAL"
	Case "CS" TranslateCountry = "SERBIA Y MONTENEGRO"
	Case "SC" TranslateCountry = "SEYCHELLES"
	Case "SL" TranslateCountry = "SIERRA LEONA"
	Case "SG" TranslateCountry = "SINGAPUR"
	Case "SK" TranslateCountry = "ESLOVAQUIA"
	Case "SI" TranslateCountry = "ESLOVENIA"
	Case "SB" TranslateCountry = "ISLAS SALOMON"
	Case "SO" TranslateCountry = "SOMALIA"
	Case "ZA" TranslateCountry = "SUR AFRICA"
	Case "GS" TranslateCountry = "GEORGIA DEL SUR Y LAS ISLAS DEL SUR DE SANDWICH"
	Case "ES" TranslateCountry = "ESPAÑA"
	Case "LK" TranslateCountry = "SRI LANKA"
	Case "SD" TranslateCountry = "SUDAN"
	Case "SR" TranslateCountry = "SURINAM"
	Case "SJ" TranslateCountry = "SVALBARD Y ENERO MAYEN"
	Case "SZ" TranslateCountry = "SWAZILANDIA"
	Case "SE" TranslateCountry = "SUECIA"
	Case "CH" TranslateCountry = "SUIZA"
	Case "SY" TranslateCountry = "REPÚBLICA ARABE DE SIRIA"
	Case "TW" TranslateCountry = "TAIWAN PROVINCIA DE CHINA"
	Case "TJ" TranslateCountry = "TAJIKISTAN"
	Case "TZ" TranslateCountry = "REPÚBLICA UNIDA DE TANZANIA"
	Case "TH" TranslateCountry = "TAILANDIA"
	Case "TL" TranslateCountry = "TIMOR-LESTE"
	Case "TG" TranslateCountry = "TOGO"
	Case "TK" TranslateCountry = "TOKELAU"
	Case "TO" TranslateCountry = "TONGA"
	Case "TT" TranslateCountry = "TRINIDAD Y TOBAGO"
	Case "TN" TranslateCountry = "TUNISIA"
	Case "TR" TranslateCountry = "TURKESTÁN"
	Case "TM" TranslateCountry = "TURKMENISTÁN"
	Case "TC" TranslateCountry = "ISLAS TURCAS Y TURCOS"
	Case "TV" TranslateCountry = "TUVALU"
	Case "UG" TranslateCountry = "UGANDA"
	Case "UA" TranslateCountry = "UCRANIA"
	Case "AE" TranslateCountry = "EMIRATOS ÁRABES UNIDOS"
	Case "GB" TranslateCountry = "REINGO UNIDO O INGLATERRA"
	Case "US" TranslateCountry = "ESTADOS UNIDOS"
	Case "UM" TranslateCountry = "ISLAS MENORES Y PERÍFERICAS DE ESTADOS UNIDOS"
	Case "UY" TranslateCountry = "URUGUAY"
	Case "UZ" TranslateCountry = "UZBEKISTÁN"
	Case "VU" TranslateCountry = "VANUATU"
	Case "VE" TranslateCountry = "VENEZUELA"
	Case "VN" TranslateCountry = "VIETNAM"
	Case "VG" TranslateCountry = "ISLAS BRITÁNICAS VIRGINIA"
	Case "VI" TranslateCountry = "ISLAS DE ESTADOS UNIDOS VIRGINIA"
	Case "WF" TranslateCountry = "WALLIS Y FUTUNA"
	Case "EH" TranslateCountry = "SAHARA OCCIDENTAL"
	Case "YE" TranslateCountry = "YEMEN"
	Case "ZM" TranslateCountry = "ZAMBIA"
	Case "ZW" TranslateCountry = "ZIMBABWE"
	Case "XX" TranslateCountry = "PAIS"
	Case "14.PAIS PROCEDENCIA" TranslateCountry = "14.PAIS PROCEDENCIA"
	Case "15.PAIS DESTINO" TranslateCountry = "15.PAIS DESTINO"
	Case Else TranslateCountry = ""
End Select
End Function

Function ShowServicesMenu(IndexService)
'IndexService toma el valor de variable MenuPos de la pagina que llama esta funcion para indicar el link seleccionado del menu
Dim Services, i, HTML, Pos
	Services = Session("Services")
	Pos = -1
	'select a.ServiceID, a.Name, a.URL, a.Typ
	if Session("CountServices")>=0 and IndexService=0 then
		IndexService = Services(0,0)
	end if
	
	for i=0 to Session("CountServices")
		if  Services(0,i)<>IndexService then
			if i=0 then
				HTML = "<td><img border='0' src='img/InitGray.gif'/></td>"
			else
				if Pos=1 then
					HTML = HTML & "<td><img border='0' src='img/WhiteGray.gif'/></td>"
				else
					HTML = HTML & "<td><img border='0' src='img/GrayGray.gif'/></td>"
				end if
			end if
			HTML = HTML & SetLink("G", Services(0,i), Services(1,i), Services(2,i), Services(3,i), Services(4,i), Services(5,i))
			Pos=0
		else
			if i=0 then
				HTML = "<td><img border='0' src='img/InitWhite.gif'/></td>"
			else
				HTML = HTML & "<td><img border='0' src='img/GrayWhite.gif'/></td>"
			end if
			HTML = HTML & SetLink("W", Services(0,i), Services(1,i), Services(2,i), Services(3,i), Services(4,i), Services(5,i))
			Pos=1
		end if
	next
	'Colocando el Final del Menu
	select case Pos
	case 0
		HTML = HTML & "<td><img border='0' src='img/FinGray.gif'/></td><td width='100%' valign='top'><div class='L'>&nbsp;</div></td>"
	case 1
		HTML = HTML & "<td><img border='0' src='img/FinWhite.gif'/></td><td width='100%' valign='top'><div class='L'>&nbsp;</div></td>"
	end select
	'HTML = HTML & "<td class='contenido1' rowspan='2' colspan='2' valign='middle'><table cellpadding='0' cellspacing='0' border='0' align='right'><tr><td><a href='javascript:help();'><img src='img/help.png' border='0' /></a></td><td><a href='javascript:help();' class='contenido1' onmouseover='this.className=link4;' onmouseout='this.className=contenido1;'>Help</a></td></tr></table></td>"
	ShowServicesMenu = HTML
End Function

Function SetLink (klass, ID, Title, Link, Typ, H, W)
	SetLink = ""
	Select Case Typ
	Case 0
		SetLink = "<td valign='top'><div class='" & klass & "'><a class='" & klass & "' href='"& HTMLBaseSSL & "tracking-services/" & Link & "?S=" & ID & "'><" & "script>TranslateServices(""" & Title & """);</script" & "></a></div></td>"		
	Case 1
		SetLink = "<td valign='top'><div class='" & klass & "'><a class='" & klass & "' href='#' onclick='Javascript:window.open("""& HTMLBaseSSL & "tracking-services/" & Link & """, ""P" & ID & """, ""menubar=0,resizable=1,scrollbars=1,width=" & W & ",height=" & H & """);return false;'>" & Replace(Title," ", "&nbsp;") & "</a></div></td>"
	End Select
End Function

Function UpdatescriptLang(Conn, Langs, CountLangs)
Dim rs, Translations, CantTranslations, i, j, ntr, Jscript, fs, fname
	ntr = chr(13) & chr(10)

	CantTranslations = -1
	Set rs = Conn.Execute("select * from Links")
	if Not rs.EOF then
		Translations = rs.GetRows
		CantTranslations = rs.RecordCount-1
	end if
	CloseOBJ rs
	
	Jscript = "var aBases = new Array();"
	for j=0 to CantTranslations
		Jscript = Jscript & ntr & "aBases[" & j & "]='" & Translations(4,j) & "';"
	next

	for i=0 to CountLangs
		Jscript = Jscript & ntr & "var a" & Langs(2,i) & " = new Array();"
		for j=0 to CantTranslations
			Jscript = Jscript & ntr & "a" & Langs(2,i) & "[" & j & "]='" & Translations(5+i,j) & "';"
		next
	next
	
	set fs=Server.CreateObject("scripting.FileSystemObject")
	set fname=fs.CreateTextFile("C:\Inetpub\wwwroot\site\Javascripts\languages.js",true)
	fname.Write(Jscript)
	fname.Close
	
	Set fname = nothing
	Set fs = nothing
	Set Translations = nothing
End Function

Function SetChatSupport
	Select Case Request.Cookies("Country")
	Case "BZ"
		response.write "http://aimarcommunicator.imserver.net/WebPanel/MsgPanel.asp?WCI=Panel&amp;NombreUsuario=managerbzl&amp;NumeroSerie=20000l&amp;LangId=1"
	Case "SV"
		response.write "http://aimarcommunicator.imserver.net/WebPanel/MsgPanel.asp?WCI=Panel&amp;NombreUsuario=trafficsv01&amp;NumeroSerie=20000&amp;LangId=1"
	Case "HN"
		response.write "http://aimarcommunicator.imserver.net/WebPanel/MsgPanel.asp?WCI=Panel&amp;NombreUsuario=traffichn04&amp;NumeroSerie=20000&amp;LangId=1"
	Case "NI"
		response.write "http://aimarcommunicator.imserver.net/WebPanel/MsgPanel.asp?WCI=Panel&amp;NombreUsuario=trafficnic01&amp;NumeroSerie=20000&amp;LangId=1"
	Case "CR"
		response.write "http://aimarcommunicator.imserver.net/WebPanel/MsgPanel.asp?WCI=Panel&amp;NombreUsuario=customercr01&amp;NumeroSerie=20000&amp;LangId=1"
	Case "PA"
		response.write "http://aimarcommunicator.imserver.net/WebPanel/MsgPanel.asp?WCI=Panel&amp;NombreUsuario=accountpa01&amp;NumeroSerie=20000&amp;LangId=1"
	Case Else
		response.write "http://aimarcommunicator.imserver.net/WebPanel/MsgPanel.asp?WCI=Panel&amp;NombreUsuario=trafficgt14&amp;NumeroSerie=20000&amp;LangId=1"
	End Select
End Function

Function SetPassword(Value)
Dim Pwd
    Pwd = replace(UCASE(Value)," ","",1,-1)
    Pwd = replace(Pwd,"0","",1,-1)
    Pwd = replace(Pwd,"1","i",1,-1)
    Pwd = replace(Pwd,"2","z",1,-1)
    Pwd = replace(Pwd,"3","m",1,-1)
    Pwd = replace(Pwd,"4","a",1,-1)
    Pwd = replace(Pwd,"5","s",1,-1)
    Pwd = replace(Pwd,"6","g",1,-1)
    Pwd = replace(Pwd,"7","t",1,-1)
    Pwd = replace(Pwd,"8","b",1,-1)
    Pwd = replace(Pwd,"9","p",1,-1)
    Pwd = replace(Pwd,"-","",1,-1)
    Pwd = replace(Pwd,",","",1,-1)
    Pwd = replace(Pwd,".","",1,-1)
    Pwd = replace(Pwd,"/","",1,-1)
    Pwd = replace(Pwd,":","",1,-1)
    Pwd = replace(Pwd,";","",1,-1)
    Pwd = replace(Pwd,"'","",1,-1)
    Pwd = replace(Pwd,"A","4",1,-1)
    Pwd = replace(Pwd,"B","8",1,-1)
    Pwd = replace(Pwd,"C","u",1,-1)
    Pwd = replace(Pwd,"D","p",1,-1)
    Pwd = replace(Pwd,"E","3",1,-1)
    Pwd = replace(Pwd,"F","t",1,-1)
    Pwd = replace(Pwd,"G","6",1,-1)
    Pwd = replace(Pwd,"I","7",1,-1)
    Pwd = replace(Pwd,"L","y",1,-1)
    Pwd = replace(Pwd,"M","w",1,-1)
    Pwd = replace(Pwd,"P","d",1,-1)
    Pwd = replace(Pwd,"R","9",1,-1)
    Pwd = replace(Pwd,"S","5",1,-1)
    Pwd = replace(Pwd,"S","z",1,-1)
    Pwd = replace(Pwd,"T","h",1,-1)
    Pwd = replace(Pwd,"U","n",1,-1)
    Pwd = replace(Pwd,"V","a",1,-1)
    Pwd = replace(Pwd,"W","m",1,-1)
    Pwd = replace(Pwd,"Z","f",1,-1)
    
    SetPassword = Left(LCASE(Pwd),6)
End Function
</script>
