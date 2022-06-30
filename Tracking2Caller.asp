<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
-->
<%
    if request("AWBID2") <> "" then        
%>    
        <!-- #INCLUDE file="Utils.asp" -->
        <!-- #INCLUDE file=Tracking2.asp -->
<%
        'response.write "Leido correctamente.<br>"
        
        
        
        dim aList1Values23        
        'LECTURA DE LOS DATOS BASICOS DE LAS GUIAS  Tracking2.asp
        aList1Values23 = GetaList1Values(Request("AWBID2"),Request("AwbType"))
        
        'response.write "(" & Request("AwbType") & ")"
        'response.write "(" & aList1Values23(0,0) & ")"
        'response.write "(" & aList1Values23(1,0) & ")"
        'response.write "(" & aList1Values23(2,0) & ")"
        'response.write "(" & aList1Values23(3,0) & ")"
        'response.write "(" & aList1Values23(4,0) & ")"
        'response.write "(" & aList1Values23(5,0) & ")"
        'response.write "(" & aList1Values23(6,0) & ")"
        'response.write "(" & aList1Values23(7,0) & ")"
        'response.write "(" & aList1Values23(8,0) & ")"
        'response.write "(" & aList1Values23(9,0) & ")"
        'response.write "<br><br>"

        'response.write CheckNum(Request("SendAgent")) & " " & CheckNum(Request("SendConsigner")) & " " & CheckNum(Request("SendShipper")) & " " & Request("AwbType") & "<br><br>"

        dim SQLQuery        
        'LECTURA DE LOS CONTACTOS INTERNOS / EXTERNOS Tracking2.asp
        SQLQuery = GetSQLQuery(CheckNum(Request("SendAgent")), CheckNum(Request("SendConsigner")), CheckNum(Request("SendShipper")), aList1Values23, Request("AwbType"))  

        'response.write "(" & SQLQuery & ")<br>"

        if SQLQuery <> "" then
        'PROCESO DE TODOS LOS DATOS YA CAPTURADOS / ELABORACION DE EMAILS / ENVIO   Tracking2.asp
            SQLQuery = SendNotification(Request("BLStatusName"), SQLQuery, Request("AwbType"), Request("Comment"), aList1Values23, CheckNum(Request("SendShipper")),Request("mode"))
        else 
            response.write "Fallo al traer el query"

            if Request("mode") = 1 then
                'response.write "\n"
            else
                response.write "<br>"
            end if

        end if

        'response.write SQLQuery 
    else
        response.write "No vienen parametros necesarios!"
    end if


    'AWBID2=&BLStatusName=FECHA DE INGRESO&Comment=Automatico&SendAgent=1&SendConsigner=1&SendShipper=1

%>

