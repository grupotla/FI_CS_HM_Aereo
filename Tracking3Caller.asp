<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
-->
<%
    if Request("AWBID2") <> "" then        
%>    
        <!-- #INCLUDE file="Utils.asp" -->
        <!-- #INCLUDE file=Tracking3.asp -->
<%
        'response.write "Leido correctamente.<br>"
        
        
        
        dim aList1Values23        

        dim AWBID2, SendAgent, SendConsigner, SendShipper, AwbType, mode

         AWBID2 = CheckNum(Request("AWBID2"))
         SendAgent = CheckNum(Request("SendAgent"))
         SendConsigner = CheckNum(Request("SendConsigner"))
         SendShipper = CheckNum(Request("SendShipper")) 
         AwbType = CheckNum(Request("AwbType")) 
         mode = CheckNum(Request("mode"))


        aList1Values23 = GetaList1Values(AWBID2, AwbType)
        
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
        SQLQuery = GetSQLQuery(SendAgent, SendConsigner, SendShipper, aList1Values23, AwbType)  

        'response.write "(" & SQLQuery & ")<br>"

        if SQLQuery <> "" then
            SQLQuery = SendNotification(Request("BLStatusName"), SQLQuery, AwbType, Request("Comment"), aList1Values23, SendShipper, mode)
        else 
            response.write "Fallo al traer el query"

            if mode = 1 then
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

