<%@ Language=VBScript %>
<%Option Explicit


'Response.Write(Request("file"))
'Response.Write(Request("name"))
'Response.End

'Response.ContentType = "application/vnd.ms-excel"
'Response.AppendHeader "content-disposition", " filename=excelTest.xls"



dim fs,f,t,x
set fs=Server.CreateObject("Scripting.FileSystemObject")
'set f=fs.CreateTextFile("c:\test.txt")
'f.write("Hello World!")
'f.close

set t=fs.OpenTextFile(Request("file"),1,false)
x=t.ReadAll
t.close
Response.Write(x)

%>
