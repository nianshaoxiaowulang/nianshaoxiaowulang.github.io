<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Const.asp" -->
<%
Dim  DBC,Conn
Set  DBC = New DataBaseClass
Set  Conn = DBC.OpenConnection()
Set  DBC = Nothing
DownLoadID=Replace(Replace(request("DownLoadID"),"'",""),Chr(39),"")
Set rs = server.createobject(G_FS_RS)
Rs.source = "select ClickNum from FS_DownLoad where DownLoadID='"&DownLoadID&"'"
Rs.open Rs.source,conn,1,1
if Not Rs.Eof then
%>
   javastr="<%=rs("ClickNum")%>"
   document.write(javastr)
<%
else
%>
   javastr="0"
   document.write(javastr)
<%
End If
Rs.close
Set Rs=nothing
Set Conn = Nothing
%>
