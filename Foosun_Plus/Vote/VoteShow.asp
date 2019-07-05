<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Function.asp" -->
<!--#include file="../../Inc/NoSqlHack.asp" -->
<%
   Dim DBC,Conn
   Set DBC = New databaseclass
   Set Conn = DBC.openconnection()
   Set DBC = Nothing
   
Dim addr,VoteObj,vote,title,style,VoteID,VoteConfigObj
Set VoteConfigObj = Conn.Execute("Select DoMain from FS_Config")
addr=VoteConfigObj("DoMain")&"/"& PlusDir &"/Vote/"
if isnull(request("VoteID")) or request("VoteID")="" then 
	set VoteObj=conn.execute("select top 1 * from FS_Vote order by VoteID desc")
	vote="vote"
else
	set VoteObj=conn.execute("select * from FS_Vote where VoteID='"&request("VoteID")&"'")
	vote="vote"&request("VoteID")
end if

response.write vote&"="&""""&""&""""&chr(10)

if VoteObj.eof then
response.write vote&"="&vote&"+"&""""&"<font color=red>暂无投票项目</font>"&""""&chr(10)
response.write "document.write ("&vote&");"
response.end
end if

title=VoteObj("Name")
if VoteObj("Type") = "1" then
	style="checkbox"
else
	style="radio"
end if
VoteID = VoteObj("VoteID")
VoteObj.close

Set VoteObj = Conn.Execute("select * from FS_VoteOption where VoteID='"&VoteID&"'")

response.write vote&"="&vote&"+"&""""&"<form name=form"&vote&" method='POST'>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<table border='0' width='100%' cellspacing='0' cellpadding='0'>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<tr>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<td width='100%' height='25'>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<p align='center'><b>"&title&"<b></td>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"</tr>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<tr>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<td width='100%'>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<div align='center'>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<center>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<table border='0' width='90%' cellspacing='0' cellpadding='0'>"&""""&chr(10)

do while not VoteObj.eof
response.write vote&"="&vote&"+"&""""&"<tr>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<td width='100%'><input type='"&style&"' name='voted' value='"&VoteObj("ID")&"' size='20'>"&VoteObj("OptionName")&"</td>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"</tr>"&""""&chr(10)
VoteObj.movenext
loop


response.write vote&"="&vote&"+"&""""&"</table>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"</center>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"</div>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"</td>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"</tr>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<tr>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<td width='100%' valign='bottom' align='center' height='30'><input type='button' value='投票' name='B1' onclick='tou"&vote&"();'>&nbsp;&nbsp;<input type='button' value='查看' name='B1' onclick=window.open('"&addr&"VoteResult.asp?VoteID="&VoteID&"','','width=650,height=220,resizable=1,scrollbars=1');></td>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"</tr>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"</table></form>"&""""&chr(10)
response.write vote&"="&vote&"+"&""""&"<\script>function tou"&vote&"(){window.open('about:blank','tou','width=650,height=220');form"&vote&".action='"&addr&"VoteCheck.asp?VoteID="&VoteID&"&style="&style&"';form"&vote&".target='tou';form"&vote&".submit();}</\script>"&""""&chr(10)

VoteObj.close
set VoteObj=nothing
Set VoteConfigObj = Nothing
Conn.close
set Conn=nothing
response.write "document.write ("&vote&")"&chr(10)
%>