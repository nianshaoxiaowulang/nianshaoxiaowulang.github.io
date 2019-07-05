<% Option Explicit %>
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<%
	Dim DBC,Conn
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->

<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070307") then Call ReturnError()
   
    Dim VoteID,CodeStr,RRsVoteConfigObj
	Set RRsVoteConfigObj = Conn.Execute("Select DoMain from FS_Config")
		VoteID = Cstr(Request("VoteID"))
		CodeStr = RRsVoteConfigObj("DoMain")&"/"& PlusDir &"/Vote/VoteShow.asp?VoteID="&VoteID&""
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<body topmargin="0" leftmargin="0">
<table width="75%" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr> 
    <td height="15">&nbsp;</td>
  </tr>
  <tr> 
    <td><font color="#0000FF" size="3">&nbsp;&nbsp;本投票项目调用代码为:</font></td>
  </tr>
  <tr> 
    <td> <div align="center"> 
        <textarea name="textfield" cols="83" rows="6"><script src=<%=CodeStr%>></script></textarea>
      </div></td>
  </tr>
  <tr> 
    <td> <div align="center"> 
        <input type="button" name="Submit3" value=" 关 闭 " onclick="window.close();">
      </div></td>
  </tr>
  <tr> 
    <td height="10">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<script>
  document.all.textfield.select();
</script>
