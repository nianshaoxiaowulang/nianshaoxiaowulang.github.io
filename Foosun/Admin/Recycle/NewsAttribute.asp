<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->

<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070103") then Call ReturnError()
Dim RsNewsObj,NewsSql,NewsID
NewsID = Request("ID")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����վ-��������</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="14">
<div align="center">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <%
		if NewsID <> "" then
			NewsSql = "Select * from FS_News where NewsID='" & NewsID & "'"
			Set RsNewsObj = Conn.Execute(NewsSql)
			if Not RsNewsObj.Eof then
%>
	<tr>
	<td height="20" colspan="2"></td>
	</tr>
    <tr> 
    	<td width="28%" rowspan="3"><div align="center"><img src="../../Images/Info.gif" width="34" height="33"></div></td>
        <td width="72%" height="27">���ű��⣺
        <% = RsNewsObj("Title") %></td>
    </tr>
	<tr>
      	<td height="28">�� �� �⣺
        <% = RsNewsObj("SubTitle") %></td>
	</tr>
    <tr> 
    	<td height="27">�������ͣ�
          <% 
	  				Dim NewsType
	  				NewsType="��������"
					if  RsNewsObj("HeadNewsTF") = 1 then NewsType="��������"
					if  RsNewsObj("PicNewsTF") = 1 then NewsType="ͼƬ����"
					Response.Write(NewsType)
%></td>
    </tr>
		<tr><td height="27" colspan="2"></td>
	<tr>
		<td height="27" colspan="2" >&nbsp;&nbsp;&nbsp;&nbsp;������Ŀ��
<% 
	 				Dim RsTemp
					Set RsTemp = Conn.Execute("Select ClassCName From FS_NewsClass Where ClassID= '"&RsNewsObj("ClassID")&"'") 
					Response.Write(RsTemp("ClassCName"))
 %>	
 		</td>
 	</tr>
	<tr height="2">
    	<td height="25" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;������ڣ�  <% = RsNewsObj("AddDate") %></td>
    </tr>
    <tr> 
    	<td height="28" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;ɾ������:   <% = RsNewsObj("DelTime") %></td>
    </tr>
	<tr>
		<td height="30" colspan="2">&nbsp;</td>
    </tr>
<%
	else
%>
    <tr> 
      <td colspan="3"> <div align="center">���Ų�����</div></td>
    </tr>
    <%
	end if
else
%>
    <tr> 
      <td colspan="3"> <div align="center">�������ݴ���</div></td>
    </tr>
    <%
end if
%>
    <tr> 
      <td height="50" colspan="3"> <div align="center"> 
          <input name="Submitasd" onClick="window.close();" type="button" id="Submitasd" value=" �� �� ">
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
Set RsTemp = Nothing
Set RsNewsObj = Nothing
Set Conn = Nothing
%>
