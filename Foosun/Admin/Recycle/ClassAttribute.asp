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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����վ-��Ŀ����</title>
</head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet">
<body topmargin="0" leftmargin="4">
<div align="center"> 
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <%
	dim ClassID,ClassSql,RsClassObj
	ClassID = Request("ID")
	if ClassID <> "" then
	ClassSql = "Select * from FS_NewsClass where ClassID='" & ClassID & "'"
	Set RsClassObj = Conn.Execute(ClassSql)
	if Not RsClassObj.Eof then
%>
	<tr>
	<td height="20" colspan="2"></td>
	</tr>
    <tr> 
      <td width="28%" rowspan="3"><div align="center"><img src="../../Images/Info.gif" width="34" height="33"></div></td>
      <td width="72%" height="22">��Ŀ���ƣ�
      <% = RsClassObj("ClassCName") %></td>
    </tr>
    <tr> 
      <td height="22">����Ŀ��
      <% 
			If RsClassObj("ParentID")=0 then 
				Response.Write("��Ŀ¼")
			else
				Dim RsTemp,TempSql
				TempSql="Select ClassCName From FS_NewsClass Where ClassID='"&RsClassObj("ParentID")&"'"
				Set RsTemp = Conn.Execute(TempSql)
				 response.write(RsTemp("ClassCName"))
				 Set RsTemp = Nothing
			end if
%></td>
    </tr>
	<tr>
      <td height="22">ģ�����ƣ�
      <% =RsClassObj("ClassTemp") %></td>
	</tr>
	<tr>
	<td height="13" colspan="2"></td>
	</tr>
    <tr> 
      <td height="22" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;���ʱ�䣺 
      <% = RsClassObj("AddTime") %></td>
    </tr>
	<tr> 
      <td height="22" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;ɾ��ʱ�䣺 
      <% = RsClassObj("DelTime") %></td>
    </tr>
	<tr>
	<td height="10" colspan="2"></td>
	</tr>
	 <tr>
	  <td height="22" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;�� �� ����
<%  
			Set RsTemp = Conn.Execute("Select Count(*) from FS_News where ClassID ='"& RsClassObj("ClassID")&"'")
			Response.Write(RsTemp(0))
			Set RsTemp = Nothing 
%>
		</td>
    </tr>
    <tr> 
      <td height="22" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;����Ŀ���� <% = RsClassObj("ChildNum") %></td>
	 </tr>
	 <tr>
      <td height="22" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;�Ƿ�����Ͷ�壺  
<%  
			dim Flag
			Flag="��"
			if RsClassObj("Contribution") = 1 then Flag = "��"
			Response.Write(Flag)
%>
		</td>
    </tr>
<%
	else
%>
    <tr> 
      <td height="21" colspan="3"><div align="center">Nothing</div></td>
    </tr>
    <%
	end if
else
%>
    <tr> 
      <td height="50" colspan="3"><div align="center"></div></td>
    </tr>
    <%
end if
%>
	<tr>
	<td height="40" colspan="2"></td>
	</tr>
    <tr> 
      <td height="30" colspan="3"><div align="center"> 
          <input name="Submitasd" onClick="window.close();" type="button" id="Submitasd" value=" �� �� ">
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
Set RsClassObj = Nothing
Set Conn = Nothing
%>