<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P030800") then Call ReturnError()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ѡ�����ű�ǩ����</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<body bgcolor="#E6E6E6" topmargin="20">
<div align="center"> 
  <table width="90%" border="0" cellpadding="0" cellspacing="5">
    <tr> 
      <td height="20"> <div align="center">�����ֶ� 
          <select style="width:60%;" name="InsertFun">
            <option selected>ѡ������ֶ�</option>
            <option value="{News_Title}">����</option>
            <option value="{News_SubTitle}">������</option>
            <option value="{News_Author}">����</option>
            <option value="{News_Content}">����</option>
            <option value="{News_TxtSource}">��Դ</option>
            <option value="{News_TxtEditer}">���α༭</option>
            <option value="{News_AddDate}">����</option>
            <option value="{News_SendFriend}">���͸�����</option>
            <option value="{News_ReviewContent}">����</option>
            <option value="{News_Review}">��������</option> 
            <option value="{News_ClickNum}">���ŵ������</option> 
            <option value="{News_Favorite}">��ӵ��ղؼ�</option> 
          </select>
        </div></td>
    </tr>
    <tr>
      <td height="5"></td>
    </tr>
    <tr> 
      <td height="20"> <div align="center"> 
          <input name="Submitdd" onClick="InsertScript(document.all.InsertFun);" type="button" id="Submitdd" value=" �� �� ">
          <input type="button" onClick="window.close();" name="Submit2" value=" ȡ �� ">
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
Set Conn = Nothing
%>
<script language="JavaScript">
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
function InsertScript(Obj)
{
	var re=/[\$]/ig;
	var TempStr=Obj.value;
	TempStr=TempStr.replace(re,'"');
	window.returnValue=TempStr;
	window.close();
}
</script>