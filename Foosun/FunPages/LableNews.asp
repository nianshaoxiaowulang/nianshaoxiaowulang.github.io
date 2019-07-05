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
<title>选择新闻标签属性</title>
</head>
<link rel="stylesheet" href="../../CSS/ModeWindow.css">
<body bgcolor="#E6E6E6" topmargin="20">
<div align="center"> 
  <table width="90%" border="0" cellpadding="0" cellspacing="5">
    <tr> 
      <td height="20"> <div align="center">插入字段 
          <select style="width:60%;" name="InsertFun">
            <option selected>选择插入字段</option>
            <option value="{News_Title}">标题</option>
            <option value="{News_SubTitle}">副标题</option>
            <option value="{News_Author}">作者</option>
            <option value="{News_Content}">内容</option>
            <option value="{News_TxtSource}">来源</option>
            <option value="{News_TxtEditer}">责任编辑</option>
            <option value="{News_AddDate}">日期</option>
            <option value="{News_SendFriend}">发送给好友</option>
            <option value="{News_ReviewContent}">评论</option>
            <option value="{News_Review}">发表评论</option> 
            <option value="{News_ClickNum}">新闻点击次数</option> 
            <option value="{News_Favorite}">添加到收藏夹</option> 
          </select>
        </div></td>
    </tr>
    <tr>
      <td height="5"></td>
    </tr>
    <tr> 
      <td height="20"> <div align="center"> 
          <input name="Submitdd" onClick="InsertScript(document.all.InsertFun);" type="button" id="Submitdd" value=" 插 入 ">
          <input type="button" onClick="window.close();" name="Submit2" value=" 取 消 ">
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