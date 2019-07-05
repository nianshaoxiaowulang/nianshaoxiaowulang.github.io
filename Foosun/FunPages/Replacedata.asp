<% Option Explicit %>
<!--#include file="../../Inc/NoSqlhack.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080700") then Call ReturnError1()
Dim NewsTableName,NewsFieldName,DldTableName,DldFieldName,strSQl
NewsTableName = Request.Form("NewsTableName")
NewsFieldName = Request.Form("NewsFieldName")
DldTableName = Request.Form("DldTableName")
DldFieldName = Request.Form("DldFieldName")

If NewsTableName="FS_News" And NewsFieldName<>"" And Request.Form("strFindNews")<>"" then
	strSQl="select "&NewsFieldName&" from "&NewsTableName
	ReplaceData strSQl,Request.Form("strFindNews"),Request.Form("strReplaceNews"),Conn
End If 
If DldTableName="FS_Download" And DldFieldName<>"" And Request.Form("strFindDld")<>"" then
	strSQl="select "&DldFieldName&" from "&DldTableName
	ReplaceData strSQl,Request.Form("strFindDld"),Request.Form("strReplaceDld"),Conn
End If 
Function ReplaceData(strSQL,strPattern,strReplace,tempConn)'strSQL提取需要替换的字段，strPattern要被替换的内容匹配字符，strReplace替换的字符,tempConn数据库连接
	If strPattern="" Then
		Response.write("<script>alert('您没有填写查找内容！');history.back();</script>")
	End If 
	Dim strFind
	Set strFind = server.createobject(G_FS_RS)
	strFind.open strSQL,tempConn,3,3
	On Error Resume Next
	Do While NOT strFind.EOF
		If Not IsNull(strFind(0)) Then 			
			strFind(0)=Replace(strFind(0),strPattern,strReplace)
			strFind.update
		End If
		strFind.moveNext
	Loop 
	Response.write("<script>alert('该字段内容已经全部替换成功！');history.back();</script>")
End Function 

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>字段内容替换</title>
</head>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<script src="../SysJS/PublicJS.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2" oncontextmenu="//return false;">
<form action="" method="post" name="replacedata" id="replacedata">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="提交修改数据" onClick="document.replacedata.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">提交</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr>
  <td valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="dddddd">
        <tr> 
          <td width="10%" height="26" class="ButtonListLeft"> 
            <div align="center">表名</div></td>
          <td width="25%" height="26" class="ButtonList"> 
            <div align="center">字段</div></td>
          <td width="35%" height="26" class="ButtonList"> 
            <div align="center">查找内容</div></td>
          <td width="35%" height="26" class="ButtonList"> 
            <div align="center">替换内容</div></td>
  </tr>

	    <tr valign="bottom" bgcolor="#FFFFFF" height="30"> 
          <td align="left"> 
            <div align="center">新闻表</div><input type="hidden" name="NewsTableName" value="FS_News">
          </td>
          <td align="left"> 
            <div align="center"><select name="NewsFieldName">
		<option value="Content">新闻内容</option>
		<option value="KeyWords">关键字</option>
		<option value="TxtSource">新闻来源</option>
		<option value="Author">新闻作者</option>
		<option value="Editer">新闻责任编辑</option>
		<option value="NewsTemplet">新闻模板文件&nbsp;</option>		
       </select> </div>
      </td>
	      <td align="center"> 
            <div align="center">	<input type="text" name="strFindNews" onmouseover="this.focus();"> </div> </td>
          <td  align="center"> 
            <div align="center"><input type="text" name="strReplaceNews" onmouseover="this.focus();"></div> </td>
	</tr>
        <tr valign="bottom" bgcolor="#FFFFFF" height="30"> 
          <td align="left"> 
            <div align="center">下载表</div>	  </td><input type="hidden" name="DldTableName" value="FS_Download">
          <td align="left"> 
            <div align="center"><select name="DldFieldName">
		<option value="Name" >下载名字</option>
		<option value="Description" >下载介绍</option>
		<option value="NewsTemplet" >模板文件名</option>
		<option value="Provider" >开发商</option>
		<option value="ProviderUrl">提供者url地址</option>
       </select> </div>
      </td>

	      <td align="center"> 
            <div align="center">	<input type="text" name="strFindDld" onmouseover="this.focus();"> 	</div>  </td>
          <td align="center"> 
            <div align="center"><input type="text" name="strReplaceDld" onmouseover="this.focus();"> </div>   </td>
</tr>
</table>
</form>
</body>
</html>
