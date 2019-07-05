<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System(FoosunCMS V3.1.0930)
'最新更新：2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'商业注册联系：028-85098980-601,项目开发：028-85098980-606、609,客户支持：608
'产品咨询QQ：394226379,159410,125114015
'技术支持QQ：315485710,66252421 
'项目开发QQ：415637671，655071
'程序开发：四川风讯科技发展有限公司(Foosun Inc.)
'Email:service@Foosun.cn
'MSN：skoolls@hotmail.com
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.cn  演示站点：test.cooin.com 
'网站通系列(智能快速建站系列)：www.ewebs.cn
'==============================================================================
'免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接
'风讯公司保留此程序的法律追究权利
'如需进行2次开发，必须经过风讯公司书面允许。否则将追究法律责任
'==============================================================================
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P040503") then Call ReturnError1()
Dim Path,FileName,EditFile,FileContent,Result
Result = Request.Form("Result")
Path = "../../../Inc"
FileName = "Const.asp"
EditFile = Server.MapPath(Path) & "\" & FileName
Dim FsoObj,FileObj,FileStreamObj
Set FsoObj = Server.CreateObject(G_FS_FSO)
Set FileObj = FsoObj.GetFile(EditFile)
if Result = "" then
	Set FileStreamObj = FileObj.OpenAsTextStream(1)
	if Not FileStreamObj.AtEndOfStream then
		FileContent = FileStreamObj.ReadAll
	else
		FileContent = ""
	end if
else
	Set FileStreamObj = FileObj.OpenAsTextStream(2)
	FileContent = Request.Form("ConstContent")
	FileStreamObj.Write FileContent
	if Err.Number <> 0 then
		%>
		<script language="JavaScript">
			alert('<% = "保存失败，请拷贝后，重新打开文件再保存" %>');window.location='SysConstSet.asp';
		</script>
		<%
		
	else
		%>
		<script language="JavaScript">
			alert('修改成功');window.location='SysConstSet.asp';
		</script>
		<%
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>系统常量设置</title>
</head>
<style type="text/css">
<!--
.SysParaButtonStyle {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-right-color: #999999;
	border-bottom-color: #999999;
	border-left-color: #FFFFFF;
	background-color: #E6E6E6;
}
-->
</style>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<body bgcolor="#FFFFFF" topmargin="2" leftmargin="2" scroll=yes  oncontextmenu="return false;">
<form action="" name="Form" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width=35 align="center" alt="保存" onClick="document.Form.submit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp;<input name="Result" type="hidden" id="Result" value="Modify">
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	 <tr> 
      <td colspan="2"><div align="left">
        <font color="#FF0000">&nbsp;&nbsp 注意：<br>&nbsp;&nbsp 设置的时候除SysRootDir和数据库路径带虚拟路径（如果有），其余均不加虚拟路径。请大家注意设置，以免给您的调试造成不必要的麻烦</font>  
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center">
          <textarea name="ConstContent" rows="34" style="width:99%;"><% = FileContent %></textarea>
        </div></td>
    </tr>
    <tr> 
      <td height="30" colspan="2"><div align="center"><font color="#FF0000">注意：主窗口里中任意单引号&quot;<font color="#0000FF">'</font>&quot;不能去掉，请不要使用回车。&quot;<font color="#0000FF">&lt;%</font>&quot;和&quot;<font color="#0000FF">%&gt;</font>&quot;不能去掉。请一定注意。建议只改字符，不要增加、删除</font></div></td>
    </tr>
</table>
</form>
</body>
</html>
<%
Set FsoObj = Nothing
Set FileObj = Nothing
Set FileStreamObj = Nothing
%>
