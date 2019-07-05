<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="inc/Config.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
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
Dim DBC,Conn,CollectConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = CollectDBConnectionStr
Set CollectConn = DBC.OpenConnection()
Set DBC = Nothing
'判断权限
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P080105") then Call ReturnError1()
'判断权限结束
Dim SelectPath
if SysRootDir = "" then
	SelectPath = "/" & UpFiles
else
	SelectPath = "/" & SysRootDir & "/" & UpFiles
end if
Dim RsEditObj,EditSql,SiteID
Set RsEditObj = Server.CreateObject ("ADODB.RecordSet")
SiteID = Request("SiteID")
if SiteID <> "" then
	EditSql="Select * from FS_Site where ID=" & SiteID
	RsEditObj.Open EditSql,CollectConn,1,3
	if RsEditObj.Eof then
		Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
		Response.end
	end if
else
	Response.write"<script>alert(""没有修改的站点"");location.href=""javascript:history.back()"";</script>"
	Response.end
end if

Dim TempClassListStr
TempClassListStr = ClassList
Function ClassList()
	Dim ClassListObj,SelectStr
	Set ClassListObj = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='0' order by ClassID desc")
	do while Not ClassListObj.Eof
		if RsEditObj("SysClass") = ClassListObj("ClassID") then
			SelectStr = "selected"
		else
			SelectStr = ""
		end if
		ClassList = ClassList & "<option " & SelectStr & " value="&ClassListObj("ClassID")&"" & ">" & ClassListObj("ClassCName") & "</option><br>"
		ClassList = ClassList & ChildClassList(ClassListObj("ClassID"),"")
		ClassListObj.MoveNext	
	loop
	ClassListObj.Close
	Set ClassListObj = Nothing
End Function

Function SiteFolderList()
	Dim ClassListObj,SelectStr
	Set ClassListObj = CollectConn.Execute("Select * from FS_SiteFolder order by ID desc")
	do while Not ClassListObj.Eof
		if RsEditObj("Folder") = ClassListObj("ID") then
			SelectStr = "selected"
		else
			SelectStr = ""
		end if
		SiteFolderList = SiteFolderList & "<option " & SelectStr & " value="&ClassListObj("ID")&"" & ">&nbsp;&nbsp;|--" & ClassListObj("SiteFolder") & "</option><br>"
		ClassListObj.MoveNext	
	loop
	ClassListObj.Close
	Set ClassListObj = Nothing
End Function

Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr,SelectStr
	Set TempRs = Conn.Execute("Select ClassID,ClassCName from FS_NewsClass where ParentID='" & ClassID & "' order by ClassID desc")
	TempStr = Temp & " |- "
	do while Not TempRs.Eof
		if RsEditObj("SysClass") = TempRs("ClassID") then
			SelectStr = "selected"
		else
			SelectStr = ""
		end if
		ChildClassList = ChildClassList & "<option " & SelectStr & " value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option><br>"
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function

Dim TempletDirectory
if SysRootDir <> "" then
	TempletDirectory = "/" & SysRootDir & "/" & TempletDir
else
	TempletDirectory = "/" & TempletDir
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自动新闻采集—站点设置</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body leftmargin="2" topmargin="2">
<form name="Form" method="post" action="SiteTwoStep.asp">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="45" align="center" alt="第二步" onClick="CheckData();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下一步</td>
			<td width=2 class="Gray">|</td>
		    <td width="35" align="center" alt="后退" onClick="history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp; <input name="SiteID" type="hidden" id="SiteID" value="<% = SiteID %>"> 
              <input name="Result" type="hidden" id="Result" value="Edit"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E6E6E6">
    <tr> 
      <td width="100" height="26" bgcolor="#F5F5F5"> 
        <div align="right">站点名称</div></td>
      <td bgcolor="#FFFFFF"> 
        <input name="SiteName" style="width:100%;" type="text" id="SiteName" value="<%=RsEditObj("sitename")%>"> 
        <div align="right"> </div></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#F5F5F5"> 
        <div align="right">采集对象页</div></td>
      <td bgcolor="#FFFFFF"> 
        <input name="objURL" type="text" id="textarea" style="width:100%;" value="<%=RsEditObj("objURL")%>" size="50"></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#F5F5F5"> 
        <div align="right">新闻摸板</div></td>
      <td bgcolor="#FFFFFF">
<input readonly name="SysTemplet" type="text" id="SysTemplet" style="width:80%;" value="<%=RsEditObj("SysTemplet")%>">
        <input name="Submitaaa" type="button" id="Submitaaa" value="选择模板" onClick="OpenWindowAndSetValue('../../FunPages/SelectFileFrame.asp?CurrPath=<% = TempletDirectory %>',400,300,window,document.Form.SysTemplet);"></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#F5F5F5"> 
        <div align="right">采集站点分类</div></td>
      <td bgcolor="#FFFFFF">
<select name="SiteFolder" style="width:100%;" id="SiteFolder">
          <option value="0">根栏目</option>
          <% = SiteFolderList %>
        </select></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#F5F5F5"> 
        <div align="right">入库目标栏目</div></td>
      <td bgcolor="#FFFFFF">
<select name="SysClass" style="width:100%;" id="select2">
          <% =TempClassListStr %>
        </select></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#F5F5F5"> 
        <div align="right">采集参数</div></td>
      <td bgcolor="#FFFFFF"> 锁定 
        <input name="islock" type="checkbox" id="islock" value="1" <%if RsEditObj("islock")=true then response.Write("checked")%>>
        保存远程图片 
        <input type="checkbox" name="SaveRemotePic" value="1" <%if RsEditObj("SaveRemotePic")=true then response.Write("checked")%>>
        新闻是否已经审核 
        <input type="checkbox" name="Audit" value="1" <%if RsEditObj("Audit")=true then response.Write("checked")%>>
        是否倒序采集 
        <input name="IsReverse" type="checkbox" id="IsReverse" value="1" <%if RsEditObj("IsReverse")="1" then response.Write("checked")%>></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#F5F5F5">
<div align="right">保存图片路径</div></td>
      <td bgcolor="#FFFFFF">
<input type="text" readonly name="SaveIMGPath" style="width:80%;" value="<% = RsEditObj("SaveIMGPath") %>"> 
        <input name="Submit111" id="SelectPath" type="button" value="选择路径" onClick="OpenWindowAndSetValue('../../FunPages/SelectPathFrame.asp?CurrPath=<% = SelectPath %>',400,300,window,document.Form.SaveIMGPath);"></td>
    </tr>
    <tr> 
      <td height="26" bgcolor="#F5F5F5">
<div align="right">过滤选项</div></td>
      <td bgcolor="#FFFFFF">HTML 
        <input type="checkbox" name="TextTF" value="1" <% if RsEditObj("TextTF") = True then Response.Write("checked")%>>
        STYLE 
        <input type="checkbox" name="IsStyle" value="1" <% if RsEditObj("IsStyle") = True then Response.Write("checked")%>>
        DIV
        <input type="checkbox" name="IsDiv" value="1" <% if RsEditObj("IsDiv") = True then Response.Write("checked")%>>
        A
        <input type="checkbox" name="IsA" value="1" <% if RsEditObj("IsA") = True then Response.Write("checked")%>>
        CLASS
        <input type="checkbox" name="IsClass" value="1" <% if RsEditObj("IsClass") = True then Response.Write("checked")%>>
        FONT
        <input type="checkbox" name="IsFont" value="1" <% if RsEditObj("IsFont") = True then Response.Write("checked")%>>
        SPAN
        <input type="checkbox" name="IsSpan" value="1" <% if RsEditObj("IsSpan") = True then Response.Write("checked")%>>
        OBJECT
        <input type="checkbox" name="IsObject" value="1" <% if RsEditObj("IsObject") = True then Response.Write("checked")%>>
        IFRAME
        <input type="checkbox" name="IsIFrame" value="1" <% if RsEditObj("IsIFrame") = True then Response.Write("checked")%>>
        SCRIPT
        <input type="checkbox" name="IsScript" value="1" <% if RsEditObj("IsScript") = True then Response.Write("checked")%>> 
      </td>
    </tr>
  </table>
</form>
</body>
</html><%
Set Conn = Nothing
Set CollectConn = Nothing
Set RsEditObj = Nothing
%>
<script language="JavaScript">
function CheckData()
{
	if (document.Form.SiteName.value==''){alert('没有填写站点名称');document.Form.SiteName.focus();return;}
	if (document.Form.objURL.value==''){alert('没有填写采集对象页');document.Form.objURL.focus();return;}
	if (document.Form.SysClass.value==''){alert('没有填写目标栏目');document.Form.SysClass.focus();return;}
	document.Form.submit();
}
</script>