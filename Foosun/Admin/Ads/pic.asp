<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/NosqlHack.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
Dim DBC,Conn,TempSysRootDir,pic1,rsObj,picSql
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
	if SysRootDir = "" then
		TempSysRootDir = ""
	else
		TempSysRootDir = "/" & SysRootDir
end if
pic1=replace(request("pic"),"'","")
set rsObj=server.createobject(G_FS_RS)
	  picSql="select * from FS_Ads where Location="&pic1&""
	  rsObj.open picSql,Conn,1,1
	  if rsobj.eof then
	     Response.Write"参数传递错误!!!"
	  else 
%>
      <a href="<%=rsobj("url")%>" target="_blank"><img src="<%=TempSysRootDir%><%=rsobj("LeftPicPath")%>" border="0"></img></a>
<% end if
   rsObj.close
   set rsObj=nothing
 %>
