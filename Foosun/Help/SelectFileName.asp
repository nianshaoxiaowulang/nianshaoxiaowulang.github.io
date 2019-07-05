<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim DBC,Conn,HelpConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + Server.MapPath("Foosun_help.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set HelpConn = DBC.OpenConnection()
Set DBC = Nothing
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

Dim FuncName,FileName,SingleContentArray,ContentArray,oPageFieldArray
FuncName = Request.QueryString("FuncName")
FileName = Request.QueryString("FileName")

Dim ZoneRs,PageField,SingleContent,Content,strSQL
Set ZoneRs = Server.CreateObject(G_FS_RS)

strSQL = "Select * From [FS_Help] Where FileName='"&FileName&"'"
If FuncName<>"" Then strSQL=" and FuncName='" &FuncName& "'"

ZoneRs.open strSQL,HelpConn,1,1
do while not ZoneRs.eof
		PageField = PageField & "{FS_Help_Split}" & ZoneRs("PageField")
		SingleContent = SingleContent & "{FS_Help_Split}" & ZoneRs("HelpSingleContent")
		Content = Content & "{FS_Help_Split}" & ZoneRs("HelpContent")
ZoneRs.movenext
loop
ZoneRs.close
Set ZoneRs = Nothing

set Conn = Nothing

PageField = Replace(PageField,vbcrlf,"")
PageField = Replace(PageField,"'","\'")

SingleContent = Replace(SingleContent,vbcrlf,"")
SingleContent = Replace(SingleContent,"'","\'")

Content = Replace(Content,vbcrlf,"")
Content = Replace(Content,"'","\'")

%>
<script language="javascript">
<!--
	var PageField = '<%=PageField%>';
	if(PageField != ''){
		parent.oPageFieldArray = '<%=PageField%>'.split("{FS_Help_Split}");
		parent.SingleContentArray = '<%=SingleContent%>'.split("{FS_Help_Split}");
		parent.ContentArray = '<%=Content%>'.split("{FS_Help_Split}");
		var oItem = parent.HelpForm.PageField;
		oItem.length = 0;
		for(var i=0;i<parent.oPageFieldArray.length;i++)
		{
			var opt = parent.document.createElement("OPTION");
			if(parent.oPageFieldArray[i]!='')
			{
				opt.text = parent.oPageFieldArray[i];
				opt.value = parent.oPageFieldArray[i];
				oItem.add(opt);
			}
		}
	}else{
		alert('联合查询 页面功能“<%=FuncName%>”和页面文件“<%=FileName%>”，没有找到帮助内容！');
	}
-->
</script>
<%Set HelpConn = nothing%>