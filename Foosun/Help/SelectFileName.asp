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
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'�������2�ο��������뾭����Ѷ��˾������������׷����������
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
		alert('���ϲ�ѯ ҳ�湦�ܡ�<%=FuncName%>����ҳ���ļ���<%=FileName%>����û���ҵ��������ݣ�');
	}
-->
</script>
<%Set HelpConn = nothing%>