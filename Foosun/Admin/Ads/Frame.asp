<% Option Explicit %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% = Request("PageTitle") %></title>
<%
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
Dim RequestItem,ParaList,FileName,Url
ParaList = ""
For Each RequestItem In Request.QueryString
	if RequestItem <> "FileName" And RequestItem <> "PageTitle" then
		if ParaList = "" then
			ParaList = RequestItem & "=" & Request.QueryString(RequestItem)
		else
			ParaList = ParaList & "&" & RequestItem & "=" & Request.QueryString(RequestItem)
		end if
	end if
Next
FileName = Request("FileName")
if FileName <> "" then
	Url = FileName & "?" & ParaList
else
	%>
	<script language="JavaScript">
		alert('�ļ�������');
		window.close();
	</script>
	<%
	Response.End
end if
%>
</head>
<body scrolling=no bgcolor="#E6E6E6" topmargin="0" leftmargin="0">
<iframe src=<% = Url %> style="width:100%;height:100%;"  frameborder=0 scrolling="auto"></iframe>
</body>
</html>
