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
'==============================================================================
   Dim UrlAdd
   UrlAdd=Request.QueryString("Url")
   Session("MemName") = ""
   Session("sName") = ""
   Session("MemPassword") = ""
   Session("MemID") = ""
   Session("email") = ""
   Session("mGroupId") = ""
   Session("RePassWord") =""
   Response.Cookies("Foosun")("MemName")=""
   Response.Cookies("Foosun")("MemPassword")=""
   Response.Cookies("Foosun")("MemID")=""
   Response.Cookies("Foosun")("GroupID")=""
	If UrlAdd<>"" then
		response.redirect UrlAdd
	Else
		Response.Redirect "../../"
	End If
   Response.End
%>