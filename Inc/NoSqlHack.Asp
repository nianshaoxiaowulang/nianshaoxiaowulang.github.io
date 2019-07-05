<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System 
'模块功能: SQL通用防注入模块
'最新更新：2005.8.4 By lzp
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-606、607,客户支持：608
'产品咨询QQ：159410,655071,66252421
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim FS_NoSqlHack_AllStr,FS_NoSqlHack_Str,FS_NoSqlHack_ComeUrlGet,FS_NoSqlHack_ComeUrlPost,FS_NoSqlHack_Get,FS_NoSqlHack_Post,FS_NoSqlHack_i
'On Error Resume Next
FS_NoSqlHack_AllStr="'|;| and |chr(|exec |insert |select |delete from|update |mid(|master."
FS_NoSqlHack_ComeUrlGet   = Request.QueryString
FS_NoSqlHack_ComeUrlPost  = Request.Form
FS_NoSqlHack_Str = Split(FS_NoSqlHack_AllStr,"|")
'Post
If FS_NoSqlHack_ComeUrlPost<>"" then
	For Each FS_NoSqlHack_Post In Request.Form
		For FS_NoSqlHack_i = 0 To Ubound(FS_NoSqlHack_Str)
			If Instr(LCase(Request.Form(FS_NoSqlHack_Post)),FS_NoSqlHack_Str(FS_NoSqlHack_i))<>0 Then
				Response.End
			End if
		Next
	Next
End if
'Get
If FS_NoSqlHack_ComeUrlGet<>"" then 
	For Each FS_NoSqlHack_Get In Request.QueryString
		For FS_NoSqlHack_i = 0 To Ubound(FS_NoSqlHack_Str)
			If Instr(LCase(Request.QueryString(FS_NoSqlHack_Get)),FS_NoSqlHack_Str(FS_NoSqlHack_i))<>0 Then
				Response.End
			End if
		Next
	Next
End if
%>
