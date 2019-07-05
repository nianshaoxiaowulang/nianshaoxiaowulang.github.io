<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
'==============================================================================
'软件名称：FoosunShop System Form FoosunCMS
'当前版本：Foosun Content Manager System 3.0 系列
'最新更新：2004.12
'==============================================================================
'商业注册联系：028-85098980-601,602 技术支持：028-85098980-605、607,客户支持：608
'产品咨询QQ：159410,394226379,125114015,655071
'技术支持:所有程序使用问题，请提问到bbs.foosun.net我们将及时回答您
'程序开发：风讯开发组 & 风讯插件开发组
'论坛支持：风讯在线论坛(http://bbs.foosun.net)
'官方网站：www.Foosun.net  演示站点：test.cooin.com    
'网站建设专区：www.cooin.com
'==============================================================================
'免费版本请在新闻首页保留版权信息，并做上本站LOGO友情连接
'==============================================================================
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Dim RsAdminConfigObj
Set RsAdminConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop,MaxContent,QPoint from FS_Config")
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070704") then Call ReturnError1()
If request("Action")="Del" Then
	If trim(Request("GID"))<>"" Then
		Dim BookStr,ParaArray,i,NumStr,NumParaArray
		BookStr = Request("GID")
		  if Right(BookStr,1)="," then
			BookStr = Left(BookStr,Len(BookStr)-1)
		  end if
		  if Left(BookStr,1)="," then
			BookStr = Right(BookStr,Len(BookStr)-1)
		  end if
		  ParaArray = Split(BookStr,",")
		For i = LBound(ParaArray) to UBound(ParaArray)
			Dim GBListObj
			Set GBListObj = Conn.execute("Select ID,UserID From FS_GBook where id="&Clng(ParaArray(i)))
		    Conn.execute("Update FS_Members Set Point = Point-"&RsAdminConfigObj("QPoint")&"  where ID="&GBListObj("UserID"))
			Conn.execute("Delete From FS_GBook Where id="&Clng(ParaArray(i)))
		Next
		Response.Write("<script>alert(""删除成功！"&CopyRight&""");location=""Sysbook.asp"";</script>")  
		Response.End
		'扣除积分
	Else
		Response.Write("<script>alert(""请选择删除的帖子！"&CopyRight&""");location=""Sysbook.asp"";</script>")  
		Response.End
	End if
End If
%>