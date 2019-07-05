<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P010601") then Call ReturnError1()
    Dim TempClassID,OldClassObj,OldClassEName
	If trim(Request("ClassID"))<>"" and Request("ClassID")<>"0" then
	   TempClassID = CStr(Request("ClassID"))
	   TempClassID = Replace(Replace(Replace(Replace(Replace(TempClassID,"'",""),"and",""),"select",""),"or",""),"union","")
		Set OldClassObj = Conn.Execute("select ClassID from FS_NewsClass where ClassID='"&TempClassID&"'")
		if OldClassObj.Eof then
		   Response.Write("<script>alert(""栏目参数传递错误"");dialogArguments.location.reload();window.close();</script>")
		   Response.End
		end if
		OldClassObj.Close
		Set OldClassObj = Nothing
	 else
	   Response.Write("<script>alert(""请先选种栏目，投稿管理中才可以新建投稿"");</script>")
	   Response.Write("<script>location.href='Contributionlist.asp'</script>")

	End If
	
Dim NewsContent
NewsContent = Replace(Replace(Request.Form("Content"),"""","%22"),"'","%27")
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>稿件添加</title>
</head>
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<body topmargin="2" leftmargin="2">
<form action="" name="NewsForm" method="post">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
           	  <td width="35" align="center" alt="保存" onClick="SubmitFun();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">保存</td>
			  <td width=2 class="Gray">|</td>
			  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
            <td>&nbsp;<input type="hidden" name="Content" value="<% =  NewsContent %>">
          <input name="action" type="hidden" id="action" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr> 
      <td width="8%">稿件标题</td>
      <td><input name="Title" type="text" id="Title" style="width:90%" value="<%=Request("Title")%>"> 
        <input name="ClassID" type="hidden" id="ClassID" value="<%=TempClassID%>"></td>
      <td width="8%">副 标 题</td>
      <td width="42%"><input name="SubTitle" type="text" id="SubTitle2" style="width:91.5%" value="<%=Request("SubTitle")%>"></td>
    </tr>
    <tr> 
      <td>新闻作者</td>
      <td><input name="Author" type="text" style="width:90%" value="<%=Request("Author")%>"></td>
      <td>关 键 字</td>
      <td><input name="KeyWords" type="text" style="width:91.5%" value="<%=Request("KeyWords")%>"> 
      </td>
    </tr>
    <tr> 
      <td colspan="4"><div align="center"> 
          <iframe id='NewsContent' src="../../Editer/NewsEditer.asp" frameborder=0 scrolling=no width='100%' height='460'></iframe>
        </div></td>
    </tr>
</table>
</form>
</body>
</html>
<script language="javascript">

function SubmitFun()
{
	frames["NewsContent"].SaveCurrPage();
	var TempContentArray=frames["NewsContent"].NewsContentArray;
	document.NewsForm.Content.value='';
	for (var i=0;i<TempContentArray.length;i++)
	{
		if (TempContentArray[i]!='')
		{
			if (document.NewsForm.Content.value=='') document.NewsForm.Content.value=TempContentArray[i];
			else document.NewsForm.Content.value=document.NewsForm.Content.value+'[Page]'+TempContentArray[i];
		} 
	}
	document.NewsForm.submit();
}
</script>
<%
  if Request.Form("action")="add" then
     Dim ITitle,IClassID,INewsTemplet,IClickNum,IAddDate,INewsAddObj,INewsAddSql
     if Request.Form("Title")<>"" then
		ITitle = Replace(Replace(Request.Form("Title"),"""",""),"'","")
	 else
	    Response.Write("<script>alert('请输入新闻标题');</script>")
		Response.End
	 end if
     if Request.Form("ClassID")<>"" then
		IClassID = Replace(Replace(Request.Form("ClassID"),"""",""),"'","")
	 else
	    Response.Write("<script>alert('栏目参数传递错误');</script>")
		Response.End
	 end if
	 if Request.Form("Content")="" or isnull(Request.Form("Content")) then
	    Response.Write("<script>alert('请输入新闻内容');</script>")
		Response.End
	 end if
	 if Request.Form("Author")="" or isnull(Request.Form("Author")) then
	    Response.Write("<script>alert('请留下您的大名');</script>")
		Response.End
	 end if
	  set INewsAddObj=server.createobject(G_FS_RS)
	  INewsAddSql="select * from FS_Contribution where 1=0"
	  INewsAddObj.open INewsAddSql,Conn,3,3
	  INewsAddObj.addnew
	  INewsAddObj("Title") =  ITitle
	  If Request.Form("SubTitle")<>"" then
		  INewsAddObj("SubTitle") = Replace(Replace(Request.Form("SubTitle"),"""",""),"'","")
	   end if
	  INewsAddObj("ClassID") =  IClassID
	  INewsAddObj("Content") =  Request.Form("Content")   '新闻内容 尚未判断
	  INewsAddObj("ContID") =  GetRandomID18    '稿件ID
	  INewsAddObj("AddTime") =  Now()
	  if Request.Form("KeyWords") <> "" then 
		  INewsAddObj("KeyWords") = Replace(Replace(Request.Form("KeyWords"),"""",""),"'","")
	  end if
	  INewsAddObj("Author") = Replace(Replace(Request.Form("Author"),"""",""),"'","")
	  INewsAddObj.Update
	  INewsAddObj.Close
	  Set INewsAddObj = Nothing
	  Response.Redirect("ContributionList.asp?ClassID=" & TempClassID)
  end if
%>