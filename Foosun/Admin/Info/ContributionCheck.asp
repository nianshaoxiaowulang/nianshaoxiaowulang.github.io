<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../Refresh/Function.asp" -->
<!--#include file="../Refresh/RefreshFunction.asp" -->
<!--#include file="../Refresh/SelectFunction.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P010603") then Call ReturnError1()
    Dim NewsID,TempClassID,OldClassObj,OldClassEName,RsContObj
	    NewsID = Cstr(Request("NewsID"))
		NewsID = Replace(Replace(Replace(Replace(Replace(NewsID,"'",""),"and",""),"select",""),"or",""),"union","")
		Set RsContObj = Conn.Execute("Select * from FS_Contribution where ContID = '"&NewsID&"'")
		If RsContObj.eof then
		   Response.Write("<script>alert(""参数传递错误"");dialogArguments.location.reload();window.close();</script>")
		   Response.End
		End If
	Dim DummyPath_Riker
	If SysRootDir<>"" then
		DummyPath_Riker = "/" & SysRootDir
	Else
		DummyPath_Riker = ""
	End If
	
    Set OldClassObj = Conn.Execute("select ClassID,ClassEName from FS_NewsClass where ClassID='"&RsContObj("ClassID")&"'")
	if Not OldClassObj.Eof then
		OldClassEName = OldClassObj("ClassEName")
	end if
	OldClassObj.Close
	Set OldClassObj = Nothing
	
dim TempClassListStr
TempClassListStr = ClassList
Function ClassList()
	Dim ClassListObj,SelectStr
	Set ClassListObj = Conn.Execute("select * from FS_newsclass where ParentID = '0'")
	do while Not ClassListObj.Eof
		if OldClassEName = ClassListObj("ClassEName") then
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

Function ChildClassList(ClassID,Temp)
	Dim TempRs,TempStr,SelectStr
	Set TempRs = Conn.Execute("Select * from FS_NewsClass where ParentID = '" & ClassID & "'")
	TempStr = Temp & " - "
	do while Not TempRs.Eof
		if OldClassEName = TempRs("ClassEName") then
			SelectStr = "selected"
		else
			SelectStr = ""
		end if
		if TempRs("ChildNum") = 0 then
			ChildClassList = ChildClassList & "<option " & SelectStr & " value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option><br>"
		else
			ChildClassList = ChildClassList & "<option " & SelectStr & " value="&TempRs("ClassID")&"" & ">" & TempStr & TempRs("ClassCName") & "</option><br>"
		end if
		ChildClassList = ChildClassList & ChildClassList(TempRs("ClassID"),TempStr)
		TempRs.MoveNext
	loop
	TempRs.Close
	Set TempRs = Nothing
End Function
	
Dim NewsContent
    If Request.Form("Content")<>"" then
		NewsContent = Replace(Replace(Request.Form("Content"),"""","%22"),"'","%27")
	Else
		NewsContent = Replace(Replace(RsContObj("Content"),"""","%22"),"'","%27")
	End If
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>稿件审核</title>
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
            <td>&nbsp; <input type="hidden" name="Content" value="<% = NewsContent %>"> 
              <input name="action" type="hidden" id="action2" value="add"></td>
        </tr>
      </table>
	  </td>
  </tr>
</table><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="4"></td>
  </tr>
</table>

  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#E4E4E4">
    <tr bgcolor="#FFFFFF"> 
      <td width="9%" height="26"> 
        <div align="center">新闻标题</div></td>
      <td width="38%"> 
        <input name="Title" type="text" id="Title" style="width:90%" value="<%=LoseHtml(RsContObj("Title"))%>"></td>
      <td width="9%"> 
        <div align="center">副 标 题</div></td>
      <td width="41%"> 
        <input name="SubTitle" type="text" id="SubTitle2" style="width:90%" value="<%=RsContObj("SubTitle")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">所属栏目</div></td>
      <td> 
        <select name="ClassID" id="select2" style="width:90%">
          <% =TempClassListStr %>
        </select></td>
      <td> 
        <div align="center">浏览权限</div></td>
      <td> 
        <select name="BrowPop" style="width:90%" onChange="ChooseExeName();">
          <option value="" selected> </option>
          <%
		Dim BrowPopObj
		set BrowPopObj = Conn.Execute("Select Name,PopLevel from FS_MemGroup order by PopLevel asc")
		while not BrowPopObj.eof
		%>
          <option value="<%=BrowPopObj("PopLevel")%>"><%=BrowPopObj("Name")%></option>
          <%
		BrowPopObj.Movenext
		Wend
		BrowPopObj.Close
		Set BrowPopObj = Nothing
		%>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">新闻模板</div></td>
      <td> 
        <input name="NewsTemplet" type="text" id="NewsTemplet" size="24" readonly value="<%If Request("NewsTemplet")="" then Response.Write("/"&TempletDir&"/NewsClass/News.htm") else Response.Write(Request("NewsTemplet"))%>" > 
        <input type="button" name="Submit" value="选择模板" onClick="OpenWindowAndSetValue('../../Funpages/SelectFileFrame.asp?CurrPath=<%=DummyPath_Riker%>/<% = TempletDir %>',400,300,window,document.NewsForm.NewsTemplet);document.NewsForm.NewsTemplet.focus();"> 
      </td>
      <td> 
        <div align="center">图片地址</div></td>
      <td> 
        <input name="PicPath" type="text" id="PicPath" size="27" value="<%=Request("PicPath")%>" > 
        <input type="button" name="PPPChoose" value="选择图片" onClick="OpenWindowAndSetValue('../../FunPages/SelectPic.asp?CurrPath=/<%= UpFiles %>',550,290,window,document.NewsForm.PicPath);"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">所属专题</div></td>
      <td> 
        <input name="SpecialIDText" type="text" size="24" readonly value="<%=Request("SpecialIDText")%>"> 
        <input type="hidden" name="SpecialID" value="<%=Request("SpecialID")%>"> 
        <select name="SelSpecialID" style="width:27%" onChange=ChooseSpecial(this.options[this.selectedIndex].value)>
          <option value="" selected> </option>
          <option value="<%=""&"***"&"Clean"%>" style="color:red">清空</option>
          <%
		Dim SpecialIDObj
		set SpecialIDObj = Conn.Execute("Select CName,SpecialID from FS_Special order by ID desc")
		while not SpecialIDObj.eof
		%>
          <option value="<%=SpecialIDObj("CName")&"***"&SpecialIDObj("SpecialID")%>"><%=SpecialIDObj("CName")%></option>
          <%
		SpecialIDObj.Movenext
		Wend
		SpecialIDObj.Close
		Set SpecialIDObj = Nothing
		%>
        </select></td>
      <td> 
        <div align="center">新闻来源</div></td>
      <td> 
        <input name="TxtSourceText" type="text" size="27" readonly value="<%=Request("TxtSource")%>"> 
        <input type="hidden" name="TxtSource" value= "<%=Request("TxtSource")%>"> 
        <select name="SelTxtSource" style="width:25%" onChange="Dosusite(this.options[this.selectedIndex].value);">
          <option value="" selected> </option>
          <option value="<%=""&"***"&"Clean"%>" style="color:red">清空</option>
          <option value="本站原创">本站原创</option>
          <option value="不详">不详</option>
          <%
		Dim TxtSourceObj
		set TxtSourceObj = Conn.Execute("select * from FS_Routine where Type=2 order by ID desc")
		While not TxtSourceObj.eof
		%>
          <option value="<%=TxtSourceObj("Name")&"***"&TxtSourceObj("Url")%>"><%=TxtSourceObj("Name")%></option>
          <%
		TxtSourceObj.Movenext
		Wend
		TxtSourceObj.Close
		Set TxtSourceObj = Nothing
		%>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">关 键 字</div></td>
      <td> 
        <input name="KeywordText" type="text" size="24" readonly value="<%=RsContObj("KeyWords")%>"> 
        <input type="hidden" name="KeyWords" value="<%=RsContObj("KeyWords")%>"> 
        <select name="SelKeyWords" style="width:27%" onChange=Dokesite(this.options[this.selectedIndex].value)>
          <option value="" selected> </option>
          <option value="Clean" style="color:red">清空</option>
          <%
		Dim KeyWordsObj
		set KeyWordsObj = Conn.Execute("select * from FS_Routine where Type=1 order by ID desc")
		while not KeyWordsObj.eof
		%>
          <option value="<%=KeyWordsObj("Name")%>"><%=KeyWordsObj("Name")%></option>
          <%
		KeyWordsObj.Movenext
		Wend
		KeyWordsObj.Close
		Set KeyWordsObj = Nothing
		%>
        </select></td>
      <td> 
        <div align="center">责任编辑</div></td>
      <td> 
        <input name="EditerText" type="text" size="27" readonly value="<%=Request("EditerText")%>"> 
        <input type="hidden" name="Editer" value="<%=Request("Editer")%>"> 
        <select name="Editer1" style="width:25%"  onChange="Editsite(this.options[this.selectedIndex].value);">
          <option value="" selected> </option>
          <option value="<%=""&"***"&"Clean"%>" style="color:red">清空</option>
          <%
		Dim EditerObj
		Set EditerObj = Conn.Execute("Select * from FS_Routine where Type=4 order by ID desc")
		while not EditerObj.eof
		%>
          <option value="<%=EditerObj("Name")&"***"&EditerObj("Url")%>"><%=EditerObj("Name")%></option>
          <%
		EditerObj.Movenext
		Wend
		EditerObj.Close
		Set EditerObj = Nothing
		%>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">新闻作者</div></td>
      <td> 
        <input name="AuthorText" type="text" size="24" readonly value="<%=RsContObj("Author")%>"> 
        <input type="hidden" name="Author" value="<%=RsContObj("Author")%>"> 
        <select name="SelAuthor" id="SelAuthor" style="width:27%" onChange="Doauthsite(this.options[this.selectedIndex].value);" disabled>
          <option value="" selected> </option>
          <option value="<%=""&"***"&"Clean"%>" style="color:red">清空</option>
          <option value="佚名">佚名</option>
          <option value="本站">本站</option>
          <option value="未知">未知</option>
          <%
		Dim AuthorObj
		set AuthorObj = Conn.Execute("select * from FS_Routine where Type=3 order by ID desc")
		while not AuthorObj.eof
		%>
          <option value="<%=AuthorObj("Name")&"***"&AuthorObj("Url")%>"><%=AuthorObj("Name")%></option>
          <%
		AuthorObj.Movenext
		Wend
		AuthorObj.Close
		Set AuthorObj=nothing
		%>
        </select></td>
      <td> 
        <div align="center">添加日期</div></td>
      <td> 
        <input name="AddDate" type="text" id="AddDate3" style="width:90%" value="<% if RsContObj("AddTime")="" then Response.Write(now()) else Response.Write(RsContObj("AddTime")) end if%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">点击次数</div></td>
      <td> 
        <input name="ClickNum" type="text" id="ClickNum" style="width:90%" value="<%if Request("ClickNum")="" then Response.Write("0") else Response.Write(Request("ClickNum")) end if %>"></td>
      <td rowspan="2"> 
        <div align="center">导航文字</div></td>
      <td rowspan="2"> 
        <textarea name="NaviWords" rows="2" id="NaviWords2" style="width:90%"><%=Request("NaviWords")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="26"> 
        <div align="center">扩 展 名</div></td>
      <td> 
        <select name="FileExtName" id="select" style="width:90%">
          <option value="htm">htm</option>
          <option value="html">html</option>
          <option value="shtm">shtm</option>
          <option value="shtml">shtml</option>
          <option value="asp">asp</option>
        </select></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="26"> <div align="center">图片新闻 
                <input type="checkbox" name="PicNewsTF" value="1" <%If Request("PicNewsTF")="1" then Response.Write("checked") end if%> onClick="ChoosePicType();">
              </div></td>
            <td><div align="center">滚动新闻 
                <input name="MarqueeNews" type="checkbox" id="MarqueeNews2" value="1" <%if Request("MarqueeNews")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">允许评论 
                <input name="ReviewTF" type="checkbox" id="ReviewTF2" value="1" onClick="ChooseRiview();" <%if Request("ReviewTF")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">显示评论 
                <input name="ShowReviewTF" type="checkbox" id="ShowReviewTF2" value="1" disabled <%if Request("ShowReviewTF")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">公告新闻 
                <input name="ProclaimNews" type="checkbox" id="ProclaimNews2" value="1" <%if Request("ProclaimNews")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">幻 灯 片 
                <input name="FilterNews" type="checkbox" id="FilterNews" value="1" <%if Request("FilterNews")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">焦点图片 
                <input name="FocusNewsTF" type="checkbox" id="FocusNewsTF" value="1" <%if Request("FocusNewsTF")=1 then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">今日头条 
                <input name="TodayNewsTF" type="checkbox" id="TodayNewsTF" value="1">
              </div></td>
          </tr>
          <tr> 
            <td height="26"><div align="center">推荐新闻 
                <input name="RecTF" type="checkbox" id="RecTF2" value="1" <%if Request("RecTF")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">是否审核 
                <input name="AuditTF" type="checkbox" id="AuditTF2" value="1" <%if Request("AuditTF")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">内部链接 
                <input name="LinkTF" type="checkbox" id="LinkTF2" value="1" <%if Request("LinkTF")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">用户投稿 
                <input name="ManuTF" type="checkbox" id="ManuTF2" value="1" disabled checked <%if Request("ManuTF")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">远程存图 
                <input type="checkbox" name="SavePic" value="1" <%if Request("SavePic")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">并排新闻 
                <input name="SBSNews" type="checkbox" id="SBSNews" value="1" <%if Request("SBSNews")="1" then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">精彩回顾 
                <input name="ClassicalNewsTF" type="checkbox" id="ClassicalNewsTF" value="1" <%if Request("ClassicalNewsTF")=1 then Response.Write("checked") end if%>>
              </div></td>
            <td><div align="center">备用参数 
                <input name="BackUp" type="checkbox" id="BackUp" value="1">
              </div></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4"> 
        <div align="center"> 
          <iframe id='NewsContent' src="../../Editer/NewsEditer.asp" frameborder=0 scrolling=no width='100%' height='460'></iframe>
        </div></td>
    </tr>
</table>
</form>
</body>
</html>
<script language="javascript">
 function ChooseRiview()
   {
      if (document.NewsForm.ReviewTF.checked==true) document.NewsForm.ShowReviewTF.disabled=false;
      else document.NewsForm.ShowReviewTF.disabled=true;
	}

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

function ChoosePicType()
{
  if (document.NewsForm.PicNewsTF.checked==true)
     {
	   document.NewsForm.PicPath.disabled=false;
	   document.NewsForm.PPPChoose.disabled=false;
	   document.NewsForm.NaviWords.disabled=false;
	   document.NewsForm.FilterNews.disabled=false;
	   document.NewsForm.FocusNewsTF.disabled=false;
	   document.NewsForm.ClassicalNewsTF.disabled=false;
	  }
   else
      {
	   document.NewsForm.PicPath.disabled=true;
	   document.NewsForm.PPPChoose.disabled=true;
	   document.NewsForm.NaviWords.disabled=true;
	   document.NewsForm.FilterNews.disabled=true;
	   document.NewsForm.FocusNewsTF.disabled=true;
	   document.NewsForm.ClassicalNewsTF.disabled=true;
	   }
 }
 
function ChooseExeName()
{
  if (document.NewsForm.BrowPop.value!='') document.NewsForm.FileExtName.disabled=true;
  else document.NewsForm.FileExtName.disabled=false;
 }
 
ChooseRiview();
ChoosePicType();
ChooseExeName();
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
     if Request.Form("NewsTemplet")<>"" then
		INewsTemplet = Replace(Replace(Request.Form("NewsTemplet"),"""",""),"'","")
	 else
	    Response.Write("<script>alert('请选择新闻模板文件');</script>")
		Response.End
	 end if 
     if Isnumeric(Request.Form("ClickNum")) then
		IClickNum = Clng(Request.Form("ClickNum"))
	 else
	    Response.Write("<script>alert('新闻初始点击次数必须为数字类型');</script>")
		Response.End
	 end if 
	 if IsDate(Request.Form("AddDate")) then
	 	IAddDate = Formatdatetime(Request.Form("AddDate"))
	 else
	    Response.Write("<script>alert('新闻添加时间类型错误,请重新输入');</script>")
		Response.End
	 end if
	 if Request.Form("PicNewsTF")="1" and Request.Form("PicPath")="" then
	    Response.Write("<script>alert('请输入图片地址');</script>")
		Response.End
	 end if
	 if Request.Form("Content")="" or isnull(Request.Form("Content")) then
	    Response.Write("<script>alert('请输入新闻内容');</script>")
		Response.End
	 end if
	Dim ConCheckNewsID
	ConCheckNewsID = GetRandomID18
	
	Dim NewsFileNames,RsNewsConfigObj
		Set RsNewsConfigObj = Conn.Execute("Select DoMain,NewsFileName from FS_Config")
		NewsFileNames = NewsFileName(RsNewsConfigObj("NewsFileName"),IClassID,ConCheckNewsID)

	  set INewsAddObj=server.createobject(G_FS_RS)
	  INewsAddSql="select * from FS_News"
	  INewsAddObj.open INewsAddSql,Conn,3,3
	  INewsAddObj.addnew
	  INewsAddObj("Title") =  ITitle
	  If Request.Form("SubTitle")<>"" then
		  INewsAddObj("SubTitle") = Replace(Replace(Request.Form("SubTitle"),"""",""),"'","")
	   end if
	  INewsAddObj("ClassID") =  IClassID
	  INewsAddObj("HeadNewsTF") =  "0"
'	  INewsAddObj("Content") = ReplaceRemoteUrl(Request.Form("Content"),UpFiles & BeyondPicDir,RsNewsConfigObj("DoMain"),DummyPath_Riker)   '新闻内容 尚未判断
	  INewsAddObj("NewsID") =  ConCheckNewsID    '新闻ID
	  INewsAddObj("ManuTF") =  "1"
	  INewsAddObj("FileName") = NewsFileNames   '新闻文件名 
	  if Request.Form("BrowPop") <> "" then
		  INewsAddObj("FileExtName") =  "asp"     '新闻文件扩展名
	  else
		  INewsAddObj("FileExtName") =  Request.Form("FileExtName")     '新闻文件扩展名
	  end if 
	  INewsAddObj("Path") =  "/" & year(now())&"/"&month(now())&"-"&day(now())             '新闻路径 
	  INewsAddObj("AddDate") =  IAddDate
	  if Request.Form("KeyWords") <> "" then 
		  INewsAddObj("KeyWords") = Replace(Replace(Request.Form("KeyWords"),"""",""),"'","")
	  end if
	  if Request.Form("TxtSource") <> "" then
		  INewsAddObj("TxtSource") = Replace(Replace(Request.Form("TxtSource"),"""",""),"'","")
	  end if
	  if Request.Form("Author") <> "" then
		  INewsAddObj("Author") = Replace(Replace(Request.Form("Author"),"""",""),"'","")
	  end if
	  if Request.Form("Editer") <> "" then
		  INewsAddObj("Editer") = Replace(Replace(Request.Form("Editer"),"""",""),"'","")
	  end if
	  INewsAddObj("ClickNum") =  IClickNum
	  if Request.Form("RecTF") = "1" then
		  INewsAddObj("RecTF") =  "1"
	  else
		  INewsAddObj("RecTF") =  "0"
	  end if
	  if Request.Form("SpecialID") <> "" then
		  INewsAddObj("SpecialID") = Replace(Replace(Request.Form("SpecialID"),"""",""),"'","")
	  end if
	  if Request.Form("PicNewsTF") = "1" then
		  INewsAddObj("PicNewsTF") =  "1"
	  else
		  INewsAddObj("PicNewsTF") =  "0"
	  end if
	  INewsAddObj("PicPath") =  Request.Form("PicPath")
	  if Request.Form("AuditTF") = "1" then
		  INewsAddObj("AuditTF") =  "0"
	  else
		  INewsAddObj("AuditTF") =  "1"
	  end if
	  INewsAddObj("DelTF") =  "0"
	  if Request.Form("BrowPop") <> "" then
		  INewsAddObj("BrowPop") =  Replace(Replace(Request.Form("BrowPop"),"""",""),"'","")
	  end if
	  if Request.Form("ShowReviewTF") = "1" then
		  INewsAddObj("ShowReviewTF") =  "1"
	  else
		  INewsAddObj("ShowReviewTF") =  "0"
	  end if
	  if Request.Form("ReviewTF") = "1" then
		  INewsAddObj("ReviewTF") =  "1"
	  else
		  INewsAddObj("ReviewTF") =  "0"
	  end if
	  if Request.Form("SBSNews") = "1" then
		  INewsAddObj("SBSNews") =  "1"
	  else
		  INewsAddObj("SBSNews") =  "0"
	  end if
	  if Request.Form("MarqueeNews") = "1" then
		  INewsAddObj("MarqueeNews") =  "1"
	  else
		  INewsAddObj("MarqueeNews") =  "0"
	  end if
	  if Request.Form("ProclaimNews") = "1" then
		  INewsAddObj("ProclaimNews") =  "1"
	  else
		  INewsAddObj("ProclaimNews") =  "0"
	  end if
	  if Request.Form("LinkTF") = "1" then
		  INewsAddObj("LinkTF") =  "1"
	  else
		  INewsAddObj("LinkTF") =  "0"
	  end if
	  If Request.Form("TodayNewsTF")<>"" then
		  INewsAddObj("TodayNewsTF") = 1
	  Else
		  INewsAddObj("TodayNewsTF") = 0
	  End If
	  if Request.Form("FilterNews") = "1" then
		  INewsAddObj("FilterNews") =  "1"
	  Else
		  INewsAddObj("FilterNews") =  "0"
	  End If
	  INewsAddObj("NewsTemplet") =  INewsTemplet
	  INewsAddObj("NaviWords") =  Request.Form("NaviWords")
	  If Request.Form("SavePic") = "1" Then
	  	  INewsAddObj("Content") = ReplaceRemoteUrl(Request.Form("Content"),"/" & UpFiles & "/" & BeyondPicDir&"/"&year(Now())&"-"&month(now())&"/"&day(Now()),RsNewsConfigObj("DoMain"),DummyPath_Riker)
	  Else
		  INewsAddObj("Content") = Request.Form("Content")   '新闻内容
	  End If
	  INewsAddObj.Update
	  INewsAddObj.Close
	  Set INewsAddObj = Nothing
	  RsContObj.Close
	  Set RsContObj = Nothing
	  Conn.Execute("Delete from FS_Contribution where ContID = '"&NewsID&"'")
	  
	  Dim CreatePageObj
	  Set CreatePageObj = Conn.Execute("Select * from FS_News where NewsID='"&ConCheckNewsID&"'")
		If Not CreatePageObj.eof then
			RefreshNews CreatePageObj
		End If	
		CreatePageObj.Close
		Set CreatePageObj = Nothing  
	  conn.execute("update FS_members set ConNumNews=ConNumNews+1 where MemName='"&Replace(Replace(Request.Form("Author"),"""",""),"'","")&"'")
	  conn.execute("update FS_members set Point=Point+"&confimsn("NumberContPoint")&" where MemName='"&Replace(Replace(Request.Form("Author"),"""",""),"'","")&"'")
	
	'	<script>
	'		top.GetNavFoldersObject().location='../Menu_Folders.asp?Action=ContentTree&OpenClassIDList=<% = IClassID ';		
	'	</script>
	  Response.Redirect("ContributionList.asp?ClassID=" & IClassID)

  end if
%>