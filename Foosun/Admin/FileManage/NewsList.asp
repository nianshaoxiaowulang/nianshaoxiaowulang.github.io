<% Option Explicit %>
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
Dim DBC,Conn,RecordConn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
DBC.ConnStr = "DBQ=" + server.mappath(RecordDataBaseConnectStr) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set RecordConn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070606") then Call ReturnError1()
Dim SearchScope,SearchType,SearchContent,SearchBeginTime,SearchEndTime
Dim News_Search_Sql
Dim Action
Action = Request("Action")
if Action = "DelNews" then
	if Not JudgePopedomTF(Session("Name"),"P070609") then Call ReturnError1()
	Dim DelNewsID
	DelNewsID = Request("DelNewsID")
	DelNewsID = Replace(Replace(Replace(Replace(Replace(DelNewsID,"'",""),"and",""),"select",""),"or",""),"union","")
	if DelNewsID <> "" then
		DelNewsID = Replace(DelNewsID,"***","','")
		RecordConn.Execute("Delete from FS_News where NewsID in ('" & DelNewsID & "')")
	end if
end if 
SearchScope = Request("SearchScope")
SearchType = Request("SearchType")
SearchContent = Request("SearchContent")
SearchBeginTime = Request("SearchBeginTime")
SearchEndTime = Request("SearchEndTime")
Select Case SearchScope
	Case "All"
		if SearchType <> "" then
			if SearchType <> "" and SearchContent <> "" then
				News_Search_Sql = " and " & SearchType & " like '%" & SearchContent & "%'"
			end if
			if SearchBeginTime <> "" and SearchEndTime <> "" then
				If IsSqlDataBase=0 then
					News_Search_Sql = News_Search_Sql & " and (AddDate between #" & SearchBeginTime & "# and #" & SearchEndTime & "#)"
				Else
					News_Search_Sql = News_Search_Sql & " and (AddDate between '" & SearchBeginTime & "' and '" & SearchEndTime & "')"
				End If
			end if
		end if
	Case "News"
		if SearchType <> "" then
			if SearchType <> "" and SearchContent <> "" then
				News_Search_Sql = " and " & SearchType & " like '%" & SearchContent & "%'"
			end if
			if SearchBeginTime <> "" and SearchEndTime <> "" then
				If IsSqlDataBase=0 then 
					News_Search_Sql = News_Search_Sql & " and (AddDate between #" & SearchBeginTime & "# and #" & SearchEndTime & "#)"
				Else
					News_Search_Sql = News_Search_Sql & " and (AddDate between '" & SearchBeginTime & "' and '" & SearchEndTime & "')"
				End If
			end if
		end if
	Case Else
		News_Search_Sql = "" 
end Select
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>内容窗口</title>
<style type="text/css">
<!--
.SearchBtnStyle {
	border: 1px solid #000000;
}
.menu {
	position:absolute;
	background: menu;
	border-top: 1px solid buttonhighlight;
	border-left: 1px solid buttonhighlight;
	border-bottom: 2px solid buttonshadow;
	border-right: 2px solid buttonshadow;
	padding: 2px;
	font: menu;
	cursor:default;
	font-size:9pt;
	width:90pt;
	visibility: hidden;
	z-index: 2;
	overflow: visible;
}
.menushow {
	position:absolute;
	visibility:visible;
	background:#EFEFEF;
	border-top: 1px solid #000000;
	border-left: 1px solid #000000;
	border-bottom: 1px solid #000000;
	border-right: 1px solid #000000;
	padding:0px;
	font: 9pt "menu";
	cursor:default;
	width:38pt;
}
-->
</style>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../../SysJS/PublicJS.js"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<script language="JavaScript">
var ContentMenuArray=new Array();
var ListObjArray=new Array();
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelContent();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.RefreshNews();','生成','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.PreviewNews();','预览','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','刷新','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
	IntialListObjArray();
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	ChangeContentMenuStatus();
}
function ChangeContentMenuStatus()
{
	var EventObjInArray=false,SelectContent='',DisabledContentMenuStr='';
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			if (ListObjArray[i].Selected==true) EventObjInArray=true;
			break;
		}
	}
	for (var i=0;i<ListObjArray.length;i++)
	{
		if (event.srcElement==ListObjArray[i].Obj)
		{
			ListObjArray[i].Obj.className='TempletSelectItem';
			ListObjArray[i].Selected=true;
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.ContentID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.ContentID;
		}
		else
		{
			if (!EventObjInArray)
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
			else
			{
				if (ListObjArray[i].Selected==true)
				{
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.ContentID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.ContentID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',修改,删除,生成,预览,';
	else
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',修改,预览,'
	}
	if ((top.MainInfo.SourceNews=='')&&(top.MainInfo.SourceDownLoad=='')) DisabledContentMenuStr=DisabledContentMenuStr+',粘贴,';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function ChangePageNO(NO,SearchStr)
{
	var LocationStr=window.location.href;
	var SearchLocation=LocationStr.lastIndexOf(SearchStr);
	if (SearchLocation!=-1)
	{
		var TempSearchLocation=LocationStr.indexOf('&',SearchLocation);
		if (TempSearchLocation!=-1)
		{
			LocationStr=LocationStr.slice(0,SearchLocation)+SearchStr+'='+NO+window.location.href.slice(TempSearchLocation);
		}
		else LocationStr=LocationStr.slice(0,SearchLocation)+SearchStr+'='+NO;
	}
	else
	{
		if (LocationStr.lastIndexOf('?')!=-1) LocationStr=LocationStr+'&'+SearchStr+'='+NO;
		else  LocationStr=LocationStr+'?'+SearchStr+'='+NO;
	}
	window.location=LocationStr;
}
function ShowSearchArea()
{
	OpenWindow('Frame.asp?FileName=ContentSearch.asp&PageTitle=搜索',400,200,window);
}
function AddLocationStr(LocationStr,Value,SearchStr)
{
	var SearchLocation=LocationStr.lastIndexOf(SearchStr);
	if (SearchLocation!=-1)
	{
		var TempSearchLocation=LocationStr.indexOf('&',SearchLocation);
		if (TempSearchLocation!=-1)
		{
			var TempLocationStr=LocationStr.slice(TempSearchLocation)
			LocationStr=LocationStr.slice(0,SearchLocation)+SearchStr+'='+Value+TempLocationStr;
		}
		else LocationStr=LocationStr.slice(0,SearchLocation)+SearchStr+'='+Value;
	}
	else
	{
		if (LocationStr.lastIndexOf('?')!=-1) LocationStr=LocationStr+'&'+SearchStr+'='+Value;
		else  LocationStr=LocationStr+'?'+SearchStr+'='+Value;
	}
	return LocationStr;
}
function SearchSubmit(FormObj)
{
	var LocationStr=window.location.href;
	LocationStr=AddLocationStr(LocationStr,FormObj.SearchScope.value,'SearchScope');
	LocationStr=AddLocationStr(LocationStr,FormObj.SearchType.value,'SearchType');
	LocationStr=AddLocationStr(LocationStr,FormObj.SearchContent.value,'SearchContent');
	LocationStr=AddLocationStr(LocationStr,FormObj.SearchBeginTime.value,'SearchBeginTime');
	LocationStr=AddLocationStr(LocationStr,FormObj.SearchEndTime.value,'SearchEndTime');
	window.location=LocationStr;
}
function CancelSearch()
{
	var LocationStr=window.location.href;
	LocationStr=AddLocationStr(LocationStr,'','SearchScope');
	LocationStr=AddLocationStr(LocationStr,'','SearchType');
	LocationStr=AddLocationStr(LocationStr,'','SearchContent');
	LocationStr=AddLocationStr(LocationStr,'','SearchBeginTime');
	LocationStr=AddLocationStr(LocationStr,'','SearchEndTime');
	window.location=LocationStr;
}
function ClickNewsOrDownLoad()
{
	var el=event.srcElement;
	var i=0;
	if ((event.ctrlKey==true)||(event.shiftKey==true))
	{
		if (event.ctrlKey==true)
		{
			for (i=0;i<ListObjArray.length;i++)
			{
				if (el==ListObjArray[i].Obj)
				{
					if (ListObjArray[i].Selected==false)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					else
					{
						ListObjArray[i].Obj.className='TempletItem';
						ListObjArray[i].Selected=false;
					}
				}
			}
		}
		if (event.shiftKey==true)
		{
			var MaxIndex=0,ObjInArray=false,EndIndex=0,ElIndex=-1;
			for (i=0;i<ListObjArray.length;i++)
			{
				if (ListObjArray[i].Selected==true)
				{
					if (ListObjArray[i].Index>=MaxIndex) MaxIndex=ListObjArray[i].Index;
				}
				if (el==ListObjArray[i].Obj)
				{
					ObjInArray=true;
					ElIndex=i;
					EndIndex=ListObjArray[i].Index;
				}
			}
			if (ElIndex>MaxIndex)
			{
				if (MaxIndex>0)
					for (i=MaxIndex-1;i<EndIndex;i++)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
				else
				{
					ListObjArray[ElIndex].Obj.className='TempletSelectItem';
					ListObjArray[ElIndex].Selected=true;
				}
			}
			else
			{
				if (ObjInArray)
				{
					for (i=EndIndex;i<MaxIndex-1;i++)
					{
						ListObjArray[i].Obj.className='TempletSelectItem';
						ListObjArray[i].Selected=true;
					}
					if (ElIndex>=0)
					{
						ListObjArray[ElIndex].Obj.className='TempletSelectItem';
						ListObjArray[ElIndex].Selected=true;
					}
				}
			}
		}
	}
	else
	{
		for (i=0;i<ListObjArray.length;i++)
		{
			if (el==ListObjArray[i].Obj)
			{
				ListObjArray[i].Obj.className='TempletSelectItem';
				ListObjArray[i].Selected=true;
			}
			else
			{
				ListObjArray[i].Obj.className='TempletItem';
				ListObjArray[i].Selected=false;
			}
		}
	}
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.ContentID!=null)
		{
			ListObjArray[ListObjArray.length]=new NewsOrDownLoadObj(CurrObj,j,false);
			j++;
		}
	}
}
function NewsOrDownLoadObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function PreviewNews()
{
	var SelectedContent='',SelectedTF=false;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			SelectedTF=true;
			window.open('Read.asp?Table=FS_News&ID='+ListObjArray[i].Obj.ContentID);
		}
	}
	if (!SelectedTF) alert('请选择要预览的内容!');
}
function RefreshNews()
{
	location='RefreshFile.asp';
}
function DelContent()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
			}
		}
	}
	if (SelectedNews!='') {if (confirm('确定要删除吗？')) location='?Action=DelNews&DelNewsID='+SelectedNews;}
	else alert('请选择删除内容');
}
function ClickMenu(MenuObj)
{
	var CurrObj=null;
	var IMGObj=document.body.getElementsByTagName('IMG');
	for (var i=0;i<IMGObj.length;i++)
	{
		CurrObj=IMGObj(i);
		if (CurrObj.className=='BtnMouseOver') CurrObj.className='';
	}
	var el=event.srcElement;
	MenuObj.releaseCapture();
	MenuObj.className="menu";
	for (var i=0;i<MenuObj.children.length;i++)
	{
		var CurrObj=MenuObj.children(i);
		for (var j=0;j<CurrObj.children.length;j++)
		{
			if (CurrObj.children(j).className=='MenuShow') {CurrObj.children(j).className='Menu';}	
		}
	}
	if (el.ExeFunction!=null) eval(el.ExeFunction);
}
</script>
<script src="../SysJS/ContentMenu.js" language="JavaScript"></script>
<body topmargin="2" leftmargin="2" scroll=auto onclick="ClickNewsOrDownLoad(event);" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width="35" align="center" alt="删除内容" onClick="DelContent();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width="35" align="center" alt="刷新新闻" onClick="RefreshNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">生成</td>
		  <td width=2 class="Gray">|</td>
		  <td width="35" align="center" alt="预览" onClick="PreviewNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">预览</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="90%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr>
    <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="40%" height="26" class="ButtonListLeft"> <div align="center">名称</div></td>
          <td nowrap class="ButtonList"> <div align="center">栏目</div></td>
          <td width="20%" height="26" class="ButtonList"> <div align="center">时间</div></td>
          <td width="20%" nowrap class="ButtonList">
<div align="center">归档时间</div></td>
        </tr>
        <%
	Dim NewsSql,RsNewsObj,PicStr,News_For_Var
	NewsSql = "Select * from FS_News order by ID desc"
	Set RsNewsObj = Server.CreateObject(G_FS_RS)
	RsNewsObj.Open NewsSql,RecordConn,1,1
	if Not RsNewsObj.Eof then
		Dim News_Page_Size,News_Page_No,News_Page_Total,News_Record_All,ContentTypeStr
		News_Page_Size = 20
		News_Page_No = Request.Querystring("News_Page_No")
		if News_Page_No <= 0 or News_Page_No = "" then News_Page_No = 1
		RsNewsObj.PageSize = News_Page_Size
		News_Page_Total = RsNewsObj.PageCount
		if (Cint(News_Page_No) > News_Page_Total) then News_Page_No = News_Page_Total
		RsNewsObj.AbsolutePage = News_Page_No
		News_Record_All = RsNewsObj.RecordCount
		for News_For_Var = 1 to RsNewsObj.PageSize
			if RsNewsObj.Eof then Exit For
			if RsNewsObj("HeadNewsTF")<>"1" and RsNewsObj("PicNewsTF")<>"1" then
			   PicStr = "../../Images/Info/WordNews.gif"
			   ContentTypeStr = "1"
			elseif RsNewsObj("HeadNewsTF")="1" then
			   PicStr = "../../Images/Info/TitleNews.gif"
			   ContentTypeStr = "2"
			else
			   PicStr = "../../Images/Info/PicNews.gif"
			   ContentTypeStr = "3"
			end if
			if RsNewsObj("FileExtName") = "asp" then
			   PicStr = "../../Images/Info/asp.gif"
			end If
%>
        <tr> 
          <td nowrap> <table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><img src="<% = PicStr %>"></td>
                <td nowrap><span ContentTypeStr="<% = ContentTypeStr %>" AuditTF="" class="TempletItem" ContentID="<% = RsNewsObj("NewsID") %>" align="center"> 
                  <% = Left(RsNewsObj("Title"),26) %>
                  </span> </td>
              </tr>
            </table></td>
          <td nowrap> <div align="center">
		  <%
		  Dim RsClassObj
		  Set RsClassObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='" & RsNewsObj("ClassID") & "'")
		  if Not RsClassObj.Eof then
		  	Response.Write(RsClassObj("ClassCName"))	
		  else
		  	Response.Write("栏目不存在")	
		  end if
		  Set RsClassObj = Nothing
		  %>
		  </div></td>
          <td nowrap> <div align="center"> 
              <% = RsNewsObj("AddDate") %>
            </div></td>
          <td nowrap><div align="center"><% = RsNewsObj("FileTime") %></div></td>
        </tr>
        <%
			RsNewsObj.MoveNext
		Next
end if
%>
      </table>
</td>
  </tr>
</table>
<table width="100%" height="26" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="26" class="buttonlistleft">翻页 <% = NewsPageStr %> </td>
        </tr>
</table>
</body>
</html> 
<%
Set RsNewsObj = Nothing
Set Conn = Nothing
Set RecordConn = Nothing
Function GetNewsOptionValue(Flag,FieldName)
	Dim GetLocation,CheckLength
	Dim CheckArray ,i
	GetLocation = 0
	CheckArray = Array("type","contribution","audit","deleted","link","rec","sbs","marquee","bulletin","filter","focus","classical","today","showreview","reviewtf")
	for i = LBound(CheckArray) to UBound(CheckArray)
		if CheckArray(i) = FieldName then
			GetLocation = i
		end if
	Next
	CheckLength = UBound(CheckArray) + 1 - GetLocation
	if Not IsNull(Flag) then
		if GetLocation > 0 then
			if Len(Flag) < CheckLength then
				GetNewsOptionValue = ""
			else
				GetNewsOptionValue = Mid(Flag,1,GetLocation)
			end if
		else
			GetNewsOptionValue=""
		end if
	else
		GetNewsOptionValue=""
	end if
End Function
Function NewsPageStr()
	NewsPageStr = "NO.<b>"& News_Page_No &"</b>,&nbsp;&nbsp;"
	NewsPageStr = NewsPageStr & "Totel:<b>"& News_Page_Total &"</b>,&nbsp;RecordCounts:<b>" & News_Record_All &"</b>&nbsp;&nbsp;&nbsp;"
	if News_Page_Total = 1 then
		NewsPageStr = NewsPageStr & "&nbsp;<img src=""../../images/FirstPage.gif"" border=0 alt=首页>&nbsp;" & Chr(13) & Chr(10)
		NewsPageStr = NewsPageStr & "&nbsp;<img src=""../../images/prePage.gif"" border=0 alt=上一页>&nbsp;" & Chr(13) & Chr(10)
		NewsPageStr = NewsPageStr & "&nbsp;<img src=""../../images/nextPage.gif"" border=0 alt=下一页>&nbsp;" & Chr(13) & Chr(10)
		NewsPageStr = NewsPageStr & "&nbsp;<img src=""../../images/endPage.gif"" border=0 alt=尾页>&nbsp;" & Chr(13) & Chr(10)
	else
		if cint(News_Page_No) <> 1 and cint(News_Page_No) <> News_Page_Total then
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('1','News_Page_No');"" style=""cursor:hand;""><img src=""../../images/FirstPage.gif"" border=0 alt=首页></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_No - 1 & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/prePage.gif border=0 alt=上一页></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_No + 1 & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_Total & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/endPage.gif border=0 alt=尾页></span>&nbsp;" & Chr(13) & Chr(10)
		elseif cint(News_Page_No) = 1 then
			NewsPageStr = NewsPageStr & "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=首页>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<img src=../../images/prePage.gif border=0 alt=上一页>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_No + 1 & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_Total & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/endpage.gif border=0 alt=尾页></span>&nbsp;" & Chr(13) & Chr(10)
		else
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('1','News_Page_No');"" style=""cursor:hand;""><img src=../../images/FirstPage.gif border=0 alt=首页></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_No - 1 & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/prePage.gif border=0 alt=上一页></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<img src=../../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<img src=../../images/endpage.gif border=0 alt=尾页>&nbsp;" & Chr(13) & Chr(10)
		end if
	end if
End Function
%>