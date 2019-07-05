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
if Not JudgePopedomTF(Session("Name"),"P010600") then Call ReturnError1()
Dim NewsSql,RsNewsObj,RsClassObj,ClassID,ClassCName,RsChildClassObj,AllowContributionTF,DisableStr
ClassID = Request("ClassID")
ClassID = Replace(Replace(Replace(Replace(Replace(ClassID,"'",""),"and",""),"select",""),"or",""),"union","")
if ClassID = "0" or ClassID = "" then
	NewsSql = "Select * from FS_Contribution order by AddTime desc"
	AllowContributionTF = True
Else	
	NewsSql = "Select * from FS_Contribution where ClassID='" & ClassID & "' order by AddTime desc"
	Set RsClassObj = Conn.Execute("Select Contribution from FS_NewsClass where ClassID='" & ClassID & "'")
	if Not RsClassObj.Eof then
		if RsClassObj("Contribution") = 1 then
			AllowContributionTF = True
		else
			AllowContributionTF = False
			DisableStr = "disabled"
		end if
	else
		AllowContributionTF = False
		DisableStr = "disabled"
	end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新闻列表</title>
</head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body topmargin="2" leftmargin="2" onclick="SelectContr();" ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35 align="center" alt="新建" onClick="CreateNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut" <% = DisableStr %>>新建</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="修改" onClick="EditNews();" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut" <% = DisableStr %>>修改</td>
		  <td width=2 class="Gray">|</td>
          <td width=35 align="center" alt="删除" onClick="DelNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut" <% = DisableStr %>>删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35  align="center" alt="审核" onClick="Audit();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut" <% = DisableStr %>>审核</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="37%" height="26" class="ButtonListLeft">
<div align="center">稿件名称</div></td>
          <td width="18%" height="26" class="ButtonList">
<div align="center">添加时间</div></td>
          <td width="18%" height="26" class="ButtonList">
<div align="center">所属栏目</div></td>
          <td width="16%" height="26" class="ButtonList">
<div align="center">作者</div></td>
          <td width="11%" height="26" class="ButtonList">
<div align="center">大小</div></td>
        </tr>
<%
if AllowContributionTF = True then
	Set RsNewsObj = Conn.Execute(NewsSql)
	do while Not RsNewsObj.Eof
	ClassCName=conn.execute("select ClassCName from FS_NewsClass where ClassID='" & RsNewsObj("Classid") & "'")(0)
%>
        <tr> 
          <td height="20">
		  <table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><img src="../../Images/Info/WordNews.gif"></td>
                <td><span class="TempletItem" NewsID="<% = RsNewsObj("ContID") %>" align="center"> 
                  <% = GotTopic(RsNewsObj("Title"),30) %>
                  </span> </td>
              </tr>
          </table>
		  </td>
          <td height="20"><div align="center" class="TempletItem"><% = RsNewsObj("AddTime") %></div></td>
		  <td height="20"><div align="center" class="TempletItem"><% = ClassCName %></div></td>
          <td height="20"><div align="center" class="TempletItem"><% = RsNewsObj("Author") %></div></td>
          <td height="20"><div align="center" class="TempletItem"><% = Len(RsNewsObj("Content")) %>
              b</div></td>
        </tr>
        <%
		RsNewsObj.MoveNext
	loop
%>
<%
else
%>
  <tr> 
    <td colspan="5" height="26"><div align="center">此栏目不允许投稿 </div>
      <div align="center"></div></td>
    </tr>
<%
end if
%>
      </table>
	</td>
  </tr>
</table>
</body>
</html>
<%
Set RsChildClassObj = Nothing
Set RsNewsObj = Nothing
Set RsClassObj = Nothing
Set Conn = Nothing
%>
<script language="javascript"> 
var ClassID = '<% = ClassID %>';
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	IntialListObjArray();
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditNews();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelNews();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Audit();','审核','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','刷新','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
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
	var EventObjInArray=false,SelectContribution='',DisabledContentMenuStr='';
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
			if (SelectContribution=='') SelectContribution=ListObjArray[i].Obj.NewsID;
			else SelectContribution=SelectContribution+'***'+ListObjArray[i].Obj.NewsID;
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
					if (SelectContribution=='') SelectContribution=ListObjArray[i].Obj.NewsID;
					else SelectContribution=SelectContribution+'***'+ListObjArray[i].Obj.NewsID;
				}
			}
		}
	}
	if (SelectContribution=='') DisabledContentMenuStr=',修改,删除,审核,';
	else
	{
		if (SelectContribution.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',修改,'
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function FolderFileObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function IntialListObjArray()
{
	var CurrObj=null,j=1;
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.NewsID!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function SelectContr()
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
function CreateNews()
{
	location='ContributionAdd.asp?ClassID='+ClassID;
}
function EditNews()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
	{
		if (SelectedNews.indexOf('***')==-1) location='ContributionModify.asp?ClassID='+ClassID+'&NewsID='+SelectedNews;
		else alert('一次只能够修改一个新闻');
	}
	else alert('请选择要修改的新闻');
}
function DelNews()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
		OpenWindow('Frame.asp?FileName=ContributionDell.asp&PageTitle=稿件删除&NewsID='+SelectedNews,220,110,window);
	else alert('请选择要删除的投稿');
}
function Audit()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.NewsID;
				else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedNews!='')
	{
		if (SelectedNews.indexOf('***')==-1) location='ContributionCheck.asp?NewsID='+SelectedNews+'&ClassID='+ClassID;
		else alert('一次只能够审核一个新闻');
	}
	else alert('请选择要审核的投稿');
}
function CutOperation()
{
	parent.MoveTF=true;
	if (NewsID!='')
	{  
	     parent.MoveOrCopySourceClass=BigClassID;
		 parent.MoveOrCopySourceNews=NewsID;
	}
}

function CopyOperation()
{
	parent.MoveTF=false;
	if (NewsID!='')
	{
	     parent.MoveOrCopySourceClass=BigClassID;
		 parent.MoveOrCopySourceNews=NewsID;
	}
}

function PasteOperation()
{
	var MoveOrCopyClassPara='MoveTF:'+parent.MoveTF+',SourceClass:'+parent.MoveOrCopySourceClass+',SourceNews:'+parent.MoveOrCopySourceNews+',ObjectClass:'+parent.MoveOrCopyObjectClass+',';
	OpenWindow('ContTip.asp?FileName=MoveOrCopyCont.asp&Titles=稿件移动或复制&MoveOrCopyClassPara='+MoveOrCopyClassPara,310,95,window);
}
</script>
