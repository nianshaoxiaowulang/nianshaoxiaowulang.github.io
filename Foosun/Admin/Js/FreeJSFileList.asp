<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
'==============================================================================
'软件名称：风讯网站信息管理系统
'当前版本：Foosun Content Manager System v3.1 
'最新更新：2004.12
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P060400") then Call ReturnError1()
Dim JSID,TempJSObj,FileListObj,FileListSql,FileClassObj,MannerStr,ENameStr,CNameStr
if Request("JsID")<>"" then
	JSID = Clng(Request("JsID"))
else
	Response.Write("<script>alert(""参数传递错误"");history.back();</script>")
	response.end
end if
Set TempJSObj = Conn.Execute("select EName,Type,CName,Manner from FS_FreeJS where ID=" & JSID & "")
If TempJSObj.eof then
	Response.Write("<script>alert(""未查询到相关记录"");history.back();</script>")
	response.end
End If
MannerStr = TempJSObj("Manner")
ENameStr = TempJSObj("EName")
CNameStr =TempJSObj("CName")
'--------删除 FreeJsFile 表的冗余数据 -----
Dim RikerLuObj,RikerNewsObj
Set RikerLuObj = Conn.Execute("Select FileName from FS_FreeJsFile where JSName='" & TempJSObj("EName") & "'")
Do While Not RikerLuObj.eof
	Set RikerNewsObj = Conn.Execute("Select NewsID from FS_News where FileName='" & RikerLuObj("FileName") & "' ")
	If RikerNewsObj.eof then
		Conn.Execute("Delete from FS_FreeJsFile where FileName='" & RikerLuObj("FileName") & "'")
	End If
	RikerNewsObj.Close
	Set RikerNewsObj = Nothing
	RikerLuObj.MoveNext
Loop
RikerLuObj.Close
Set RikerLuObj = Nothing
Set FileListObj=server.createobject(G_FS_RS)
'---------------------------------------------
FileListSql="Select * from FS_FreeJsFile where JSName='" & TempJSObj("EName") & "'"
FileListObj.open FileListSql,Conn,1,1
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>自由JS列表</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script language="JavaScript" src="../../SysJS/ContentMenu.js"></script>
<body onclick="SelectJS(false);" topmargin="2" leftmargin="2" ondragstart="return false;" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35 align="center" alt="删除" onClick="DelFreeJSFile();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="返回" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">返回</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
  <tr> 
    <td valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="41%" height="26" class="ButtonListLeft">
<div align="center">新闻名称</div></td>
          <td width="13%" height="26" class="ButtonList">
<div align="center">所属栏目名称</div></td>
<td width="13%" height="26" class="ButtonList">
<div align="center">所属JS名称</div></td>
          <td width="11%" height="26" class="ButtonList">
<div align="center">状态</div></td>
          <td width="23%" height="26" class="ButtonList">
<div align="center">加入JS时间</div></td>
        </tr>
<% 
do while Not FileListObj.Eof
Dim FileClassName
Set FileClassObj = Conn.Execute("Select ClassCName from FS_NewsClass where ClassID='"&FileListObj("ClassID")&"'")
If FileClassObj.eof then
	FileClassName = "已删除"
Else
	FileClassName = FileClassObj("ClassCName")
End If
Dim TempTypeFlag,RsTempObj,FlagTyype
Set RsTempObj = Conn.Execute("Select HeadNewsTF,PicNewsTF from FS_News where FileName='"&FileListObj("FileName")&"'")
If Not RsTempObj.eof then
   If RsTempObj("HeadNewsTF")="1" then
	  TempTypeFlag = "<img src=""../../Images/Info/TitleNews.gif"" border=""0"">" '标题新闻标志
   elseif RsTempObj("HeadNewsTF")="0" and RsTempObj("PicNewsTF")="0" then
	  TempTypeFlag = "<img src=""../../Images/Info/WordNews.gif"" border=""0"">" '文字新闻标志
   else
	  TempTypeFlag = "<img src=""../../Images/Info/PicNews.gif"" border=""0"">" '图片新闻标志
   end if
Else
	TempTypeFlag = ""
End  If
If FileListObj("DelFlag")=1 then
	FlagTyype = "回收站"
Else
	FlagTyype = "可用"
End If
%>
        <tr> 
          <td> 
            <table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><% = TempTypeFlag %></td>
                <td><span NewsID="<%=FileListObj("ID")%>" JsType="<%=TempJSObj("Type")%>" class="TempletItem" align="center"><%=GotTopic(FileListObj("Title"),40)%></span></td>
              </tr>
            </table>
		  </td>
          <td> 
            <div align="center" class="TempletItem"><%=FileClassName%></div></td>
		  <td> 
            <div align="center" class="TempletItem"><%=CNameStr%></div></td>
          <td> 
            <div align="center" class="TempletItem"><%=FlagTyype%></div></td>
          <td> 
            <div align="center" class="TempletItem"><%=FileListObj("ToJsTime")%></div></td>
        </tr>
<%
	FileClassObj.Close
	FileListObj.MoveNext
loop
FileListObj.close
set FileListObj=nothing
%>
      </table></td>
  </tr>
</table>
</body>
<%
Set Conn = Nothing
%>
<script>
var NewsID = '';
var JsType = '';
var TempENameStr='<%=ENameStr%>';
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelFreeJSFile();",'删除','disabled');
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
	var EventObjInArray=false,SelectAds='',DisabledContentMenuStr='';
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
			if (SelectAds=='') SelectAds=ListObjArray[i].Obj.NewsID;
			else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.NewsID;
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
					if (SelectAds=='') SelectAds=ListObjArray[i].Obj.NewsID;
					else SelectAds=SelectAds+'***'+ListObjArray[i].Obj.NewsID;
				}
			}
		}
	}
	if (SelectAds=='') DisabledContentMenuStr=',删除,';
	else
	{
		if (SelectAds.indexOf('***')==-1) DisabledContentMenuStr='';
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
function SelectJS(MouseRight)
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
function DelFreeJSFile()
{
	var SelectedJSFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.NewsID!=null)
			{
				if (SelectedJSFile=='') SelectedJSFile=ListObjArray[i].Obj.NewsID;
				else  SelectedJSFile=SelectedJSFile+'***'+ListObjArray[i].Obj.NewsID;
			}
		}
	}
	if (SelectedJSFile!='')
		OpenWindow('Frame.asp?FileName=FileModify.asp&Types=Del&PageTitle=删除JS新闻&JsID='+SelectedJSFile,225,105,window);
	else alert('请选择要删除的新闻');
}
function CutOperation()
{
	parent.MoveTF=true;
	if (NewsID!='')
	{
		parent.MoveOrCopySourceClass=TempENameStr;
		parent.MoveOrCopySourceNews=NewsID;
	}
}

function CopyOperation()
{
	parent.MoveTF=false;
	if (NewsID!='')
	{
		parent.MoveOrCopySourceClass=TempENameStr;
		parent.MoveOrCopySourceNews=NewsID;
	}
}

function PasteOperation()
{
	var MoveOrCopyClassPara='MoveTF:'+parent.MoveTF+',SourceClass:'+parent.MoveOrCopySourceClass+',SourceNews:'+parent.MoveOrCopySourceNews+',ObjectClass:'+parent.MoveOrCopyObjectClass+',';
	OpenWindow('JsTip.asp?FileName=MoveOrCopyNews.asp&Title=移动或复制JS新闻&MoveOrCopyClassPara='+MoveOrCopyClassPara,300,105,window);
}
</script>
