<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
Dim RsMenuConfigObj,HaveValueTF
Set RsMenuConfigObj = Conn.execute("Select IsShop From FS_Config")
if RsMenuConfigObj("IsShop") = 0 then
	Response.Write("<script>alert(""Sorry!商城未开放,将转到新闻列表!!!\n\n请在栏目修改中设置转向页面!!"&CopyRight&""");location=""NewsList.asp?ClassID="&Request("ClassID")&""";</script>")  
	Response.End
End if
Set RsMenuConfigObj = Nothing
if Not JudgePopedomTF(Session("Name"),"" & Request("ClassID") & "") then Call ReturnError1()
if Not JudgePopedomTF(Session("Name"),"P010500") then Call ReturnError1()
Dim ClassID
ClassID = Request("ClassID")
If Request("Action")="Check" then
	Conn.execute("Update FS_Shop_Products Set Islock=0 where id in("&Request("ID")&")")
	Response.Write("<script>alert(""取消锁定成功！"&CopyRight&""");location=""ProductList.asp?ClassID="&Request("ClassID")&""";</script>")
	Response.end
ElseIf Request("Action")="Lock" then
	Conn.execute("Update FS_Shop_Products Set Islock=1 where id in("&Request("ID")&")")
	Response.Write("<script>alert(""锁定成功！"&CopyRight&""");location=""ProductList.asp?ClassID="&Request("ClassID")&""";</script>")
	Response.end
End if
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
	width:50pt;
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
var ClassID='<% = ClassID %>';
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.RefreshProducts();','生成','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditContent();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.PreviewNews();','预览','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelContent();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Audit(true);','审核','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Audit(false);','锁定','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CutContent();','剪切','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CopyContent();','复制','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.PasteContent();','粘贴','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ConvertClass();','转移栏目','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.MoveNewsToFile();','归档','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ReViewManage();','评论管理','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.RefreshList();','刷新','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
	IntialListObjArray();
}

function RefreshList()
{
	location.href=location.href;
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
	if (SelectContent=='') DisabledContentMenuStr=',修改,删除,生成,预览,审核,锁定,';
	else
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',修改,'
	}
//	if ((top.MainInfo.SourceNews=='')&&(top.MainInfo.SourceProduct=='')) DisabledContentMenuStr=DisabledContentMenuStr+',粘贴,';
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
	OpenWindow('Frame.asp?FileName=ContentSearch.asp&PageTitle=搜索',400,170,window);
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
function auditcontent()
{
	var LocationStr=window.location.href;
	LocationStr = LocationStr.replace(/&Audit=IsAuditTF/g,"").replace(/&Audit=NoAuditTF/g,"");
	LocationStr=LocationStr+"&Audit=IsAuditTF"
	window.location=LocationStr;
}
function noauditcontent()
{
	var LocationStr=window.location.href;
	LocationStr = LocationStr.replace(/&Audit=IsAuditTF/g,"").replace(/&Audit=NoAuditTF/g,"");
	LocationStr=LocationStr+"&Audit=NoAuditTF"
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
function ClickNewsOrProduct()
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
			ListObjArray[ListObjArray.length]=new NewsOrProductObj(CurrObj,j,false);
			j++;
		}
	}
}
function NewsOrProductObj(Obj,Index,Selected)
{
	this.Obj=Obj;
	this.Index=Index;
	this.Selected=Selected;
}
function EditContent()
{
	var SelectedContent='',SelectContentObj=null;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ContentID!=null)
			{
				if (SelectedContent=='') SelectedContent=ListObjArray[i].Obj.ContentID;
				else  SelectedContent=SelectedContent+'***'+ListObjArray[i].Obj.ContentID;
			}
			SelectContentObj=ListObjArray[i].Obj;
		}
	}
	if (SelectedContent!='')
	{
		if (SelectedContent.indexOf('***')==-1)
		{
			if (SelectContentObj.ContentTypeStr!=null)
			{
				location='../mall/mall_AddProducts.asp?ID='+SelectedContent+'&ClassID='+ClassID;
			}
		}
		else alert('一次只能够修改一条新闻');
	}
	else alert('请选择要修改的新闻');
}
function Audit(Action)
{
	var SelectedContent='',SelectContentObj=null;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ContentID!=null)
			{
				if (SelectedContent=='') SelectedContent=ListObjArray[i].Obj.ContentID;
				else  SelectedContent=SelectedContent+','+ListObjArray[i].Obj.ContentID;
			}
			SelectContentObj=ListObjArray[i].Obj;
		}
	}
	if (SelectedContent!='')
	{
		if (SelectContentObj.ContentTypeStr!=null)
		{
			if (Action==true)
			{
				location='ProductList.asp?Action=Check&ID='+SelectedContent+'&ClassID='+ClassID;
			}
			else
				location='ProductList.asp?Action=Lock&ID='+SelectedContent+'&ClassID='+ClassID;
		}
	}
	else alert('请选择要操作的商品');
}
function PreviewNews()
{
	var SelectedContent='',SelectedTF=false;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			SelectedTF=true;
			window.open('Read.asp?Table=Product&ID='+ListObjArray[i].Obj.ContentID);
		}
	}
	if (!SelectedTF) alert('请选择要预览的内容!');
}
function DelContent()
{
	var SelectedProduct='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if (SelectedProduct=='') SelectedProduct=ListObjArray[i].Obj.ContentID;
				else  SelectedProduct=SelectedProduct+'***'+ListObjArray[i].Obj.ContentID;
			}
		}
	}
	if (SelectedProduct!='')
	{
		location='../mall/Mall_DelProduct.asp?action=del&ProductsID='+SelectedProduct+'&ClassID='+ClassID;
		//location.href=location.href;
	}
	else
	{
		alert('请选择删除内容');
	}
}
function ReViewManage()
{
	var SelectedProduct='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if (SelectedProduct=='') SelectedProduct=ListObjArray[i].Obj.ContentID;
				else  SelectedProduct=SelectedProduct+'***'+ListObjArray[i].Obj.ContentID;
			}
		}
	}
	location='../mall/mall_comment.asp?ID='+SelectedProduct;
}
function RefreshProducts()
{
	var SelectedProduct='',SelectContentObj=null;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ContentID!=null)
			{
				if (SelectedProduct=='') SelectedProduct=ListObjArray[i].Obj.ContentID;
				else  SelectedProduct=SelectedProduct+'***'+ListObjArray[i].Obj.ContentID;
			}
		}
	}
	if (SelectedProduct!='') 
	{
		OpenWindow('Frame.asp?FileName=NewsRefresh.asp&PageTitle=生成&ProductID='+SelectedProduct,220,105,window);
	}
	else alert('请选择要生成的内容');
}
function AddToSpec()
{
	var SelectedProduct='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if (SelectedProduct=='') SelectedProduct=ListObjArray[i].Obj.ContentID;
				else  SelectedProduct=SelectedProduct+'***'+ListObjArray[i].Obj.ContentID;
			}
		}
	}
	if (SelectedProduct!='')
	{
//		location='../mall/mall_productsmanage.asp?action=del&ProductsID='+SelectedProduct+'&ClassID='+ClassID;
		OpenWindow('Frame.asp?FileName=ProductToSpecial.asp&PageTitle=添加到专区&ProductID='+SelectedProduct,350,120,window);
		location.href=location.href;
	}
	else
	{
		alert('请选择要加入专区的内容');
	}
}

function ShowAddMenu()
{
	var MenuObj=document.all.AddContentMenu;
	var el=event.srcElement;
	MenuObj.style.display='';
	MenuObj.style.posLeft=el.offsetLeft;
	MenuObj.style.posTop=el.offsetHeight;
	MenuObj.className="menushow";
	MenuObj.setCapture();
}
function MouseOverRightMenu() 
{   
	var el=event.srcElement;
	if (el.tagName!='TD') return;
	if (el.ExeFunction==null) return;
	if (el.style.backgroundColor=="highlight") {el.style.backgroundColor='';el.style.color='black';}
	else {el.style.backgroundColor="highlight";el.style.color='white';}
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
function AuditOneContent(NewsID,ProductID)
{
	OpenWindow('Frame.asp?FileName=CheckContent.asp&OperateType=Check&PageTitle=审核新闻&NewsID='+NewsID+'&ProductID='+ProductID,220,105,window);
	location.href=location.href;
}
function UnAuditOneContent(NewsID,ProductID)
{
	OpenWindow('Frame.asp?FileName=CheckContent.asp&OperateType=UnCheck&PageTitle=q取消审核&NewsID='+NewsID+'&ProductID='+ProductID,220,105,window);
	location.href=location.href;
}
</script>
<body topmargin="2" leftmargin="2" onclick="ClickNewsOrProduct(event);" onselectstart="return false;">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=55 align="center" alt="添加栏目" onClick="top.GetEkMainObject().location='ClassAdd.asp?ParentID=<% = ClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">添加栏目</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="添加内容" onClick="ShowAddMenu();" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">添加内容</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="新闻列表" onClick="top.GetEkMainObject().location='NewsList.asp?ClassID=<% = ClassID %>';" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">新闻列表</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="下载列表" onClick="top.GetEkMainObject().location='DownloadList.asp?ClassID=<% = ClassID %>';" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">下载列表</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="加入专区" onClick="AddToSpec();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">加入专区</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="预览" onClick="Audit(1);" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">审核</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="预览" onClick="Audit(0);" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">锁定</td>
		  <td width=2 class="Gray">|</td>

		  <td width=35 align="center" alt="预览" onClick="PreviewNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">预览</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="搜索" onClick="ReViewManage();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">评论管理</td>
 		  <td width=2 class="Gray" style="display:none">|</td>
		  <td width=55 align="center" style="display:none" alt="取消搜索"onClick="CancelSearch();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">取消搜索</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="后退" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">后退</td>
		  <td>&nbsp; </td>
        </tr>
      </table>
	  </td>
  </tr>
</table>
<table width="100%" height="95%" border="0" cellpadding="0" cellspacing="0">
  <tr>
  	<td colspan="5" height="2"></td>
  </tr>
 <tr>
    <td valign="top">
	<table width="100%" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td width="40%" height="26" class="ButtonListLeft"> <div align="center">商品名称</div></td>
          <td nowrap class="ButtonList"> <div align="center">发布时间</div></td>
          <td width="10%" height="26" class="ButtonList"> <div align="center">状态</div></td>
          <td width="10%" height="26" class="ButtonList"> <div align="center">商品编号</div></td>
          <td width="10%" class="ButtonList"><div align="center">销售量</div></td>
          <td width="15%" class="ButtonList"> <div align="center">库存</div></td>
        </tr>
        <%

	Dim ProductSql,RsProductObj,Product_For_Var
	ProductSql = "Select * from FS_Shop_Products where ClassID='" & ClassID & "' order by Products_AddTime desc"
	Set RsProductObj = Server.CreateObject(G_FS_RS)
	RsProductObj.Open ProductSql,Conn,1,1
	if Not RsProductObj.Eof then
		Dim Product_Page_Size,Product_Page_No,Product_Page_Total,Product_Record_All
		Product_Page_Size = Conn.Execute("Select PerPageNum from FS_Shop_Config")(0)
		If Product_Page_Size="" or IsNull(Product_Page_Size) then Product_Page_Size=15
		Product_Page_No = Request.Querystring("Product_Page_No")
		if Product_Page_No <= 0 or Product_Page_No = "" then Product_Page_No = 1
		RsProductObj.PageSize = Product_Page_Size
		Product_Page_Total = RsProductObj.PageCount
		if (Cint(Product_Page_No) > Product_Page_Total) then Product_Page_No = Product_Page_Total
		RsProductObj.AbsolutePage = Product_Page_No
		Product_Record_All = RsProductObj.RecordCount
		for Product_For_Var = 1 to RsProductObj.PageSize
			if RsProductObj.Eof then Exit For
%>
        <tr> 
          <td height="22" nowrap> <table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><img src="../../Images/Folder/Ffolderclosed.gif" width="21" height="15"></td>
                <td nowrap><span ContentTypeStr="5" AuditTF="" class="TempletItem" ContentID="<% = RsProductObj("ID") %>" align="center"> 
                  <% = Left(RsProductObj("Product_Name"),23) %>
                  </span> </td>
              </tr>
            </table></td>
          <td nowrap> <div align="center" class="TempletItem"><%=RsProductObj("Products_AddTime")%> </div></td>
          <td height="22" nowrap> <div align="center" class="TempletItem"> 
              <% if RsProductObj("IsLock")=0 then
			  response.Write("<font color=blue>开放</font>")
		  else
			  response.Write("<font color=red>锁定</font>")
		  end if
		  %>
            </div></td>
          <td height="22" nowrap> <div align="center" class="TempletItem"> 
              <% = RsProductObj("Products_serial") %>
            </div></td>
          <td height="22" nowrap><div align="center"><%=RsProductObj("SaleNum")%></div></td>
          <%If RsProductObj("Products_Stockpile")<= RsProductObj("MinNum") then%>
          <td><div align="center" class="TempletItem"><font color=red><a href="../Mall/AllData.asp"><img src="../../Images/MinNum.gif" alt="库存少于警戒库存" width="7" height="12" border="0"></a><%=RsProductObj("Products_Stockpile")%></font></a></div></td>
          <%Else%>
          <td><div align="center" class="TempletItem"><%=RsProductObj("Products_Stockpile")%></div></td>
          <%End If%>
        </tr>
        <%
			RsProductObj.MoveNext
		Next
end if
%>
      </table>
</td>
  </tr>
  <tr> 
    <td height="20" class="ButtonListLeft">
<table width="100%" height="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td><div align="right"><% = ProductPageStr %>
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html> 
<div id="AddContentMenu" class="menu" onMouseOver="MouseOverRightMenu();" onMouseOut="MouseOverRightMenu();" onclick="ClickMenu(this);" style="display:none;"> 
  <table width="100%;" height="80" border="0" cellspacing="0" cellpadding="0" bgcolor="#eeeeee">
  <tr align="center">
    <td height="20" style="cursor:hand;" ExeFunction="top.GetEkMainObject().location='NewsWords.asp?ClassID='+ClassID;">文字新闻</td>
  </tr>
  <tr align="center">
    <td height="20" ExeFunction="top.GetEkMainObject().location='NewsPic.asp?ClassID='+ClassID;">图片新闻</td>
  </tr>
  <tr align="center">
    <td height="20" ExeFunction="top.GetEkMainObject().location='NewsTitle.asp?ClassID='+ClassID;">标题新闻</td>
  </tr>
  <tr align="center">
    <td height="20" ExeFunction="top.GetEkMainObject().location='download.asp?ClassID='+ClassID;">下&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;载</td>
  </tr>
  <tr align="center">
    <td height="20" ExeFunction="top.GetEkMainObject().location='../mall/mall_addProducts.asp?ClassID='+ClassID;">商&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;品</td>
  </tr>

</table>
</div>
<%
Set RsProductObj = Nothing
Set Conn = Nothing
Function ProductPageStr()
	ProductPageStr = "位置:<b>"& Product_Page_No &"</b>/<b>"& Product_Page_Total &"</b>&nbsp;&nbsp;&nbsp;"
	if Product_Page_Total = 1 then
		ProductPageStr = ProductPageStr & "&nbsp;<img src=""../../images/FirstPage.gif"" border=0 alt=首页>&nbsp;" & Chr(13) & Chr(10)
		ProductPageStr = ProductPageStr & "&nbsp;<img src=""../../images/prePage.gif"" border=0 alt=上一页>&nbsp;" & Chr(13) & Chr(10)
		ProductPageStr = ProductPageStr & "&nbsp;<img src=""../../images/nextPage.gif"" border=0 alt=下一页>&nbsp;" & Chr(13) & Chr(10)
		ProductPageStr = ProductPageStr & "&nbsp;<img src=""../../images/endPage.gif"" border=0 alt=尾页>&nbsp;" & Chr(13) & Chr(10)
	else
		if cint(Product_Page_No) <> 1 and cint(Product_Page_No) <> Product_Page_Total then
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('1','Products_Page_No');"" style=""cursor:hand;""><img src=""../../images/FirstPage.gif"" border=0 alt=首页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_No - 1 & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/prePage.gif border=0 alt=上一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_No + 1 & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_Total & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/endPage.gif border=0 alt=尾页></span>&nbsp;" & Chr(13) & Chr(10)
		elseif cint(Product_Page_No) = 1 then
			ProductPageStr = ProductPageStr & "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=首页>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<img src=../../images/prePage.gif border=0 alt=上一页>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_No + 1 & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_Total & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/endpage.gif border=0 alt=尾页></span>&nbsp;" & Chr(13) & Chr(10)
		else
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('1','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/FirstPage.gif border=0 alt=首页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<span onclick=""ChangePageNO('" & Product_Page_No - 1 & "','Product_Page_No');"" style=""cursor:hand;""><img src=../../images/prePage.gif border=0 alt=上一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<img src=../../images/nextPage.gif border=0 alt=下一页></span>&nbsp;" & Chr(13) & Chr(10)
			ProductPageStr = ProductPageStr & "&nbsp;<img src=../../images/endpage.gif border=0 alt=尾页>&nbsp;" & Chr(13) & Chr(10)
		end if
	end if
End Function
%>