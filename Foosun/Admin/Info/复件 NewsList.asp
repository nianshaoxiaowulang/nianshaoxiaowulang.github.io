<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
Dim RsMenuConfigObj,sHaveValueTF
Set RsMenuConfigObj = Conn.execute("Select IsShop From FS_Config")
if RsMenuConfigObj("IsShop") = 1 then
	sHaveValueTF = True
Else
	sHaveValueTF = False
End if
Set RsMenuConfigObj = Nothing
if Not JudgePopedomTF(Session("Name"),"" & Request("ClassID") & "") then Call ReturnError1()
if Not JudgePopedomTF(Session("Name"),"P010500") then Call ReturnError1()
Dim ClassID
ClassID = Request("ClassID")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ݴ���</title>
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
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.RefreshNews();','����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditContent();",'�޸�','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.PreviewNews();','Ԥ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelContent();",'ɾ��','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Audit(true);','���','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.Audit(false);','ȡ�����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CutContent();','����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.CopyContent();','����','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.PasteContent();','ճ��','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ConvertClass();','ת����Ŀ','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.MoveNewsToFile();','�鵵','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.RefreshList();','ˢ��','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'��ҳ��·������\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','·������','');
	IntialListObjArray();
}
function ConvertClass()
{
	var SelectedDownLoad='',SelectedNews='',SelectContentObj=null;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ContentID!=null)
			{
				if (ListObjArray[i].Obj.ContentTypeStr=='4')
				{
					if (SelectedDownLoad=='') SelectedDownLoad=ListObjArray[i].Obj.ContentID;
					else SelectedDownLoad=SelectedDownLoad+'***'+ListObjArray[i].Obj.ContentID;
				}
				else
				{
					if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
					else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
				}
			}
		}
	}
	if ((SelectedDownLoad!='')||(SelectedNews!='')) OpenWindow('Frame.asp?FileName=ConvertClass.asp&PageTitle=ת��&NewsID='+SelectedNews+'&DownLoadID='+SelectedDownLoad,220,105,window);
	else alert('��ѡ��Ҫת�������ݣ�');
location.href=location.href;
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
	if (SelectContent=='') DisabledContentMenuStr=',ת����Ŀ,�޸�,ɾ��,����,����,����,Ԥ��,���,ȡ�����,';
	else
	{
		if (SelectContent.indexOf('***')==-1) DisabledContentMenuStr='';
		else DisabledContentMenuStr=',�޸�,'
	}
	if ((top.MainInfo.SourceNews=='')&&(top.MainInfo.SourceDownLoad=='')) DisabledContentMenuStr=DisabledContentMenuStr+',ճ��,';
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
	OpenWindow('Frame.asp?FileName=ContentSearch.asp&PageTitle=����',400,170,window);
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
function AddContent()
{
	top.GetEkMainObject().location.href='ContentEdit.asp?ClassID='+ClassID;
}
function MoveNewsToFile()
{
	var SelectedContent='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ContentID!=null)
			{
				switch (ListObjArray[i].Obj.ContentTypeStr)
				{
					case '1':
						if (SelectedContent=='') SelectedContent=ListObjArray[i].Obj.ContentID;
						else  SelectedContent=SelectedContent+'***'+ListObjArray[i].Obj.ContentID;
						break;
					case '2':
						if (SelectedContent=='') SelectedContent=ListObjArray[i].Obj.ContentID;
						else  SelectedContent=SelectedContent+'***'+ListObjArray[i].Obj.ContentID;
						break;
					case '3':
						if (SelectedContent=='') SelectedContent=ListObjArray[i].Obj.ContentID;
						else  SelectedContent=SelectedContent+'***'+ListObjArray[i].Obj.ContentID;
						break;
				}
			}
		}
	}
	if (SelectedContent!='') 
		OpenWindow('Frame.asp?FileName=MoveNewsToFile.asp&PageTitle=���Ź鵵&NewsID='+SelectedContent,220,105,window);
	else
		//location='MoveNewsToFile.asp?ClassID='+ClassID;
		OpenWindow('Frame.asp?FileName=MoveNewsToFile.asp&PageTitle=���Ź鵵&ClassID='+ClassID,220,105,window);
	location.href=location.href;
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
				switch (SelectContentObj.ContentTypeStr)
				{
					case '1':
						location='NewsWords.asp?NewsID='+SelectedContent+'&ClassID='+ClassID;
						break;
					case '2':
						location='NewsTitle.asp?NewsID='+SelectedContent+'&ClassID='+ClassID;
						break;
					case '3':
						location='NewsPic.asp?NewsID='+SelectedContent+'&ClassID='+ClassID;
						break;
					case '4':
						location='DownLoad.asp?DownLoadID='+SelectedContent+'&ClassID='+ClassID;
						break;
				}
			}
		}
		else alert('һ��ֻ�ܹ��޸�һ������');
	}
	else alert('��ѡ��Ҫ�޸ĵ�����');
}
function PreviewNews()
{
	var SelectedContent='',SelectedTF=false;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			SelectedTF=true;
			switch (ListObjArray[i].Obj.ContentTypeStr)
			{
				case '1':
					window.open('Read.asp?Table=News&ID='+ListObjArray[i].Obj.ContentID);
					break;
				case '2':
					window.open('Read.asp?Table=News&ID='+ListObjArray[i].Obj.ContentID);
					break;
				case '3':
					window.open('Read.asp?Table=News&ID='+ListObjArray[i].Obj.ContentID);
					break;
				case '4':
					window.open('Read.asp?Table=DownLoad&ID='+ListObjArray[i].Obj.ContentID);
					break;
			}
		}
	}
	if (!SelectedTF) alert('��ѡ��ҪԤ��������!');
}
function ReviewManage()
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
				if (SelectContentObj.ContentTypeStr!='4') location='Review.asp?NewsID='+SelectedContent;
				else location='Review.asp?DownloadID='+SelectedContent;
			}
		}
		else alert('һ��ֻ�ܹ�����һ�����Ż����ص�����');
	}
	else location='Review.asp';
}
function Audit(Flag)
{
	var SelectedNews='',SelectedDownLoad='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if (ListObjArray[i].Obj.ContentTypeStr=='4')
				{
					if (SelectedDownLoad=='') SelectedDownLoad=ListObjArray[i].Obj.ContentID;
					else  SelectedDownLoad=SelectedDownLoad+'***'+ListObjArray[i].Obj.ContentID;
				}
				else
				{
					if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
					else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
				} 
			}
		}
	}
	if ((SelectedNews!='')||(SelectedDownLoad!=''))
	{
		if (Flag) OpenWindow('Frame.asp?FileName=CheckContent.asp&OperateType=Check&PageTitle=�������&NewsID='+SelectedNews+'&DownLoadID='+SelectedDownLoad,220,105,window);
		else  OpenWindow('Frame.asp?FileName=CheckContent.asp&OperateType=UnCheck&PageTitle=�������&NewsID='+SelectedNews+'&DownLoadID='+SelectedDownLoad,220,105,window);
		location.href=location.href;
	}
	else
	{
		alert('��ѡ���������');
	}
}
function RefreshNews()
{
	var SelectedNews='',SelectedDownLoad='',SelectContentObj=null;
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.ContentID!=null)
			{
				switch (ListObjArray[i].Obj.ContentTypeStr)
				{
					case '1':
						if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
						else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
						break;
					case '3':
						if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
						else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
						break;
					case '4':
						if (SelectedDownLoad=='') SelectedDownLoad=ListObjArray[i].Obj.ContentID;
						else  SelectedDownLoad=SelectedDownLoad+'***'+ListObjArray[i].Obj.ContentID;
						break;
				}
			}
		}
	}
	if ((SelectedDownLoad!='')||(SelectedNews!='')) OpenWindow('Frame.asp?FileName=NewsRefresh.asp&PageTitle=����&NewsID='+SelectedNews+'&DownLoadID='+SelectedDownLoad,220,105,window);
	else alert('��ѡ��Ҫ���ɵ�����');
}
function DelContent()
{
	var SelectedNews='',SelectedDownLoad='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if (ListObjArray[i].Obj.ContentTypeStr=='4')
				{
					if (SelectedDownLoad=='') SelectedDownLoad=ListObjArray[i].Obj.ContentID;
					else  SelectedDownLoad=SelectedDownLoad+'***'+ListObjArray[i].Obj.ContentID;
				}
				else
				{
					if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
					else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
				} 
			}
		}
	}
	if ((SelectedNews!='')||(SelectedDownLoad!=''))
	{
		OpenWindow('Frame.asp?FileName=DelContent.asp&Operation=DelContent&PageTitle=ɾ������&NewsID='+SelectedNews+'&DownLoadID='+SelectedDownLoad,220,105,window);
		location.href=location.href;
	}
	else
	{
		alert('��ѡ��ɾ������');
	}
}
function CutContent()
{
	var SelectedNews='',SelectedDownLoad='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if (ListObjArray[i].Obj.ContentTypeStr=='4')
				{
					if (SelectedDownLoad=='') SelectedDownLoad=ListObjArray[i].Obj.ContentID;
					else  SelectedDownLoad=SelectedDownLoad+'***'+ListObjArray[i].Obj.ContentID;
				}
				else
				{
					if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
					else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
				} 
			}
		}
	}
	if ((SelectedNews!='')||(SelectedDownLoad!=''))
	{
		top.MainInfo.SourceClass='';
		top.MainInfo.SourceNews=SelectedNews;
		top.MainInfo.SourceDownLoad=SelectedDownLoad;
		top.MainInfo.ObjectClass='';
		top.MainInfo.MoveTF=true;
		top.MainInfo.OperationType='Content';
	}
	else
	{
		alert('��ѡ��Ҫ���е�����');
	}
}
function CopyContent()
{
	var SelectedNews='',SelectedDownLoad='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if (ListObjArray[i].Obj.ContentTypeStr=='4')
				{
					if (SelectedDownLoad=='') SelectedDownLoad=ListObjArray[i].Obj.ContentID;
					else  SelectedDownLoad=SelectedDownLoad+'***'+ListObjArray[i].Obj.ContentID;
				}
				else
				{
					if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
					else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
				} 
			}
		}
	}
	if ((SelectedNews!='')||(SelectedDownLoad!=''))
	{
		top.MainInfo.SourceClass='';
		top.MainInfo.SourceNews=SelectedNews;
		top.MainInfo.SourceDownLoad=SelectedDownLoad;
		top.MainInfo.ObjectClass='';
		top.MainInfo.MoveTF=false;
		top.MainInfo.OperationType='Content';
	}
	else
	{
		alert('��ѡ��Ҫ���Ƶ�����');
	}
}
function PasteContent()
{
	if (ClassID=='')
	{
		alert('Ŀ����Ŀ������');
		return;
	}
	if ((top.MainInfo.SourceNews=='')&&(top.MainInfo.SourceDownLoad=='')) {alert('��ѡ���ȼ��л��߸�������');return;}
	top.MainInfo.ObjectClass=ClassID;
	var MoveOrCopyClassPara='OperationType:'+top.MainInfo.OperationType+',MoveTF:'+top.MainInfo.MoveTF+',SourceNews:'+top.MainInfo.SourceNews+',SourceDownLoad:'+top.MainInfo.SourceDownLoad+',ObjectClass:'+top.MainInfo.ObjectClass+',';
	OpenWindow('Frame.asp?FileName=MoveOrCopyNewsClass.asp&PageTitle=ճ������&MoveOrCopyClassPara='+MoveOrCopyClassPara,260,100,window);
	location.href=location.href;
}
function AddToSpecial()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if ((ListObjArray[i].Obj.ContentTypeStr=='1')||(ListObjArray[i].Obj.ContentTypeStr=='3'))
				{
					if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
					else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
				}
			}
		}
	}
	if (SelectedNews!='') OpenWindow('Frame.asp?FileName=NewsToSpecial.asp&PageTitle=��ӵ�����ר��&NewsID='+SelectedNews,350,120,window);
	else alert('��ѡ������');
}
function AddToJS()
{
	var SelectedNews='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if ((ListObjArray[i].Obj.ContentID!=null)&&(ListObjArray[i].Obj.ContentTypeStr!=null))
			{
				if ((ListObjArray[i].Obj.ContentTypeStr=='1')||(ListObjArray[i].Obj.ContentTypeStr=='3'))
				{
					if (SelectedNews=='') SelectedNews=ListObjArray[i].Obj.ContentID;
					else  SelectedNews=SelectedNews+'***'+ListObjArray[i].Obj.ContentID;
				}
			}
		}
	}
	if (SelectedNews!='') OpenWindow('Frame.asp?FileName=NewsToJs.asp&Types=PicJs&PageTitle=��ӵ�JS&NewsID='+SelectedNews,350,135,window);
	else alert('��ѡ������');
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
function AuditOneContent(NewsID,DownloadID)
{
	OpenWindow('Frame.asp?FileName=CheckContent.asp&OperateType=Check&PageTitle=�������&NewsID='+NewsID+'&DownloadID='+DownloadID,220,105,window);
	location.href=location.href;
}
function UnAuditOneContent(NewsID,DownloadID)
{
	OpenWindow('Frame.asp?FileName=CheckContent.asp&OperateType=UnCheck&PageTitle=qȡ�����&NewsID='+NewsID+'&DownloadID='+DownloadID,220,105,window);
	location.href=location.href;
}
</script>
<body topmargin="2" leftmargin="2" onclick="ClickNewsOrDownLoad(event);" onselectstart="return false;">
<%
Dim SearchScope,SearchType,SearchContent,SearchBeginTime,SearchEndTime
Dim News_Search_Sql,DownLoad_Search_Sql
Dim SearchDisplayStr,AdvanceSearchDisplayStr,BtnOpenAdvanceSearchDisplayStr
SearchScope = Request("SearchScope")
SearchType = Request("SearchType")
SearchContent = Request("SearchContent")
SearchBeginTime = Request("SearchBeginTime")
SearchEndTime = Request("SearchEndTime")
if SearchContent <> "" then
	SearchDisplayStr = ""
	if SearchBeginTime <> "" and SearchBeginTime <> "" then
		AdvanceSearchDisplayStr = ""
		BtnOpenAdvanceSearchDisplayStr = "none"
	else
		AdvanceSearchDisplayStr = "none"
		BtnOpenAdvanceSearchDisplayStr = ""
	end if
else
	if SearchBeginTime <> "" and SearchBeginTime <> "" then
		SearchDisplayStr = ""
		AdvanceSearchDisplayStr = ""
		BtnOpenAdvanceSearchDisplayStr = "none"
	else
		SearchDisplayStr = "none"
		AdvanceSearchDisplayStr = "none"
		BtnOpenAdvanceSearchDisplayStr = ""
	end if
end if
Select Case SearchScope
	Case "All"
		if SearchType <> "" then
			if SearchType <> "" and SearchContent <> "" then
				News_Search_Sql = " and " & SearchType & " like '%" & SearchContent & "%'"
				DownLoad_Search_Sql = ""
			end if
			if SearchBeginTime <> "" and SearchEndTime <> "" then
				If IsSqlDataBase=0 then
					News_Search_Sql = News_Search_Sql & " and (AddDate between #" & SearchBeginTime & "# and #" & SearchEndTime & "#)"
					DownLoad_Search_Sql = " and (AddTime between #" & SearchBeginTime & "# and #" & SearchEndTime & "#)"
				Else
					News_Search_Sql = News_Search_Sql & " and (AddDate between '" & SearchBeginTime & "' and '" & SearchEndTime & "')"
					DownLoad_Search_Sql = " and (AddTime between '" & SearchBeginTime & "' and '" & SearchEndTime & "')"
				End If
			end if
		end if
	Case "News"
		if SearchType <> "" then
			if SearchType <> "" and SearchContent <> "" then
				News_Search_Sql = " and " & SearchType & " like '%" & SearchContent & "%'"
				DownLoad_Search_Sql = ""
			end if
			if SearchBeginTime <> "" and SearchEndTime <> "" then
				If IsSqlDataBase=0 then
					News_Search_Sql = News_Search_Sql & " and (AddDate between #" & SearchBeginTime & "# and #" & SearchEndTime & "#)"
				Else
					News_Search_Sql = News_Search_Sql & " and (AddDate between '" & SearchBeginTime & "' and '" & SearchEndTime & "')"
				End If
				DownLoad_Search_Sql = ""
			end if
		end if
	Case "DownLoad"
		if SearchType <> "" then
			if SearchType <> "" and SearchContent <> "" then
				News_Search_Sql = ""
				DownLoad_Search_Sql = ""
			end if
			if SearchBeginTime <> "" and SearchEndTime <> "" then
				News_Search_Sql = ""
				If IsSqlDataBase=0 then
					DownLoad_Search_Sql = " and (AddTime between #" & SearchBeginTime & "# and #" & SearchEndTime & "#)"
				Else
					DownLoad_Search_Sql = " and (AddTime between '" & SearchBeginTime & "' and '" & SearchEndTime & "')"
				End If
			end if
		end if
	Case Else
		News_Search_Sql = "" 
		DownLoad_Search_Sql = "" 
end Select

If Request.QueryString("Audit")="IsAuditTF" then 
	News_Search_Sql =News_Search_Sql & " and auditTF=1"
	DownLoad_Search_Sql =News_Search_Sql & " and auditTF=1"
ElseIf Request.QueryString("Audit")="NoAuditTF" then
	News_Search_Sql =News_Search_Sql & " and auditTF=0"
	DownLoad_Search_Sql =News_Search_Sql & " and auditTF=0"	
Else
End If
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=55 align="center" alt="�����Ŀ" onClick="top.GetEkMainObject().location='ClassAdd.asp?ParentID=<% = ClassID %>';" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�����Ŀ</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="�������" onClick="ShowAddMenu();" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�������</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="�����б�" onClick="top.GetEkMainObject().location='DownloadList.asp?ClassID=<% = ClassID %>';" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�����б�</td>
		  <td width=2 class="Gray">|</td>
		  <%If sHaveValueTF = True then%>
		  <td width=55 align="center" alt="��Ʒ�б�" onClick="top.GetEkMainObject().location='ProductList.asp?ClassID=<% = ClassID %>';" onMouseMove="BtnMouseOver(this);//ShowAddMenu();" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">��Ʒ�б�</td>
		  <td width=2 class="Gray">|</td>
		  <%End If%>
          <td width=45 align="center" alt="��ʾ���������" onClick="auditcontent();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">�����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=45  align="center" alt="��ʾδ�������" onClick="noauditcontent();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">δ���</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="���" onClick="Audit(true);" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���</td>
		  <td width=2 class="Gray">|</td>
          <td width=55 align="center" alt="ȡ�����" onClick="Audit(false);" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ȡ�����</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="����ר��" onClick="AddToSpecial();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����ר��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="����JS" onClick="AddToJS();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����JS</td>
		  <td width=2 class="Gray">|</td>
		  <td width=55 align="center" alt="���۹���" onClick="ReviewManage();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">���۹���</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="Ԥ��" onClick="PreviewNews();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">Ԥ��</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="ShowSearchArea();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
 		  <td width=2 class="Gray" style="display:none">|</td>
		  <td width=55 align="center" style="display:none" alt="ȡ������"onClick="CancelSearch();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">ȡ������</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="����" onClick="top.GetEkMainObject().history.back();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">����</td>
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
          <td width="40%" height="26" class="ButtonListLeft"> <div align="center">����</div></td>
          <td nowrap class="ButtonList"> <div align="center">����</div></td>
          <td width="10%" height="26" class="ButtonList"> <div align="center">״̬</div></td>
          <td width="20%" height="26" class="ButtonList"> <div align="center">ʱ��</div></td>
          <td width="15%" class="ButtonList"> <div align="center">����</div></td>
        </tr>
        <%
	Dim NewsSql,RsNewsObj,PicStr,News_For_Var
	NewsSql = "Select * from FS_News where ClassID='" & ClassID & "' and DelTF=0 " & News_Search_Sql & " order by ID desc"
	'Response.Write(NewsSql)
	'Response.End
	Set RsNewsObj = Server.CreateObject(G_FS_RS)
	RsNewsObj.Open NewsSql,Conn,1,1
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
        <tr onmouseover="//this.style.backgroundColor='#F3F3F3';this.style.color='red'" onmouseout="//this.style.backgroundColor='';this.style.color=''"> 
          <td nowrap> <table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><img src="<% = PicStr %>"></td>
                <td nowrap><span ContentTypeStr="<% = ContentTypeStr %>" AuditTF="" class="TempletItem" ContentID="<% = RsNewsObj("NewsID") %>" align="center"> 
                  <% = Left(RsNewsObj("Title"),26) %>
                  </span> </td>
              </tr>
            </table></td>
          <td nowrap> <div align="center">���� </div></td>
          <td nowrap> <div align="center"> 
              <% if RsNewsObj("AuditTF") = "1" then
				  response.Write("<font color=blue>������</font>")
			  else
				  response.Write("<font color=red>δ����</font>")
			  end if
			  %>
            </div></td>
          <td nowrap> <div align="center"> 
              <% = RsNewsObj("AddDate") %>
            </div></td>
			<%if RsNewsObj("AuditTF") = 0 then%>
          <td align="center" nowrap onClick="AuditOneContent('<%=RsNewsObj("NewsID")%>','')" style="cursor:hand;">���</td>
          <%Else%><td align="center" nowrap onClick="UnAuditOneContent('<%=RsNewsObj("NewsID")%>','')" style="cursor:hand;">ȡ�����</td>
		  <%End if%>
        </tr>
        <%
			RsNewsObj.MoveNext
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
          <td height="26" align="right"><% = NewsPageStr %> </td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html> 
<div id="AddContentMenu" class="menu" onMouseOver="MouseOverRightMenu();" onMouseOut="MouseOverRightMenu();" onclick="ClickMenu(this);" style="display:none;"> 
  <table width="100%;" height="80" border="0" cellspacing="0" cellpadding="0" bgcolor="#eeeeee">
  <tr align="center">
    <td height="20" style="cursor:hand;" ExeFunction="top.GetEkMainObject().location='NewsWords.asp?ClassID='+ClassID;">��������</td>
  </tr>
  <tr align="center">
    <td height="20" ExeFunction="top.GetEkMainObject().location='NewsPic.asp?ClassID='+ClassID;">ͼƬ����</td>
  </tr>
  <tr align="center">
    <td height="20" ExeFunction="top.GetEkMainObject().location='NewsTitle.asp?ClassID='+ClassID;">��������</td>
  </tr>
  <tr align="center">
    <td height="20" ExeFunction="top.GetEkMainObject().location='DownLoad.asp?ClassID='+ClassID;">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</td>
  </tr>
 <%if sHaveValueTF = True then%>
   <tr align="center">
    <td height="20" ExeFunction="top.GetEkMainObject().location='../mall/mall_addProducts.asp?ClassID='+ClassID;">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ʒ</td>
  </tr>
 <%End if%>
</table>
</div>
<%
Set RsNewsObj = Nothing
Set Conn = Nothing
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
	NewsPageStr = "λ��:<b>"& News_Page_No &"</b>/<b>"& News_Page_Total &"</b>&nbsp;&nbsp;&nbsp;"
	if News_Page_Total = 1 then
		NewsPageStr = NewsPageStr & "&nbsp;<img src=""../../images/FirstPage.gif"" border=0 alt=��ҳ>&nbsp;" & Chr(13) & Chr(10)
		NewsPageStr = NewsPageStr & "&nbsp;<img src=""../../images/prePage.gif"" border=0 alt=��һҳ>&nbsp;" & Chr(13) & Chr(10)
		NewsPageStr = NewsPageStr & "&nbsp;<img src=""../../images/nextPage.gif"" border=0 alt=��һҳ>&nbsp;" & Chr(13) & Chr(10)
		NewsPageStr = NewsPageStr & "&nbsp;<img src=""../../images/endPage.gif"" border=0 alt=βҳ>&nbsp;" & Chr(13) & Chr(10)
	else
		if cint(News_Page_No) <> 1 and cint(News_Page_No) <> News_Page_Total then
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('1','News_Page_No');"" style=""cursor:hand;""><img src=""../../images/FirstPage.gif"" border=0 alt=��ҳ></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_No - 1 & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/prePage.gif border=0 alt=��һҳ></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_No + 1 & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/nextPage.gif border=0 alt=��һҳ></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_Total & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/endPage.gif border=0 alt=βҳ></span>&nbsp;" & Chr(13) & Chr(10)
		elseif cint(News_Page_No) = 1 then
			NewsPageStr = NewsPageStr & "&nbsp;<img src=../../images/FirstPage.gif border=0 alt=��ҳ>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<img src=../../images/prePage.gif border=0 alt=��һҳ>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_No + 1 & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/nextPage.gif border=0 alt=��һҳ></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_Total & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/endpage.gif border=0 alt=βҳ></span>&nbsp;" & Chr(13) & Chr(10)
		else
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('1','News_Page_No');"" style=""cursor:hand;""><img src=../../images/FirstPage.gif border=0 alt=��ҳ></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<span onclick=""ChangePageNO('" & News_Page_No - 1 & "','News_Page_No');"" style=""cursor:hand;""><img src=../../images/prePage.gif border=0 alt=��һҳ></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<img src=../../images/nextPage.gif border=0 alt=��һҳ></span>&nbsp;" & Chr(13) & Chr(10)
			NewsPageStr = NewsPageStr & "&nbsp;<img src=../../images/endpage.gif border=0 alt=βҳ>&nbsp;" & Chr(13) & Chr(10)
		end if
	end if
End Function
%>