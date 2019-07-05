<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Function.asp" -->
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
<!--#include file="../../Inc/Session.asp" -->
<!--#include file="../../Inc/CheckPopedom.asp" -->
<%
Dim RsMenuConfigObj,HaveValueTF,ShowNode
Set RsMenuConfigObj = Conn.execute("Select IsShop From FS_Config")
if RsMenuConfigObj("IsShop") = 1 then
	HaveValueTF = True
	ShowNode="node"
Else
	HaveValueTF = False
	ShowNode="LastNode"
End if
Set RsMenuConfigObj = Nothing
Dim Action
Action = Request("Action")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>目录</title>
</head>
<script language="JavaScript">
//公共JS
var ContentMenuArray=new Array();
var SelectedClassObj=null;
var DocumentReadyTF=false;
<% if Action = "ContentTree" then %>
	var OpenClassIDList='<% = Request("OpenClassIDList") %>';
	var ClassListVarObjectInstance=null;
<% end if %>
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	<% if Action = "ContentTree" then %>
		ClassListVarObjectInstance=new ClassListVarObject('0','0')
		OpenAllParentClassList();
		InitialClassListContentMenu();
	<% end if %>
	<% if Action = "Special" then %>
		InitialSpecialListContentMenu();
	<% end if %>
	<% if Action = "UpLoad" then %>
		InitialUpLoadContentMenu();
	<% end if %>
	<% if Action = "JSManage" then %>
		InitialJSContentMenu();
	<% end if %>
	DocumentReadyTF=true;
}
function ContentMenuShowEvent()
{
	<% if Action = "ContentTree" then %>
	ClassListCilckContentMenu();
	<% end if %>
	<% if Action = "Special" then %>
	SpecialListCilckContentMenu();
	<% end if %>
	<% if Action = "UpLoad" then %>
	UpLoadCilckContentMenu();
	<% end if %>
	<% if Action = "JSManage" then %>
	JSCilckContentMenu();
	<% end if %>
}
function ClickClassImg(ClickObj,ClassID)
{
	var ImgSrc=ClickObj.src,OpenTF;
	var FolderObj=ClickObj.parentElement.children(ClickObj.parentElement.children.length-1);
	if (ImgSrc.indexOf('Close.gif')!=-1) {ClickObj.src='../Images/Folder/Open.gif';OpenTF=true}
	if (ImgSrc.indexOf('EndClose.gif')!=-1) {ClickObj.src='../Images/Folder/EndOpen.gif';OpenTF=true};
	if (ImgSrc.indexOf('Open.gif')!=-1) {ClickObj.src='../Images/Folder/Close.gif';OpenTF=false;}
	if (ImgSrc.indexOf('EndOpen.gif')!=-1) {ClickObj.src='../Images/Folder/EndClose.gif';OpenTF=false;}
	if (OpenTF) 
	{
		if (FolderObj.src.indexOf('folderclosed.gif')!=-1) FolderObj.src='../Images/Folder/folderopen.gif';
		ShowChildClass(ClassID);
	}
	else
	{
		if (FolderObj.src.indexOf('folderopen.gif')!=-1) FolderObj.src='../Images/Folder/folderclosed.gif';
		HideChildClass(ClassID);
	}
}
function ChangeImg(Obj,OpenTF)
{
	var CurrObj=null,ImgSrc='';
	for (var i=0;i<Obj.all.length;i++)
	{
		CurrObj=Obj.all(i);
		if (CurrObj.tagName.toLowerCase()=='img')
		{
			ImgSrc=CurrObj.src;
			if (OpenTF==true)
			{
				if (ImgSrc.indexOf('Close.gif')!=-1) CurrObj.src='../Images/Folder/Open.gif';
				if (ImgSrc.indexOf('EndClose.gif')!=-1) CurrObj.src='../Images/Folder/EndOpen.gif';
				if (ImgSrc.indexOf('Open.gif')!=-1) return;
				if (ImgSrc.indexOf('EndOpen.gif')!=-1) return;
				if (ImgSrc.indexOf('folderopen.gif')!=-1) return;
				if (ImgSrc.indexOf('folderclosed.gif')!=-1) CurrObj.src='../Images/Folder/folderopen.gif';
			}
			else
			{
				if (ImgSrc.indexOf('Close.gif')!=-1) return;
				if (ImgSrc.indexOf('EndClose.gif')!=-1) return;
				if (ImgSrc.indexOf('Open.gif')!=-1) CurrObj.src='../Images/Folder/Close.gif';
				if (ImgSrc.indexOf('EndOpen.gif')!=-1) CurrObj.src='../Images/Folder/EndClose.gif';
				if (ImgSrc.indexOf('folderopen.gif')!=-1) CurrObj.src='../Images/Folder/folderclosed.gif';
				if (ImgSrc.indexOf('folderclosed.gif')!=-1) return;
			}
		}
	}
}
function HideChildClass(ID)
{
	var CurrObj=null;
	var TRObj=document.body.getElementsByTagName('TR');
	for (var i=0;i<TRObj.length;i++)
	{
		CurrObj=TRObj(i);
		if (CurrObj.AllParentID!=null)
		{
			if (CurrObj.AllParentID.indexOf(ID)!=-1) CurrObj.style.display='none';
		}
	}
}

function ShowChildClass(ID)
{
	var CurrObj=null;
	var TRObj=document.body.getElementsByTagName('TR');
	for (var i=0;i<TRObj.length;i++)
	{
		CurrObj=TRObj(i);
		if (CurrObj.ParentID==ID)
		{
			if (CurrObj.tagName.toLowerCase()=='tr')
			{
				CurrObj.style.display='';
				ChangeImg(CurrObj,false);
			}
		}
	}
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ClickBtn(Obj,TypeStr)
{
	if (Obj!=SelectedClassObj)
	{
		Obj.className='TempletSelectItem';
		if (SelectedClassObj!=null) SelectedClassObj.className='TempletItem';
		SelectedClassObj=Obj;
	}
	top.GetEkMainObject().location=GetLocation(TypeStr,Obj);
}
//公共JS
//内容JS
function ClickFolderTxt(ClickObj,ClassID,ParentID,RedirectList)
{
	var NodeObj=ClickObj.parentElement.parentElement.children(0).children(ClickObj.parentElement.parentElement.children(0).children.length-2);
	var FolderObj=ClickObj.parentElement.parentElement.children(0).children(ClickObj.parentElement.parentElement.children(0).children.length-1);
	if (SelectedClassObj!=null) SelectedClassObj.className='TempletItem';
	ClickObj.className='TempletSelectItem';
	SelectedClassObj=ClickObj;
	if (ParentID=='110110')
	{
		if (SelectedClassObj.ClassID!=ClassID) top.GetEkMainObject().location='Info/ShowOutClass.asp';
	}
	else
	{
		if (SelectedClassObj.ClassID!=ClassID)
		{
			switch (RedirectList) 
			{
				case '1' :
					top.GetEkMainObject().location='Info/NewsList.asp?ClassID='+ClassID;break;
				case '2' :
					top.GetEkMainObject().location='Info/DownloadList.asp?ClassID='+ClassID;break;
				case '3' :
					top.GetEkMainObject().location='Info/ProductList.asp?ClassID='+ClassID;break;
				default :
					top.GetEkMainObject().location='Info/NewsList.asp?ClassID='+ClassID;
			} 

		}
		
		//if (SelectedClassObj.ClassID!=ClassID) top.GetEkMainObject().location='Info/NewsList.asp?ClassID='+ClassID;
	}
	ClassListVarObjectInstance.ClassID=ClassID;
	ClassListVarObjectInstance.ParentID=ParentID;
}
function InitialClassListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddClass();",'添加','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditClass();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelClass();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.ViewClass();",'预览','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ClassCut();','剪切','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ClassPaste();','粘贴','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.MergeClass();','合并','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ClassInit();','栏目初始化','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.ContributionMan();','投稿管理','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetNavFoldersObject().location.reload();','刷新','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
}
function InitialSpecialListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddSpecial();",'添加','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditSpecial();",'修改','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelSpecial();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetNavFoldersObject().location.reload();','刷新','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.SpecialInit();','初始化','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.MergeSpecial();','合并到','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
}
function InitialUpLoadContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddFolder();",'新建栏目','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelFolder();",'删除栏目','disabled');
	//ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditFolder();",'重命名','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetNavFoldersObject().location.reload();','刷新','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
}
function InitialJSContentMenu()
{
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddFreeJSStyle();",'新建自由JS','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditFreeJSStyle();",'修改自由JS','disabled');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.DelFreeJSStyle();",'删除自由JS','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetNavFoldersObject().location.reload();','刷新','');
//	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
}
function AddFreeJSStyle()
{
	top.GetEkMainObject().location.href='JS/JsAdd.asp';
}

function EditFreeJSStyle()
{
	if (SelectedClassObj!=null) top.GetEkMainObject().location.href='JS/JsModify.asp?JsID='+SelectedClassObj.JsID;
	else alert('请选择修改的JS');
}
function DelFreeJSStyle()
{
	if (SelectedClassObj!=null) OpenWindow('JS/Frame.asp?FileName=JsDell.asp&PageTitle=删除自由JS&JsID='+SelectedClassObj.JsID,220,95,window);
	else alert('请选择删除的JS');
}
function AddFolder()
{
	var TempPath='';
	if (SelectedClassObj==null) TempPath=document.all.RootDir.RootDir;
	else TempPath=SelectedClassObj.Path;
	OpenWindow('../FunPages/Frame.asp?PageTitle=创建目录&FileName=AddFolder.asp&Path='+TempPath,200,90,window);
}
function EditFolder()
{

}
function DelFolder()
{
	var TempPath=SelectedClassObj.Path;
	if ((TempPath=='')||(TempPath=='/')) {alert('路径错误，请重试');return;}
	if (TempPath.lastIndexOf('/')!=0)
	{
		var DelFolder=TempPath.slice(TempPath.lastIndexOf('/')+1);
		var Path=TempPath.slice(0,TempPath.lastIndexOf('/'));
	}
	else
	{
		var DelFolder=TempPath.slice(TempPath.lastIndexOf('/')+1);
		var Path='\\';
	}
	if ((DelFolder!='')&&(Path!=''))
	OpenWindow('../FunPages/Frame.asp?PageTitle=删除目录&FileName=DelFolderAndFile.asp&Path='+Path+'&DelFolder='+DelFolder,200,90,window);
	else alert('请选择要删除的栏目');
}
function AddSpecial()
{
	top.GetEkMainObject().location.href='Info/SpecialAdd.asp';
}
function EditSpecial()
{
	if (SelectedClassObj!=null) top.GetEkMainObject().location.href='Info/SpecialModify.asp?SpecialID='+SelectedClassObj.SpecialID;
	else alert('请选择修改的专题');
}
function DelSpecial()
{
	if (SelectedClassObj!=null) OpenWindow('Info/Frame.asp?FileName=SpecialDell.asp&PageTitle=删除专题&SpecialID='+SelectedClassObj.SpecialID,220,95,window);
	else alert('请选择修改的专题');
}
function ClassListVarObject(ParentID,ClassID)
{
	this.ParentID=ParentID;
	this.ClassID=ClassID;
}
function ClassListCilckContentMenu()
{
	var DisabledContentMenuStr='';
	if ((typeof(event.srcElement.onclick)=='function')&&(event.srcElement.tagName.toLowerCase()=='span')) event.srcElement.onclick();
	else
	{
		if (SelectedClassObj!=null) SelectedClassObj.className='TempletItem';
		ClassListVarObjectInstance.ClassID='0';
		ClassListVarObjectInstance.ParentID='0';
	}
	if (top.MainInfo!=null)
	{
		if ((top.MainInfo.SourceClass!='')||(top.MainInfo.SourceNews!='')||(top.MainInfo.SourceDownLoad!=''))
			DisabledContentMenuStr+='';
		else DisabledContentMenuStr+=',粘贴,';
		if (top.MainInfo.SourceClass!='') DisabledContentMenuStr+='';
		else DisabledContentMenuStr+=',合并,';
	}
	if (ClassListVarObjectInstance.ClassID!='0') DisabledContentMenuStr+='';
	else DisabledContentMenuStr+=',删除,修改,剪切,栏目初始化,预览,';
	if (ClassListVarObjectInstance.ParentID=='110110')
	{
		DisabledContentMenuStr=',添加,粘贴,合并,剪切,栏目初始化,投稿管理,预览,';
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function MergeClass()
//合并栏目
{
	var Url='Info/Frame.asp?PageTitle=合并栏目&FileName=MergeClass.asp&ObjectClass='+ClassListVarObjectInstance.ClassID;
	Url=Url+'&SourceClass='+top.MainInfo.SourceClass;
	OpenWindow(Url,330,110,window);
	top.GetNavFoldersObject().location.href=top.GetNavFoldersObject().location.href
}

function MergeSpecial()
{//合并专题
	var Url='Info/Frame.asp?PageTitle=合并专题&FileName=MergeSpecial.asp&SourceSpecial='+SelectedClassObj.SpecialID;
	OpenWindow(Url,280,110,window);
	top.GetNavFoldersObject().location.href=top.GetNavFoldersObject().location.href
}
function SpecialInit()
{//专题初始化
	var Url='Info/Frame.asp?PageTitle=专题初始化&FileName=SpecialInit.asp&SpecialID='+SelectedClassObj.SpecialID;
	OpenWindow(Url,280,110,window);
	top.GetEkMainObject().location.href=top.GetEkMainObject().location.href;
}
function ClassInit()
{//栏目初始化
	var Url='Info/Frame.asp?PageTitle=栏目初始化&FileName=ClassInit.asp&ClassID='+ClassListVarObjectInstance.ClassID;
	OpenWindow(Url,280,110,window);
	top.GetEkMainObject().location.href=top.GetEkMainObject().location.href;
}
function SpecialListCilckContentMenu()
{
	var DisabledContentMenuStr='';
	if ((typeof(event.srcElement.onclick)=='function')&&(event.srcElement.tagName.toLowerCase()=='span')) event.srcElement.onclick();
	else 
	{
		if (SelectedClassObj!=null) SelectedClassObj.className='TempletItem';
		SelectedClassObj=null;
	}
	if (SelectedClassObj==null)
	{
		DisabledContentMenuStr+=',修改,删除,合并到,初始化,';
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function UpLoadCilckContentMenu()
{
	var DisabledContentMenuStr='';
	if ((typeof(event.srcElement.onclick)=='function')&&(event.srcElement.tagName.toLowerCase()=='span')) event.srcElement.onclick();
	else 
	{
		if (SelectedClassObj!=null) SelectedClassObj.className='TempletItem';
		SelectedClassObj=null;
	}
	if (SelectedClassObj==null)
	{
		DisabledContentMenuStr+=',删除栏目,重命名,';
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function JSCilckContentMenu()
{
	var DisabledContentMenuStr='';
	if ((typeof(event.srcElement.onclick)=='function')&&(event.srcElement.tagName.toLowerCase()=='span')) event.srcElement.onclick();
	else 
	{
		if (SelectedClassObj!=null) SelectedClassObj.className='TempletItem';
		SelectedClassObj=null;
	}
	if (SelectedClassObj==null)
	{
		DisabledContentMenuStr+=',修改自由JS,删除自由JS';
	}
	else
	{
		if (SelectedClassObj.JsID!=null) DisabledContentMenuStr+='';
		else DisabledContentMenuStr+=',修改自由JS,删除自由JS';
	}
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function AddClass()
{
	top.GetEkMainObject().location.href='Info/ClassAdd.asp?ParentID='+ClassListVarObjectInstance.ClassID;
}

function ContributionMan()
{
	top.GetEkMainObject().location.href='Info/ContributionList.asp?ClassID='+ClassListVarObjectInstance.ClassID;
}
function EditClass()
{
	top.GetEkMainObject().location.href='Info/ClassEdit.asp?ClassID='+ClassListVarObjectInstance.ClassID;
}
function ViewClass()
{
	window.open('Info/read.asp?Table=NewsClass&ID='+ClassListVarObjectInstance.ClassID);
}
function DelClass()
{
	if (ClassListVarObjectInstance.ClassID!='')
	{
		OpenWindow('Info/Frame.asp?FileName=DelContent.asp&PageTitle=删除栏目&Operation=DelClass&ClassID='+ClassListVarObjectInstance.ClassID,300,101,window);
	}
}
function ClassCut()
{
	top.MainInfo.SourceClass=ClassListVarObjectInstance.ClassID;
	top.MainInfo.SourceNews='';
	top.MainInfo.SourceDownLoad='';
	top.MainInfo.ObjectClass='';
	top.MainInfo.MoveTF=true;
	top.MainInfo.OperationType='Class';
}
function ClassPaste()
{
	top.MainInfo.ObjectClass=ClassListVarObjectInstance.ClassID;
	var MoveOrCopyClassPara='OperationType:'+top.MainInfo.OperationType+',MoveTF:'+top.MainInfo.MoveTF+',SourceClass:'+top.MainInfo.SourceClass+',ObjectClass:'+top.MainInfo.ObjectClass+',';
	OpenWindow('Info/Frame.asp?FileName=MoveOrCopyNewsClass.asp&Titles=新闻栏目移动复制&MoveOrCopyClassPara='+MoveOrCopyClassPara,300,150,window);
}
function OpenAllParentClassList()
{
	var CurrObj=null;
	var TempArray=OpenClassIDList.split(',');
	var ClickParentClassIDStr='';
	var ClickEndClassIDStr='';
	for (var j=0;j<TempArray.length;j++)
	{
		if (j<TempArray.length-1)
		{
			if (ClickParentClassIDStr=='') ClickParentClassIDStr=TempArray[j];
			else  ClickParentClassIDStr=ClickParentClassIDStr+','+TempArray[j];
		}
		else ClickEndClassIDStr=TempArray[j];
	}
	if ((ClickParentClassIDStr!='')||(ClickEndClassIDStr!=''))
	{
		var TRObj=document.body.getElementsByTagName('TR');
		for (var i=0;i<TRObj.length;i++)
		{
			CurrObj=TRObj(i);
			if ((CurrObj.ClassID!=null)&&(CurrObj.tagName.toLowerCase()=='tr')&&(CurrObj.ClassID!='0'))
			{
				if (ClickParentClassIDStr.indexOf(CurrObj.ClassID)!=-1)
				{
					CurrObj.children(0).children(0).children(0).children(0).children(0).children(CurrObj.children(0).children(0).children(0).children(0).children(0).children.length-2).onclick();
				}
				if (CurrObj.ClassID==ClickEndClassIDStr)
				{
					var ParentObj=CurrObj.children(0).children(0).children(0).children(0).children(0).children(CurrObj.children(0).children(0).children(0).children(0).children(0).children.length-2);
					ParentObj.onclick();
					CurrObj.children(0).children(0).children(0).children(0).children(1).children(0).onclick();
				}
			}
		}
	}
}
//内容JS
//Select Page
function GetLocation(TypeStr,Obj)
{
	var LocationStr='';
//	if (TypeStr.slice(0,6)=='FreeJS')
//	{
//		LocationStr='JS/FreeJSFileList.asp?JSID='+TypeStr.slice(6);
//		return LocationStr;
//	}
	switch (TypeStr)
	{
		case 'FreeLable':
			LocationStr='Templet/Templet_FreeLable.asp';
			break;
		case "PictureTools":
			LocationStr='System/Tool_PictureModify.asp';
			break;
		case 'VirtualFolder':
			LocationStr='Info/UpFileList.asp?Path='+Obj.Path;
			break;
		case 'TempletManage':
			LocationStr='Templet/NewsTemplet_List.asp?Path='+Obj.Path;
			break;
		case 'AddLable':
			LocationStr='Templet/Templet_LableList.asp';
			break;
		case 'BackUpLable':
			LocationStr='Templet/Templet_LableBackUp.asp';
			break;
		case 'MallStyle':
			LocationStr='Templet/Templet_MallStyleList.asp';
			break;
		case 'DownLoadStyle':
			LocationStr='Templet/Templet_DownStyleList.asp';
			break;
		case 'RefreshIndex':
			LocationStr='Refresh/RefreshIndex.asp';
			break;
		case 'RefreshClass':
			LocationStr='Refresh/RefreshClass.asp';
			break;
		case 'RefreshNews':
			LocationStr='Refresh/RefreshNews.asp';
			break;
		case 'RefreshSpecial':
			LocationStr='Refresh/RefreshSpecial.asp';
			break;
		case 'RefreshDownLoad':
			LocationStr='Refresh/RefreshDownLoad.asp';
			break;
		case 'RefreshDoMain':
			LocationStr='Refresh/RefreshDoMain.asp';
			break;
		case 'RefreshJS':
			LocationStr='Refresh/RefreshAllJS.asp';
			break;
		case 'Mall_Refresh':
			LocationStr='Refresh/Mall_Refresh.asp';
			break;
		case 'Special':
			LocationStr='Info/SpecialFileList.asp?SpecialID='+Obj.SpecialID;
			break;
		case 'AdminGroup':
			LocationStr='System/SysAdminGroup.asp';
			break;
		case 'Admin':
			LocationStr='System/SysAdminList.asp';
			break;
		case 'ShortCut':
			LocationStr='System/SetShortCut.asp';
			break;
		case 'NewsSystemPara':
			LocationStr='System/SysParameter.asp';
			break;
		case 'DownLoadSystemPara':
			LocationStr='System/DownLoadParameter.asp';
			break;
		case 'NetStationPara':
			LocationStr='System/SysConstSet.asp';
			break;
		case 'DataStat':
			LocationStr='System/DataBase_Statistic.asp';
			break;
		case 'DBSpace':
			LocationStr='System/DataBase_Space.asp';
			break;
		case 'DBBackAndPress':
			LocationStr='System/DataBase_Operate.asp';
			break;
		case 'ExeSql':
			LocationStr='System/DataBase_ExeCuteSql.asp';
			break;
		case 'LogManage':
			LocationStr='System/DataBase_LogManage.asp';
			break;
		case 'UserGroup':
			LocationStr='System/SysUserGroup.asp';
			break;
		case 'User':
			LocationStr='System/SysUserList.asp';
			break;
		case 'UserNews':
			LocationStr='System/SysUserNews.asp';
			break;
		case 'RecyleManage':
			LocationStr='Recycle/Folder.asp';
			break;
		case 'AdsManage':
			LocationStr='Ads/AdsList.asp';
			break;
		case 'VoteManage':
			LocationStr='Vote/VoteList.asp';
			break;
		case 'KeyWordManage':
			LocationStr='Info/OrdinaryList.asp?Type=1';
			break;
		case 'SourceManage':
			LocationStr='Info/OrdinaryList.asp?Type=2';
			break;
		case 'AuthorManage':
			LocationStr='Info/OrdinaryList.asp?Type=3';
			break;
		case 'EditerManage':
			LocationStr='Info/OrdinaryList.asp?Type=4';
			break;
		case 'InnerLinkManage':
			LocationStr='Info/OrdinaryList.asp?Type=5';
			break;
		case 'FriendLinkManage':
			LocationStr='Info/OrdinaryFriendLink.asp';
			break;
		case 'CollectSiteSet':
			LocationStr='Collect/Site.asp';
			break;
		case 'CollectKeyWork':
			LocationStr='Collect/Rule.asp';
			break;
		case 'CollectAuditData':
			LocationStr='Collect/Check.asp';
			break;
		case 'CollectHistoryData':
			LocationStr='Collect/NewsOfHistory.asp';
			break;
		case 'CollectRuleManage':
			LocationStr='Collect/UpDateManage.asp';
			break;
		case 'CollectDataMove':
			LocationStr='Collect/MoveDataManage.asp';
			break;
		case 'PlusManage':
			LocationStr='Plus/PlusList.asp';
			break;
		case 'HelpManage':
			LocationStr='../help/SearchManage.asp';
			break;
		case 'ClassJS':
			LocationStr='JS/ClassJsList.asp?Types=Class';
			break;
		case 'SystemJS':
			LocationStr='JS/ClassJsList.asp?Types=System';     
			break;
		case 'ClassJSCode':
			LocationStr='JS/CodeSysJSList.asp?Types=Class';
			break;
		case 'SysJSCode':
			LocationStr='JS/CodeSysJSList.asp?Types=System';
			break;
		case 'CodeFreeJS':
			LocationStr='JS/CodeFreeJsList.asp';
			break;
		case 'AdsJSCode':
			LocationStr='JS/CodeAdsList.asp';
			break;
		case 'SFreeJSList':
			LocationStr='JS/FreeJSList.asp';
			break;
		case 'CountStatDayStat':
			LocationStr='System/Visit_DaysStatistic.asp';
			break;
		case 'CountStatVisitStat':
			LocationStr='System/Visit_VisitorList.asp';
			break;
		case 'CountStatSourceStat':
			LocationStr='System/Visit_SourceStatistic.asp';
			break;
		case 'CountStatAreaStat':
			LocationStr='System/Visit_AreaStatistic.asp';
			break;
		case 'CountStatBrowerStat':
			LocationStr='System/Visit_SystemStatistic.asp';
			break;
		case 'CountStatMonthStat':
			LocationStr='System/Visit_MonthsStatistic.asp';
			break;
		case 'CountStat24HoursStat':
			LocationStr='System/Visit_HoursStatistic.asp';
			break;
		case 'CountStatShortData':
			LocationStr='System/Visit_DataStatistic.asp';
			break;
		case 'CountStatNetManage':
			LocationStr='System/Visit_WebMaintenance.asp';
			break;
		case 'CountStatGetCode':
			LocationStr='System/Visit_ObtainCode.asp';
			break;
		case 'FileManage':
			LocationStr='FileManage/NewsList.asp';
			break;
		case 'ReplaceData':
			LocationStr='../funpages/Replacedata.asp';
			break;
		case 'DWHelp':
			LocationStr='templet/DWHelp.asp';
			break;
		case 'Mall_AddProducts':
			LocationStr='Mall/Mall_AddProducts.asp';
			break;
		case 'Mall_Products':
			LocationStr='Mall/Mall_Products.asp';
			break;
		case 'Mall_Class':
			LocationStr='Mall/Mall_Class.asp';
			break;
		case 'Mall_Special':
			LocationStr='Mall/Mall_Special.Asp';
			break;
		case 'Mall_Factory':
			LocationStr='Mall/Mall_Factory.asp';
			break;
		case 'Mall_Order':
			LocationStr='Mall/Mall_Order.asp';
			break;
		case 'Mall_integral':
			LocationStr='Mall/Mall_integral.asp';
			break;
		case 'Mall_Help':
			LocationStr='Mall/Mall_Help.asp';
			break;
		case 'Mall_Pay':
			LocationStr='Mall/Mall_Pay.asp';
			break;
		case 'system_Book':
			LocationStr='System/SysBook.asp';
			break;
		case 'Mall_Review':
			LocationStr='Mall/Mall_Review.asp';
			break;
		case 'Mall_AllData':
			LocationStr='Mall/AllData.asp';
			break;
		case 'Mall_Pmf':
			LocationStr='Mall/pmf.asp';
			break;
		case 'Mall_Config':
			LocationStr='Mall/Mall_Config.asp';
			break;
		case 'Mall_OnlinePay':
			LocationStr='Mall/Mall_OnlinePay.asp';
			break;
		default:
			LocationStr='';
			break;
	}
	return LocationStr;
}
//Select Page
</script>
<script src="../SysJS/PublicJS.js" language="JavaScript"></script>
<% if Action = "ContentTree" OR Action = "Special" OR Action = "UpLoad" then %>
<script language="JavaScript">
var ObjPopupMenu=window.createPopup();
document.oncontextmenu=new Function("return ShowMouseRightMenu(window.event);");
function ShowMouseRightMenu(event)
{
	ContentMenuShowEvent();
	var width=100;
	var height=0;
	var lefter=event.clientX;
	var topper=event.clientY;
	var ObjPopDocument=ObjPopupMenu.document;
	var ObjPopBody=ObjPopupMenu.document.body;
	var MenuStr='';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (ContentMenuArray[i].ExeFunction=='seperator')
		{
			MenuStr+=FormatSeperator();
			height+=16;
		}
		else
		{
			MenuStr+=FormatMenuRow(ContentMenuArray[i].ExeFunction,ContentMenuArray[i].Description,ContentMenuArray[i].EnabledStr);
			height+=20;
		}
	}
	MenuStr="<TABLE border=0 cellpadding=0 cellspacing=0 class=Menu width=100>"+MenuStr
	MenuStr=MenuStr+"<\/TABLE>";
	ObjPopDocument.open();
	ObjPopDocument.write("<head><link href=\"../../CSS/ContentMenu.css\" type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\" onConTextMenu=\"event.returnValue=false;\" onselectstart=\"event.returnValue=false;\">"+MenuStr);
	ObjPopDocument.close();
	height+=4;
	if(lefter+width > document.body.clientWidth) lefter=lefter-width;
	ObjPopupMenu.show(lefter, topper, width, height, document.body);
	return false;
}
function FormatSeperator()
{
	var MenuRowStr="<tr><td height=16 valign=middle><hr><\/td><\/tr>";
	return MenuRowStr;
}
function FormatMenuRow(MenuOperation,MenuDescription,EnabledStr)
{
	var MenuRowStr="<tr "+EnabledStr+"><td align=left height=20 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut'; valign=middle"
	if (EnabledStr=='') MenuRowStr+=" onclick=\""+MenuOperation+"parent.ObjPopupMenu.hide();\">&nbsp;&nbsp;&nbsp;&nbsp;";
	else MenuRowStr+=">&nbsp;&nbsp;&nbsp;&nbsp;";
	MenuRowStr=MenuRowStr+MenuDescription+"<\/td><\/tr>";
	return MenuRowStr;
}
</script>
<% end if%>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<body topmargin="0" leftmargin="2" onselectstart="return false;">
<%
If Action = "mall" then
	if Not JudgePopedomTF(Session("Name"),"P090000") then Call ReturnError1()
	If HaveValueTF=False Then
	    Response.Write("<script>alert(""商城未开放！"&CopyRight&""");location=""javascript:history.back()"";</script>")  
		Response.End
	End If
	%>
	
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/folderopen.gif"></td>
          <td class="TempletItem">商城</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Mall_Special');" Type="Mall_Special">专区管理</span></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="11"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Mall_Factory');" Type="Mall_Factory">厂家管理</span></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Mall_Order');" Type="Mall_Order">定单管理</span></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Mall_integral');" Type="Mall_integral">积分/金币</span></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Mall_OnlinePay');" Type="Mall_OnlinePay">在线支付设置</span></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Mall_Pay');" Type="Mall_Pay">邮寄资料</span></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Mall_AllData');" Type="Mall_AllData">综合统计</span></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Mall_Pmf');" Type="Mall_AllData">配送须知</span></td>
        </tr>
      </table></td>
  </tr>
</table>
	<%
elseif Action = "ContentTree" then
	if Not JudgePopedomTF(Session("Name"),"P010000") then Call ReturnError1()
	Dim ClassSql,RsClassObj
	ClassSql = "Select * from FS_NewsClass where parentID='0' and DelFlag=0 order by Orders Desc"
	Set RsClassObj = Server.CreateObject(G_FS_RS)
	RsClassObj.Open ClassSql,Conn,1,1
%>
<table border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr ParentID="0" ClassID="0" align="left" class="TempletItem">
          <td><img src="../Images/Station.gif" width="24" height="22"></td>
          <td>站点</td>
        </tr>
      </table></td>
  </tr>
  <%
	Dim ClassNumber,TempImageSrc,TempFolderImageSrc,ContributionNum,ContributionStr,SecondDoMainTF
	ClassNumber = 1
	do while Not RsClassObj.Eof
		if ClassNumber = RsClassObj.RecordCount then
			TempImageSrc = "../Images/Folder/EndClose.gif"
		else
			TempImageSrc = "../Images/Folder/Close.gif"
		end if
		if (Not IsNull(RsClassObj("DoMain"))) And (RsClassObj("DoMain") <> "") then
			TempFolderImageSrc = "../Images/DoMain.gif"
			SecondDoMainTF = "1"
		else
			TempFolderImageSrc = "../Images/Folder/folderclosed.gif"
			SecondDoMainTF = "0"
		end if
		if RsClassObj("Contribution") = 1 then
			ContributionNum = Conn.Execute("Select Count(ContID) from FS_Contribution where ClassID='" & RsClassObj("ClassID") & "'")(0)
			ContributionStr = "(" & ContributionNum & ")"
		else
			ContributionStr = ""
		end if
%>
  <tr AllParentID="0" ParentID="<% = RsClassObj("ParentID") %>" ClassID="<% = RsClassObj("ClassID") %>"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr align="left">
			<td nowrap><img onClick="ClickClassImg(this,'<% = RsClassObj("ClassID") %>');" src="<% = TempImageSrc %>"><img src="<% = TempFolderImageSrc %>"></td>
			<%
			If RsClassObj("IsOutClass")=0 then
			%>
				<td nowrap><span DoMain="<% = SecondDoMainTF %>" onClick="ClickFolderTxt(this,'<% = RsClassObj("ClassID") %>','<% = RsClassObj("ParentID") %>','<%=RsClassObj("RedirectList")%>');" class="TempletItem"><% = RsClassObj("ClassCName") & ContributionStr %></span></td>
			<%
			Else
			%>
				<td nowrap><span DoMain="<% = SecondDoMainTF %>" onClick="ClickFolderTxt(this,'<% = RsClassObj("ClassID") %>','110110','0');" class="TempletItem"><% = RsClassObj("ClassCName") %></span></td>
			<%
			End If
			%>
			
        </tr>
      </table></td>
  </tr>
  <%
		if ClassNumber = RsClassObj.RecordCount then
			Response.Write(GetChildClassList(RsClassObj("ClassID"),"",true,""))
		else
			Response.Write(GetChildClassList(RsClassObj("ClassID"),"",False,""))
		end if
		ClassNumber = ClassNumber + 1
		RsClassObj.MoveNext
	loop
	Set RsClassObj = Nothing
%>
</table>
<%
elseif Action = "Special" then
	if Not JudgePopedomTF(Session("Name"),"P020000") then Call ReturnError1()
	Dim RsSpecialObj,SpecialNumber,TempSpecialImageSrc
	SpecialNumber = 1
	Set RsSpecialObj = Server.CreateObject(G_FS_RS)
	RsSpecialObj.Open "select * from FS_Special order by ID desc",Conn,1,1
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../Images/Folder/folderopen.gif"></td>
          <td class="TempletItem">频道/专题管理</td>
        </tr>
      </table></td>
  </tr>
<%
	do while Not RsSpecialObj.Eof
		if SpecialNumber = RsSpecialObj.RecordCount then
			TempSpecialImageSrc = "../Images/Folder/lastnode.gif"
		else
			TempSpecialImageSrc = "../Images/Folder/Node.gif"
		end if
%>
  <tr height="20"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr align="left">
          <td><img src="<% = TempSpecialImageSrc %>"><img src="../Images/Folder/folderclosed.gif"></td>
          <td><span onClick="ClickBtn(this,'Special');" SpecialID="<%=RsSpecialObj("SpecialID")%>" class="TempletItem"><% = RsSpecialObj("CName") %></span></td>
        </tr>
      </table></td>
  </tr>
  <%
  		SpecialNumber = SpecialNumber + 1
		RsSpecialObj.MoveNext
	loop
	RsSpecialObj.close
	Set RsSpecialObj=nothing
%>
</table>
<%
elseif Action = "OrdinaryManage" then
	if Not(JudgePopedomTF(Session("Name"),"P070000") Or JudgePopedomTF(Session("Name"),"P080000")) then Call ReturnError1()
	if JudgePopedomTF(Session("Name"),"P070000") then
		Dim ContralImage,HRImage
		If JudgePopedomTF(Session("Name"),"P080000") Then
			ContralImage = "../Images/Folder/Open.gif"
			HRImage = "../Images/Folder/HR.gif"
		Else
			ContralImage = "../Images/Folder/EndClose.gif"
			HRImage = "../Images/Folder/blank.gif"
		End if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" height="22">
   <tr>
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../Images/Folder/folderopen.gif"></td>
          <td class="TempletItem">常规管理</td>
        </tr>
      </table></td>
  </tr>
  
  <tr AllParentID="0" ParentID="0" ClassID="OrdinaryMan"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img onClick="ClickClassImg(this,'OrdinaryMan');" src="<%=ContralImage%>" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" onClick="ClickFolder(this,'OrdinaryManage')" width="18" height="18"></td>
          <td class="TempletItem">辅助管理</td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="OrdinaryMan" ParentID="OrdinaryMan" ClassID="OrdinaryManContainer">
  <td>
	<table width="100%" height="22" border="0" cellpadding="0" cellspacing="0">
        <tr AllParentID="0" ParentID="0" ClassID="RecyleManage" style="display:;">
          <td height="22"><table width="116" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="50"><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td width="66"><span class="TempletItem" onClick="ClickBtn(this,'system_Book');" Type="Mall_Book">留言本管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="RecyleManage" style="display:;"> 
          <td height="22"><table border="0" cellpadding="0" cellspacing="0" height="22">
              <tr> 
                <td height="22"><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'RecyleManage');" class="TempletItem">回收站管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="FileManage" style="display:;"> 
          <td height="22"> <table border="0" cellspacing="0" cellpadding="0" height="22">
              <tr> 
                <td height="22"><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'FileManage');" class="TempletItem">归档管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="AdsManage" style="display:;"> 
          <td><table border="0" cellpadding="0" cellspacing="0" height="22">
              <tr> 
                <td height="22"><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'AdsManage');" class="TempletItem">广告管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="VoteManage" style="display:;"> 
          <td><table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'VoteManage');" class="TempletItem">投票管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="KeyWords" style="display:;"> 
          <td> <table border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'KeyWordManage');" class="TempletItem">关键字管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="Source" style="display:;"> 
          <td> <table height="19" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'SourceManage');" class="TempletItem">来源管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="Author" style="display:;"> 
          <td> <table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'AuthorManage');" class="TempletItem">作者管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="Editer" style="display:;"> 
          <td> <table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'EditerManage');" class="TempletItem">责任编辑管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="InnerLink" style="display:;"> 
          <td> <table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'InnerLinkManage');" class="TempletItem">内部链接管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="FriendLink" style="display:;"> 
          <td><table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'FriendLinkManage');" class="TempletItem">友情链接管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="FriendLink" style="display:;"> 
          <td><table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'PlusManage');" class="TempletItem">插件管理</span></td>
              </tr>
            </table></td>
        </tr>
        <tr AllParentID="0" ParentID="0" ClassID="FriendLink" style="display:;"> 
          <td><table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="<%=HRImage%>" width="16" height="22"><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
                <td><span onClick="ClickBtn(this,'HelpManage');" class="TempletItem">帮助管理</span></td>
              </tr>
            </table></td>
        </tr>
      </table>
</td>
</tr>
</table>
<%
	end if
	if JudgePopedomTF(Session("Name"),"P080000") then
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr AllParentID="0" ParentID="0" ClassID="Toolers"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img onClick="ClickClassImg(this,'Toolers');" src="../Images/Folder/EndClose.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td class="TempletItem">辅助工具</td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="Toolers" ParentID="Toolers" ClassID="ToolersContainer"  style="display:none">
  <td>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr AllParentID="0" ParentID="0" ClassID="FileManage"> 
		<td> <table border="0" cellspacing="0" cellpadding="0">
			<tr> 
			  <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			    <td><span onClick="ClickBtn(this,'DWHelp');" class="TempletItem">DW插件辅助</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="0" ParentID="0" ClassID="FileManage"> 
		<td> <table border="0" cellspacing="0" cellpadding="0">
			<tr> 
			  <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'ReplaceData');" class="TempletItem">字段内容替换</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="0" ParentID="0" ClassID="NewsCollect"> 
		<td> <table border="0" cellspacing="0" cellpadding="0">
			<tr> 
			  <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img onClick="ClickClassImg(this,'NewsCollect');" src="../Images/Folder/Open.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td class="TempletItem">新闻采集</td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="NewsCollect" ParentID="NewsCollect" ClassID="SiteSet" style="display:;"> 
		<td> <table border="0" cellspacing="0" cellpadding="0">
			<tr> 
			  <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CollectSiteSet');" class="TempletItem">站点设置</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="NewsCollect" ParentID="NewsCollect" ClassID="KeyWordsManage" style="display:;"> 
		<td> <table border="0" cellspacing="0" cellpadding="0">
			<tr> 
			  <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CollectKeyWork');" class="TempletItem">关键字过滤</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="NewsCollect" ParentID="NewsCollect" ClassID="AuditData" style="display:;"> 
		<td> <table border="0" cellspacing="0" cellpadding="0">
			<tr> 
			  <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CollectAuditData');" class="TempletItem">审核数据</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="NewsCollect" ParentID="NewsCollect" ClassID="HistoryData" style="display:;"> 
		<td> <table border="0" cellspacing="0" cellpadding="0">
			<tr> 
			  <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CollectHistoryData');" class="TempletItem">历史数据</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="0" ParentID="0" ClassID="CountStat"> 
		<td> <table border="0" cellspacing="0" cellpadding="0">
			<tr> 
			  <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img onClick="ClickClassImg(this,'CountStat');" src="../Images/Folder/EndOpen.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td class="TempletItem">流量统计</td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="GetCode" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr> 
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStatGetCode');" class="TempletItem">获取代码</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="NetStationManage" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr> 
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStatNetManage');" class="TempletItem">网站维护</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="DataMove" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr> 
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStatShortData');" class="TempletItem">简要数据</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="ShortData" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr> 
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStat24HoursStat');" class="TempletItem">24小时统计</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="DataMove" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr>
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStatDayStat');" class="TempletItem">日统计</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="DataMove" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr>
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStatMonthStat');" class="TempletItem">月统计</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="DataMove" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr>
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStatBrowerStat');" class="TempletItem">系统/浏览器统计</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="DataMove" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr>
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStatAreaStat');" class="TempletItem">地区统计</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="DataMove" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr>
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStatSourceStat');" class="TempletItem">来源统计</span></td>
			</tr>
		  </table></td>
	  </tr>
	  <tr AllParentID="CountStat" ParentID="CountStat" ClassID="DataMove" style="display:;"> 
		<td> <table border="0" cellpadding="0" cellspacing="0">
			<tr> 
			    <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
			  <td><span onClick="ClickBtn(this,'CountStatVisitStat');" class="TempletItem">来访者信息统计</span></td>
			</tr>
		  </table></td>
	  </tr>
	</table>
  </td>
</tr>
</table>
<%
	end if
elseif Action = "NetStation" then
	Dim TempletDirectory
	if Not JudgePopedomTF(Session("Name"),"P030000") then Call ReturnError1()
	if SysRootDir <> "" then
		TempletDirectory = "/" & SysRootDir & "/" & TempletDir
	else
		TempletDirectory = "/" & TempletDir
	end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="100%" colspan="2"><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/folderopen.gif" width="18" height="18"></td>
          <td class="TempletItem">站点管理</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td colspan="2"> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" Path="<% = TempletDirectory %>" onClick="ClickBtn(this,'TempletManage');">模板管理</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="0" ParentID="0" ClassID="LableMan"> 
    <td colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr AllParentID="0"> 
          <td><img onClick="ClickClassImg(this,'LableMan');" src="../Images/Folder/Open.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" Type="LableList">标签管理</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="LableMan" ParentID="LableMan"> 
    <td colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td nowrap><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'AddLable');" Type="LableList">自定义标签</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="LableMan" ParentID="LableMan"> 
    <td> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td nowrap><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'FreeLable');" Type="LableList">自由标签</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="LableMan" ParentID="LableMan"> 
    <td colspan="2"><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'BackUpLable');" Type="CopyLable">备份标签</span></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'DownLoadStyle');" Type="DownListStyle">下载列表样式</span></td>
        </tr>
      </table></td>
  </tr>
  <%If HaveValueTF=True Then%>
  <tr> 
    <td colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'MallStyle');" Type="MallStyle">商城列表样式</span></td>
        </tr>
      </table></td>
  </tr>
  <%End If%>
  <tr AllParentID="0" ParentID="0" ClassID="RefreshMan"> 
    <td colspan="2"><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img onClick="ClickClassImg(this,'RefreshMan');" src="../Images/Folder/EndOpen.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td class="TempletItem">发布管理</td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="RefreshMan" ParentID="RefreshMan"> 
    <td colspan="2"> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td width="55"><span class="TempletItem" onClick="ClickBtn(this,'RefreshIndex');" Type="RefreshIndex">发布首页</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="RefreshMan" ParentID="RefreshMan"> 
    <td colspan="2"> <table height="19" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'RefreshClass');" Type="RefreshClass">发布栏目</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="RefreshMan" ParentID="RefreshMan"> 
    <td colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'RefreshNews');" Type="RefreshNews">发布新闻</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="RefreshMan" ParentID="RefreshMan"> 
    <td colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'RefreshSpecial');" Type="RefreshSpecial">发布专题</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="RefreshMan" ParentID="RefreshMan"> 
    <td colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'RefreshDownLoad');" Type="RefreshDownLoad">发布下载</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="RefreshMan" ParentID="RefreshMan"> 
    <td colspan="2"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/<%=ShowNode%>.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'RefreshJS');" Type="RefreshSpecial">发布JS</span></td>
        </tr>
      </table></td>
  </tr>
  <%If HaveValueTF=True Then%>
  <tr AllParentID="RefreshMan" ParentID="RefreshMan">
	<td><table border="0" cellspacing="0" cellpadding="0">
		<tr>
		  <td width="50"><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif"></td>
		  <td width="75"><span class="TempletItem" onClick="ClickBtn(this,'Mall_Refresh');" Type="Mall_Refresh">发布商城</span></td>
		</tr>
	  </table></td>
  </tr>
  <%End if %>
</table>
<%
elseif Action = "UpLoad" then
	if Not JudgePopedomTF(Session("Name"),"P050000") then Call ReturnError1()
	Dim UpFilesPath,FS,FolderObj,SubFolderObj,FolderItem,UpLoadNumber,TempUpLoadImgSrc
	if SysRootDir <> "" then
		UpFilesPath = "/" & SysRootDir & "/"
	else
		UpFilesPath = "/"
	end if
	UpLoadNumber = 1
	Set FS = Server.CreateObject(G_FS_FSO)
	Set FolderObj = FS.GetFolder(Server.MapPath(UpFilesPath))
	Set SubFolderObj = FolderObj.SubFolders
%>
<table id="RootDir" RootDir="<% = UpFilesPath %>" width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../Images/Folder/folderopen.gif" width="18" height="18"></td>
          <td><span Path="<% = UpFilesPath %>" onClick="ClickBtn(this,'VirtualFolder');" class="TempletItem">根目录</span></td>
	    </tr>
</table>
</td>
  </tr>
<%
For Each FolderItem in SubFolderObj
	if UpLoadNumber = SubFolderObj.Count then
		TempUpLoadImgSrc = "../Images/Folder/EndClose.gif"
	else
		TempUpLoadImgSrc = "../Images/Folder/Close.gif"
	end if
%>
  <tr AllParentID="<% = UpFilesPath %>" ParentID="<% = UpFilesPath %>" ClassID="<% = UpFilesPath & FolderItem.Name %>">
    <td> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img Depth="1" onClick="ClickClassImg(this,'<% = UpFilesPath & FolderItem.Name %>');" src="<% = TempUpLoadImgSrc %>" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span Path="<% = UpFilesPath & FolderItem.Name %>" onClick="ClickBtn(this,'VirtualFolder');" class="TempletItem"><% = FolderItem.Name %></span></td>
		</tr>
	  </table>
    </td>
  </tr>
<%
	if UpLoadNumber = SubFolderObj.Count then
		Response.Write(GetChildFolderList(UpFilesPath & FolderItem.Name,"",true,""))
	else
		Response.Write(GetChildFolderList(UpFilesPath & FolderItem.Name,"",False,""))
	end if
	UpLoadNumber = UpLoadNumber + 1
Next
Set FS = Nothing
Set FolderObj = Nothing
Set SubFolderObj = Nothing
%>
</table>
</td>
  </tr>
</table>
<%
elseif Action = "JSManage" then
	if Not JudgePopedomTF(Session("Name"),"P060000") then Call ReturnError1()
	Dim FreeJSSql,RsFreeJSObj,FreeJSNumber,TempFreeJSImgSrc
	FreeJSNumber = 1
	FreeJSSql = "Select * from FS_FreeJS order by Type asc,ID asc"
	Set RsFreeJSObj = Server.CreateObject(G_FS_RS)
	RsFreeJSObj.Open FreeJSSql,Conn,1,1
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../Images/Folder/folderopen.gif" width="18" height="18"></td>
          <td class="TempletItem">JS管理</td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="0" ParentID="0" ClassID="CMSJSManage"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr align="left"> 
          <td><img onClick="ClickClassImg(this,'CMSJSManage');" src="../Images/Folder/Open.gif" width="16" height="22"><img src="../Images/Folder/folderopen.gif" width="18" height="18"></td>
          <td class="TempletItem">JS设置</td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="CMSJSManage" ParentID="CMSJSManage" ClassID="ClassJS"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr align="left"> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'ClassJS');">栏目JS</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="CMSJSManage" ParentID="CMSJSManage" ClassID="SystemJS"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr align="left"> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SystemJS');">系统JS</span></td>
        </tr>
      </table></td>
  </tr>
    <tr AllParentID="CMSJSManage" ParentID="CMSJSManage" ClassID="FreeJS"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr align="left"> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SFreeJSList');">自由JS</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="0" ParentID="0" ClassID="JSCode"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr align="left"> 
          <td><img onClick="ClickClassImg(this,'JSCode');" src="../Images/Folder/EndOpen.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem">代码调用</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="JSCode" ParentID="JSCode"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr align="left"> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'ClassJSCode');">栏目JS</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="JSCode" ParentID="JSCode"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr align="left"> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'SysJSCode');">系统JS</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="JSCode" ParentID="JSCode"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr align="left"> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'CodeFreeJS');">自由JS</span></td>
        </tr>
      </table></td> 
  </tr> 
  <tr AllParentID="JSCode" ParentID="JSCode"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr align="left"> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'AdsJSCode');">广告JS</span></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
elseif Action = "System" then
	if Not JudgePopedomTF(Session("Name"),"P040000") then Call ReturnError1()
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/folderopen.gif" width="18" height="18"></td>
          <td class="TempletItem">系统管理</td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="0" ParentID="0" ClassID="UserManage"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img onClick="ClickClassImg(this,'UserManage');" src="../Images/Folder/Open.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td class="TempletItem">用户管理</td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="UserManage" ParentID="UserManage" ClassID="AdminGroup" style="display:;"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'AdminGroup');" class="TempletItem">管理员组</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="UserManage" ParentID="UserManage" ClassID="Admin" style="display:;"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'Admin');" class="TempletItem">管理员</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="UserManage" ParentID="UserManage" ClassID="UserGroup" style="display:;"> 
    <td> <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'UserGroup');" class="TempletItem">会员组</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="UserManage" ParentID="UserManage" ClassID="Users" style="display:;"> 
    <td><table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'User');" class="TempletItem">会员</span></td>
        </tr>
      </table></td>
  </tr>
  <%if HaveValueTF = True then%>
  <tr AllParentID="UserManage" ParentID="UserManage" ClassID="Users" style="display:;"> 
    <td><table width="101" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="50"><img src="../Images/Folder/HR.gif" width="16" height="22"><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td width="51"><span onClick="ClickBtn(this,'UserNews');" class="TempletItem">会员公告</span></td>
        </tr>
      </table></td>
  </tr>
  <%End if%>
  <tr AllParentID="0" ParentID="0" ClassID="SysPara"> 
    <td> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img onClick="ClickClassImg(this,'SysPara');" src="../Images/Folder/Open.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td class="TempletItem">系统参数</td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="SysPara" ParentID="SysPara" ClassID="NewsPara" style="display:;"> 
    <td> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td nowrap><img src="../Images/Folder/HR.gif" width="16" height="22"><img onClick="ClickClassImg(this,'NewsPara');" src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'NewsSystemPara');" class="TempletItem">新闻系统参数</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="SysPara" ParentID="SysPara" ClassID="DownLoadPara" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img onClick="ClickClassImg(this,'NewsPara');" src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'DownLoadSystemPara');" class="TempletItem">下载系统参数</span></td>
        </tr>
      </table></td>
  </tr>
  <%If HaveValueTF=True Then%>
  <tr> 
    <td><table width="127" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="52"><img src="../Images/Folder/HR.gif" width="16" height="22"><img onClick="ClickClassImg(this,'NewsPara');" src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span class="TempletItem" onClick="ClickBtn(this,'Mall_Config');" Type="Mall_Config">商城参数设置</span></td>
        </tr>
      </table></td>
  </tr>
  <%End if%>
  <tr AllParentID="SysPara" ParentID="SysPara" ClassID="ConstSet" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/HR.gif" width="16" height="22"><img onClick="ClickClassImg(this,'ConstSet');" src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'NetStationPara');" class="TempletItem">站点常量设置</span></td>
        </tr>
      </table></td>
  </tr>
  <tr style="display:none;" AllParentID="0" ParentID="0" ClassID="SysPara"> 
    <td> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'ShortCut');" class="TempletItem">快捷菜单管理</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="0" ParentID="0" ClassID="DBMan"> 
    <td> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img onClick="ClickClassImg(this,'DBMan');" src="../Images/Folder/EndOpen.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td class="TempletItem">数据库管理</td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="DBMan" ParentID="DBMan" ClassID="DataStat" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'DataStat');" class="TempletItem">数据统计</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="DBMan" ParentID="DBMan" ClassID="Space" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'DBSpace');" class="TempletItem">空间占用</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="DBMan" ParentID="DBMan" ClassID="ExeSql" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/Node.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'ExeSql');" class="TempletItem">执行SQL脚本</span></td>
        </tr>
      </table></td>
  </tr>
  <tr AllParentID="DBMan" ParentID="DBMan" ClassID="LogMan" style="display:;"> 
    <td><table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../Images/Folder/blank.gif" width="16" height="22"><img src="../Images/Folder/lastnode.gif" width="16" height="22"><img src="../Images/Folder/folderclosed.gif" width="18" height="18"></td>
          <td><span onClick="ClickBtn(this,'LogManage');" class="TempletItem">后台日志管理</span></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
end if
%>



</body>
</html>
<%
Set Conn = Nothing
Set RsMenuConfigObj = Nothing
Function GetChildClassList(ClassID,Str,EndNodeTF,TempAllParentID)
	Dim Sql,RsTempObj,TempImageStr,ImageStr,ChildClassNumber,AllParentID
	Dim TempSrc,TempEndNodeTF,ContributionNum,ContributionStr
	if EndNodeTF = True then
		TempSrc = "<img src=""../Images/Folder/blank.gif"">"
	else
		TempSrc = "<img src=""../Images/Folder/HR.gif"">"
	end if
	ChildClassNumber = 1
	AllParentID = TempAllParentID & "," & ClassID
	Sql = "Select * from FS_NewsClass where ParentID='" & ClassID & "' and DelFlag=0 order by Orders Desc"
	ImageStr = Str & TempSrc
	Set RsTempObj = Server.CreateObject(G_FS_RS)
	RsTempObj.Open Sql,Conn,1,1
	do while Not RsTempObj.Eof
		if ChildClassNumber = RsTempObj.RecordCount then
			TempEndNodeTF = True
			TempImageStr = "<img onClick=""ClickClassImg(this,'" & RsTempObj("ClassID") & "')"" src=""../Images/Folder/EndClose.gif""><img src=""../Images/Folder/folderclosed.gif"">"
		else
			TempEndNodeTF = False
			TempImageStr = "<img onClick=""ClickClassImg(this,'" & RsTempObj("ClassID") & "')"" src=""../Images/Folder/Close.gif""><img src=""../Images/Folder/folderclosed.gif"">"
		end if
		if RsTempObj("Contribution") = 1 then
			ContributionNum = Conn.Execute("Select Count(ContID) from FS_Contribution where ClassID='" & RsTempObj("ClassID") & "'")(0)
			ContributionStr = "(" & ContributionNum & ")"
		else
			ContributionStr = ""
		end if
		GetChildClassList = GetChildClassList & "<tr AllParentID=""" & AllParentID & """ ParentID=""" & RsTempObj("ParentID") & """ ClassID=""" & RsTempObj("ClassID") & """ style=""display:none;""><td><table border=""0"" cellspacing=""0"" cellpadding=""0""><tr align=""left"" class=""TempletItem""><td>" & ImageStr & TempImageStr & "</td><td nowrap><span DoMain=""0"" onClick=""ClickFolderTxt(this,'" & RsTempObj("ClassID") & "','" & RsTempObj("ParentID") & "','"&RsTempObj("RedirectList")&"');"">" & RsTempObj("ClassCName") & ContributionStr & "</span></td></tr></table></td></tr>" & Chr(13) & Chr(10)
		GetChildClassList = GetChildClassList & GetChildClassList(RsTempObj("ClassID"),ImageStr,TempEndNodeTF,AllParentID)
		ChildClassNumber = ChildClassNumber + 1
		RsTempObj.MoveNext
	loop
	Set RsTempObj = Nothing
End Function

Function GetChildFolderList(FolderID,Str,EndNodeTF,TempAllParentID)
	Dim TempImageStr,ImageStr,ChildFolderNumber,AllParentID
	Dim TempSrc,TempEndNodeTF
	Dim FS,FolderObj,SubFolderObj,FolderItem
	if EndNodeTF = True then
		TempSrc = "<img src=""../Images/Folder/blank.gif"">"
	else
		TempSrc = "<img src=""../Images/Folder/HR.gif"">"
	end if
	ChildFolderNumber = 1
	AllParentID = TempAllParentID & "," & FolderID
	ImageStr = Str & TempSrc
	Set FS = Server.CreateObject(G_FS_FSO)
	FolderID=replace(FolderID,"//","/")
	Set FolderObj = FS.GetFolder(Server.MapPath(FolderID))
	Set SubFolderObj = FolderObj.SubFolders
	For Each FolderItem in SubFolderObj
		if ChildFolderNumber = SubFolderObj.Count then
			TempEndNodeTF = True
			TempImageStr = "<img onClick=""ClickClassImg(this,'" & FolderID & "/" & FolderItem.Name & "')"" src=""../Images/Folder/EndClose.gif""><img src=""../Images/Folder/folderclosed.gif"">"
		else
			TempEndNodeTF = False
			TempImageStr = "<img onClick=""ClickClassImg(this,'" & FolderID & "/" & FolderItem.Name & "')"" src=""../Images/Folder/Close.gif""><img src=""../Images/Folder/folderclosed.gif"">"
		end if
		GetChildFolderList = GetChildFolderList & "<tr AllParentID=""" & AllParentID & """ ParentID=""" & FolderID & """ ClassID=""" & FolderID & "/" & FolderItem.Name & """ style=""display:none;""><td><table border=""0"" cellspacing=""0"" cellpadding=""0""><tr align=""left"" class=""TempletItem""><td>" & ImageStr & TempImageStr & "</td><td nowrap><span Path=""" & FolderID & "/" & FolderItem.Name & """ onClick=""ClickBtn(this,'VirtualFolder');"" class=""TempletItem"">" & FolderItem.Name & "</span></td></tr></table></td></tr>" & Chr(13) & Chr(10)
		GetChildFolderList = GetChildFolderList & GetChildFolderList(FolderID & "/" & FolderItem.Name,ImageStr,TempEndNodeTF,AllParentID)
		ChildFolderNumber = ChildFolderNumber + 1
	Next
	Set FS = Nothing
	Set FolderObj = Nothing
	Set SubFolderObj = Nothing
End Function
%>