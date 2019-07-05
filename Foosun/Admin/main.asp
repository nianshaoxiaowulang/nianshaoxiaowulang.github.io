<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
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
Dim DBC,conn,sConn
Set DBC = new databaseclass
Set Conn = DBC.openconnection()
Set DBC = Nothing
Dim RsConfigLoginobj,SiteName
SiteName = ""
Set RsConfigLoginobj=Conn.execute("Select SiteName from FS_Config")
if Not RsConfigLoginobj.Eof then
	SiteName = RsConfigLoginobj("SiteName")
end if
Set RsConfigLoginobj = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE><% = SiteName %>___网站内容管理系统___后台管理-Powered by Foosun Inc.</TITLE>
<meta name="keywords" content="Foosun,FoosunCMS,Foosun Inc.,风讯,风讯网站内容管理系统,风讯系统,风讯新闻系统，风讯商城,风讯b2c,新闻系统,CMS,域名空间,asp,jsp,asp.net,SQL,SQL SERVER">
</head>
<script language="JavaScript">
var m_SelectMainWindowTimerId=0;
var m_MakeNavTreeVisibleTimerId=0;
function StartShrink(e)
{
	var NavTreeObject = GetNavTreeObject();
	if (NavTreeObject)	NavTreeObject.StartShrink(e);
}
function StartEnlarge(e)
{
	var NavTreeObject = GetNavTreeObject();
	if (NavTreeObject) NavTreeObject.StartEnlarge(e);
} 
function GetNavTreeObject()
{
	return (frames["fs_nav_bottom"]);
}
function ValidateObjectsFrames(testObject)
{
	return ((typeof(testObject.frames) == "object")&&(testObject.frames.length > 0));
}
function CanShowNavTree()
{
	var EkMainObject = GetEkMainObject();
	if (EkMainObject)
	{
		if (typeof(EkMainObject.CanShowNavTree) == "function")
		{
			return (EkMainObject.CanShowNavTree());
		}
	}
	return (true);
}
function GetEkMainObject()
{
	if (ValidateObjectsFrames(this)&&ValidateObject(frames["fs_main"], false))
	{
		return (frames["fs_main"]);
	}
	return (null)
}
function ValidateObjectsFrames(testObject)
{
	return ((typeof(testObject.frames)=="object")&&(testObject.frames.length > 0));
}
function ValidateObject(testObject, tryLoadedFunction, validateDocument)
{
	var retVal = (typeof(testObject) == "object");
	if (typeof(tryLoadedFunction) == "undefined") tryLoadedFunction = false;
	if (retVal && tryLoadedFunction) retVal = ((typeof(testObject.IsLoaded) == "function")&&testObject.IsLoaded());
	if (typeof(validateDocument) == "undefined") validateDocument = false;
	if (retVal && validateDocument)	retval = (((typeof(testObject.document))!="undefined")&&(testObject.document != null) );
	return (retVal);
}
function EnableIcon(iconNumber, enableFlag)
{
	var NavIconbarObject = GetNavIconbarObject();
	if (NavIconbarObject)
	{
		if (NavIconbarObject.document.getElementById('icon0'))
		{
			NavIconbarObject.EnableIcon(iconNumber,enableFlag);
			return true;
		}
		else return false;
	}
	return false;
}
function GetNavIconbarObject()
{
	return (frames["fs_nav_bottom"]["NavIframeContainer"]["nav_icon_area"]);
}
function GetNavButtonObject()
{
	return (frames["fs_nav_bottom"]["NavIframeContainer"]["nav_button_area"]);
}
function ShrinkFrame()
{
	var NavTreeObject = GetNavTreeObject();
	if (NavTreeObject) NavTreeObject.ShrinkFrame();
	else PostFunctionCallback("ShrinkFrame();");
}
function IsBrowserIE()
{
	return (document.all ? true : false);
}

function CancelFunctionCallback(timerId)
{
	clearTimeout(timerId);
}
function CanNavigate()
{
	var EkMainObject = GetEkMainObject();
	return true;
}
function MakeNavTreeVisible(TreeName)
{
	var NavTreeObject = GetNavTreeObject();
	var NavFoldersObject = GetNavFoldersObject();
	if (m_MakeNavTreeVisibleTimerId)
	{
		CancelFunctionCallback(m_MakeNavTreeVisibleTimerId);
		m_MakeNavTreeVisibleTimerId = 0;
	}
	if (NavTreeObject && NavFoldersObject) NavTreeObject.MakeNavTreeVisible(TreeName, NavFoldersObject);
	else PostFunctionCallback("MakeNavTreeVisible('" + TreeName + "');");
}
function GetNavFoldersObject()
{
	return (frames["fs_nav_bottom"]["NavIframeContainer"]["nav_folder_area"]);
}

function ChangeOperationWindowLocation(ClassID,Action)
{
	var MainWindow=GetEkMainObject();
	MainWindow.location='Content.asp?Action='+Action+'&ClassID='+ClassID;
}
//帮助功能，不要删除
var ResizeObj=null;
var FSHelpFrameHeight=0;
var frameSizes=null;
var UpdateFrameTimeID=null;
var FSHelpFrameMaxHeight=0;
var IsBusy=false;
function GetFSHelpObject()
{
	return (frames["fs_nav_bottom"]["NavIframeContainer"]["FSHelp"]);
}
function ResizeFrame(Direction)
{
	ResizeObj=GetFSHelpObject().parent.document.getElementById("nav_divider");
	var bodySize=ResizeObj.rows
	frameSizes=bodySize.split(",");
	FSHelpFrameHeight=parseInt(frameSizes[3]);
	if (FSHelpFrameHeight>100) FSHelpFrameMaxHeight=FSHelpFrameHeight;
	if (Direction==null) return;
	if (!IsBusy) {UpdateFrameTimeID=setInterval("UpdateFrameSize("+Direction+");",1);IsBusy=true;}
}

function UpdateFrameSize(Direction)
{
	ResizeObj.rows=frameSizes[0]+","+frameSizes[1]+","+frameSizes[2]+","+FSHelpFrameHeight;
	if (Direction==1)
	{
		FSHelpFrameHeight=FSHelpFrameHeight-3;
		if (FSHelpFrameHeight<0) {clearInterval(UpdateFrameTimeID);IsBusy=false;}
	}
	else
	{
		FSHelpFrameHeight=FSHelpFrameHeight+3;
		if (FSHelpFrameHeight>FSHelpFrameMaxHeight) {clearInterval(UpdateFrameTimeID);IsBusy=false;}
	}
}
//帮助功能，不要删除
//拷贝、粘贴功能，不要删除
function MainInfoObj(SourceClass,SourceNews,SourceDownLoad,ObjectClass,MoveTF,OperationType)
{
	this.SourceClass=SourceClass;
	this.SourceNews=SourceNews;
	this.SourceDownLoad=SourceDownLoad;
	this.ObjectClass=ObjectClass;
	this.MoveTF=MoveTF;
	this.OperationType=OperationType;
}
var MainInfo=null;
var DocumentReadyTF=false;
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	MainInfo=new MainInfoObj('','','','',true,'');
	DocumentReadyTF=true;
}
var LoginStr='当前用户：<%=Session("Name")%>';
window.status="当前用户：<%=Session("Name")%> ";
//拷贝、粘贴功能，不要删除
</script>
<frameset rows="60,*" border="0" frameborder="no" framespacing="0">
	<frame name="fs_nav_top" src="top.asp" noresize scrolling="no" frameborder="no" marginheight="0" marginwidth="0">
	<frameset id="BottomFrameSet"  cols="192,*" border="2" frameborder="no" framespacing="0" bordercolor="#EEEEEE">
		<frame name="fs_nav_bottom" src="Menu_Container.asp" scrolling="no" marginwidth="0" marginheight="0" style="cursor:col-resize;border-left:0px;border-top:0px;border-bottom:0px;border-right:6px solid #E4E4E4" frameborder="no">
		<frame name="fs_main" id="fs_main" src="Sys_main.asp" scrolling="auto" marginwidth="0" marginheight="0" frameborder="no" style="cursor:col-resize;border-left:4px OUTset #ffffff;border-top:0px;border-bottom:0px;border-right:0px;">
 	</frameset>
</frameset><noframes></noframes>
<body>
</body>
</html>