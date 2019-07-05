<% Option Explicit %>
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
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
if Not JudgePopedomTF(Session("Name"),"P050300") then Call ReturnError1()
Dim UpFilesPath,FS,FolderObj,FileObj,FileItem,FolderItem,FileIconDic
UpFilesPath = Request("Path")
Set FS = Server.CreateObject(G_FS_FSO)
Set FolderObj = FS.GetFolder(Server.MapPath(UpFilesPath))
Set FileObj = FolderObj.Files

Set FileIconDic = CreateObject("Scripting.Dictionary")
FileIconDic.Add "txt","../../Images/FileIcon/txt.gif"
FileIconDic.Add "gif","../../Images/FileIcon/gif.gif"
FileIconDic.Add "exe","../../Images/FileIcon/exe.gif"
FileIconDic.Add "asp","../../Images/FileIcon/asp.gif"
FileIconDic.Add "html","../../Images/FileIcon/html.gif"
FileIconDic.Add "htm","../../Images/FileIcon/html.gif"
FileIconDic.Add "jpg","../../Images/FileIcon/jpg.gif"
FileIconDic.Add "jpeg","../../Images/FileIcon/jpg.gif"
FileIconDic.Add "pl","../../Images/FileIcon/perl.gif"
FileIconDic.Add "perl","../../Images/FileIcon/perl.gif"
FileIconDic.Add "zip","../../Images/FileIcon/zip.gif"
FileIconDic.Add "rar","../../Images/FileIcon/zip.gif"
FileIconDic.Add "gz","../../Images/FileIcon/zip.gif"
FileIconDic.Add "doc","../../Images/FileIcon/doc.gif"
FileIconDic.Add "xml","../../Images/FileIcon/xml.gif"
FileIconDic.Add "xsl","../../Images/FileIcon/xml.gif"
FileIconDic.Add "dtd","../../Images/FileIcon/xml.gif"
FileIconDic.Add "vbs","../../Images/FileIcon/vbs.gif"
FileIconDic.Add "js","../../Images/FileIcon/vbs.gif"
FileIconDic.Add "wsh","../../Images/FileIcon/vbs.gif"
FileIconDic.Add "sql","../../Images/FileIcon/script.gif"
FileIconDic.Add "bat","../../Images/FileIcon/script.gif"
FileIconDic.Add "tcl","../../Images/FileIcon/script.gif"
FileIconDic.Add "eml","../../Images/FileIcon/mail.gif"
FileIconDic.Add "swf","../../Images/FileIcon/flash.gif"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>上传文件列表</title>
</head>
<script src="../../SysJS/PublicJS.js" language="JavaScript"></script>
<script src="../../SysJS/ContentMenu.js" language="JavaScript"></script>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<body onselectstart="return false;" onClick="ClickFileName();" topmargin="2" leftmargin="2">
<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#EEEEEE"> 
    <td height="26" colspan="5" valign="middle">
      <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <tr>
		  <td width=35  align="center" alt="上传文件" onClick="UpLoad();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">导入</td>
		  <td width=2 class="Gray">|</td>
		  <td width=35 align="center" alt="删除" onClick="DelFolderFile();" onMouseMove="BtnMouseOver(this);" onMouseOut="BtnMouseOver(this);" class="BtnMouseOut">删除</td>
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
          <td height="26" class="ButtonListLeft"> 
            <div align="center">名称</div></td>
          <td height="26" class="ButtonList">
<div align="center">大小</div></td>
          <td height="26" class="ButtonList">
<div align="center">类型</div></td>
          <td height="26" class="ButtonList">
<div align="center">修改时间</div></td>
        </tr>
        <%
For Each FileItem in FileObj
	Dim FileIcon,FileExtName
	FileExtName = Mid(CStr(FileItem.Name),Instr(CStr(FileItem.Name),".")+1)
	FileIcon = FileIconDic.Item(LCase(FileExtName))
	if FileIcon = "" then
		FileIcon = "../../Images/FileIcon/unknown.gif"
	end if
%>
        <tr> 
          <td> <div align="left"> 
              <table border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td><img src="<% = FileIcon %>" width="16" height="16"></td>
                  <td><span class="TempletItem" FileName="<% = FileItem.Name %>">
                    <% = FileItem.Name %>
                    </span></td>
                </tr>
              </table>
            </div></td>
          <td><div align="center"> 
              <% = FileItem.Size %>
              字节 </div></td>
          <td><div align="center"> 
              <% = FileItem.Type %>
            </div></td>
          <td><div align="center"> 
              <% = FileItem.DateLastModified %>
            </div></td>
        </tr>
        <%
Next
%>
      </table></td>
  </tr>
</table>
</td>
</tr>
</table>
</body>
</html>
<%
Set FS = Nothing
Set FolderObj = Nothing
Set FileObj = Nothing
Set FileIconDic = Nothing
%>
<script language="JavaScript">
var UpFilePath='<% = UpFilesPath %>';
var DocumentReadyTF=false;
var ListObjArray = new Array();
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialContentListContentMenu();
	DocumentReadyTF=true;
}
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function InitialContentListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.UpLoad();",'导入','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('parent.DelFolderFile();','删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('seperator','','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('top.GetEkMainObject().location.reload();','刷新','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction('prompt(\'本页面路径属性\',\'<%=Request.ServerVariables("SCRIPT_NAME")%>\');','路径属性','');
	IntialListObjArray();
}
function ContentMenuShowEvent()
{
	ChangeSpecialMenuStatus();
}
function ChangeSpecialMenuStatus()
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
			if (SelectContent=='') SelectContent=ListObjArray[i].Obj.NewsID;
			else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.NewsID;
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
					if (SelectContent=='') SelectContent=ListObjArray[i].Obj.NewsID;
					else SelectContent=SelectContent+'***'+ListObjArray[i].Obj.NewsID;
				}
			}
		}
	}
	if (SelectContent=='') DisabledContentMenuStr=',删除,';
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
		if (CurrObj.FileName!=null)
		{
			ListObjArray[ListObjArray.length]=new FolderFileObj(CurrObj,j,false);
			j++;
		}
	}
}
function ClickFileName()
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
function UpLoad()
{
	OpenWindow('../../FunPages/Frame.asp?FileName=UpFileForm.asp&PageTitle=上传文件&Path='+UpFilePath,350,170,window);
}

function DelFolderFile()
{
	var SelectedFile='';
	for (i=0;i<ListObjArray.length;i++)
	{
		if (ListObjArray[i].Selected==true)
		{
			if (ListObjArray[i].Obj.FileName!=null)
			{
				if (SelectedFile=='') SelectedFile=ListObjArray[i].Obj.FileName;
				else  SelectedFile=SelectedFile+'***'+ListObjArray[i].Obj.FileName;
			}
		}
	}
	if (SelectedFile!='')
		OpenWindow('../../FunPages/Frame.asp?PageTitle=删除文件&Path='+UpFilePath+'&FileName=DelFolderAndFile.asp&DelFile='+SelectedFile,200,90,window);
	else alert('请选择文件');
}
</script>