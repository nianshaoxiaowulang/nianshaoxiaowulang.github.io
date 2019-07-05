<% Option Explicit %>
<!--#include file="../../Inc/Const.asp" -->
<!--#include file="../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../Inc/Session.asp" -->
<%
Dim FsoObj,OType
Set FsoObj = Server.CreateObject(G_FS_FSO)
OType = Request("Type")
Dim CurrPath,SubFolderObj,FolderObj,i,FsoItem
Dim ParentPath
if OType <> "" then
	Dim Path
	if OType = "Del" then
		Path = Request("Path") 
		if Path <> "" then
			Path = Server.MapPath(Path)
			if FsoObj.FolderExists(Path) = true then FsoObj.DeleteFolder Path
		end if
	elseif OType = "AddFolder" then
		Path = Request("Path")
		if Path <> "" then
			Path = Server.MapPath(Path)
			if FsoObj.FolderExists(Path) = True then
				Response.Write("<script>alert('目录已经存在');</script>")
			else
				FsoObj.CreateFolder Path
			end if
		end if
	elseif OType = "FolderReName" then
		Dim NewPathName,OldPathName,PhysicalPath,FileObj
		Path = Request("Path")
		if Path <> "" then
			NewPathName = Request("NewPathName")
			OldPathName = Request("OldPathName")
			if (NewPathName <> "") And (OldPathName <> "") then
				PhysicalPath = Server.MapPath(Path) & "\" & OldPathName
				if FsoObj.FolderExists(PhysicalPath) = True then
					PhysicalPath = Server.MapPath(Path) & "\" & NewPathName
					if FsoObj.FolderExists(PhysicalPath) = False then
						Set FileObj = FsoObj.GetFolder(Server.MapPath(Path) & "\" & OldPathName)
						FileObj.Name = NewPathName
						Set FileObj = Nothing
					end if
				end if
			end if
		end if
	end if
end if

CurrPath = Request("CurrPath")
if CurrPath = "" then
	CurrPath = "/"
	ParentPath = ""
else
	ParentPath = Mid(CurrPath,1,InstrRev(CurrPath,"/")-1)
	if ParentPath = "" then
		ParentPath = "/"
	end if
end if
On Error Resume Next
Set FolderObj = FsoObj.GetFolder(Server.MapPath(CurrPath))
Set SubFolderObj = FolderObj.SubFolders
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文件和目录列表</title>
</head>
<link href="../../CSS/FS_css.css" rel="stylesheet">
<script language="JavaScript" src="../SysJS/PublicJS.js"></script>
<script language="JavaScript">
var ObjPopupMenu=window.createPopup();
document.oncontextmenu=new Function("return ShowMouseRightMenu(window.event);");
var DocumentReadyTF=false;
var ContentMenuArray=new Array();
function document.onreadystatechange()
{
	if (DocumentReadyTF) return;
	InitialClassListContentMenu();
	DocumentReadyTF=true;
}
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
function ContentMenuFunction(ExeFunction,Description,EnabledStr)
{
	this.ExeFunction=ExeFunction;
	this.Description=Description;
	this.EnabledStr=EnabledStr;
}
function ContentMenuShowEvent()
{
	SelectFolder();
}
function InitialClassListContentMenu()
{
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.AddFolderOperation();",'新建','');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("if (confirm('确定要删除吗？')==true) parent.DelFolderFile();",'删除','disabled');
	ContentMenuArray[ContentMenuArray.length]=new ContentMenuFunction("parent.EditFolder();",'重命名','disabled');
}
</script>
<body topmargin="0" leftmargin="0" onClick="SelectFolder();">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="30%" class="ButtonListLeft"> <div align="center">名称</div></td>
    <td width="40%" class="ButtonList"> <div align="center">类型</div></td>
    <td class="ButtonList"> <div align="center">大小</div></td>
  </tr>
  <%
if Err.Number = 0 then
	for each FsoItem In SubFolderObj
%>
  <tr> 
    <td width="30%"><table border="0" cellspacing="0" cellpadding="0">
        <tr title="双击鼠标进入此目录"> 
          <td><img src="../Images/Folder/folderclosed.gif"></td>
          <td nowrap> <span class="TempletItem" Path="<% = FsoItem.name %>" onDblClick="OpenFolder(this);"> 
            <% = FsoItem.name %>
            </span> </td>
        </tr>
      </table></td>
    <td><div align="center">文件夹</div></td>
    <td><div align="center"> 
        <% = FsoItem.Size %>
      </div></td>
  </tr>
  <%
	next
else
%>
  <tr> 
    <td height="20" colspan="3">
      <div align="center"><% = "路径不存在" %></div></td>
  </tr>
  <%
end if
%>
</table>
</body>
</html>
<%
Set FsoObj = Nothing
Set SubFolderObj = Nothing
%>
<script language="JavaScript">
var CurrPath='<% = CurrPath %>';
var SelectedFolder='';
function SelectFolder()
{
	Obj=event.srcElement,DisabledContentMenuStr='';
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='TempletSelectItem') document.all(i).className='TempletItem';
	}
	if (Obj.Path!=null)
	{
		Obj.className='TempletSelectItem';
		SelectedFolder=Obj.Path;
	}
	else
	{
		SelectedFolder='';
	}
	if (SelectedFolder!='')
		DisabledContentMenuStr='';
	else
		DisabledContentMenuStr=',删除,重命名,';
	for (var i=0;i<ContentMenuArray.length;i++)
	{
		if (DisabledContentMenuStr.indexOf(ContentMenuArray[i].Description)!=-1) ContentMenuArray[i].EnabledStr='disabled';
		else  ContentMenuArray[i].EnabledStr='';
	}
}
function OpenParentFolder(Obj)
{
	location.href='SelectPath.asp?CurrPath='+Obj.Path;
	SearchOptionExists(parent.document.all.FolderSelectList,Obj.Path);
}

function OpenFolder(Obj)
{
	var SubmitPath='';
	if (CurrPath=='/') SubmitPath=CurrPath+Obj.Path;
	else SubmitPath=CurrPath+'/'+Obj.Path;
	location.href='SelectPath.asp?CurrPath='+SubmitPath;
	AddFolderList(parent.document.all.FolderSelectList,SubmitPath,SubmitPath);
}

function SelectUpFolder(Obj)
{
	for (var i=0;i<document.all.length;i++)
	{
		if (document.all(i).className=='TempletSelectItem') document.all(i).className='TempletItem';
	}
	Obj.className='TempletSelectItem';
}
function AddFolderList(SelectObj,Lable,LableContent)
{
	var i=0,AddOption;
	if (!SearchOptionExists(SelectObj,Lable))
	{
		AddOption = document.createElement("OPTION");
		AddOption.text=Lable;
		AddOption.value=LableContent;
		SelectObj.add(AddOption);
		SelectObj.options(SelectObj.length-1).selected=true;
	}
}
function DelFolderList(SelectObj,Lable,LableContent)
{
	var i,SelectIndex=-1;
	for(i=0;i<SelectObj.length;i++)
	{
		if (SelectObj.options(i).text==Lable) SelectIndex=i;
	}
	if (SelectIndex!=-1)
	{
		SelectObj.options.remove(SelectObj.selectedIndex);
	}
}
function SearchOptionExists(Obj,SearchText)
{
	var i;
	for(i=0;i<Obj.length;i++)
	{
		if (Obj.options(i).text==SearchText)
		{
			Obj.options(i).selected=true;
			return true;
		}
	}
	return false;
}
function AddFolderOperation()
{
	var ReturnValue=prompt('新建目录名：','');
	if ((ReturnValue!='') && (ReturnValue!=null))
		window.location.href='?Type=AddFolder&Path='+CurrPath+'/'+ReturnValue+'&CurrPath='+CurrPath;
}
function DelFolderFile()
{
	var ReturnValue='';
	if (SelectedFolder!='') 
		window.location.href='?Type=Del&Path='+CurrPath+'/'+SelectedFolder+'&CurrPath='+CurrPath;
	else alert('请选择要删除的目录');
}
function EditFolder()
{
	if (SelectedFolder!='')
	{
		var ReturnValue=prompt('修改的目录名：',SelectedFolder);
		if ((ReturnValue!='') && (ReturnValue!=null))
			window.location.href='?Type=FolderReName&Path='+CurrPath+'&CurrPath='+CurrPath+'&OldPathName='+SelectedFolder+'&NewPathName='+ReturnValue;
	}
	else alert('请填写要更名的目录名称');
}
</script>